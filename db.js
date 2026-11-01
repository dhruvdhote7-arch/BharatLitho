const path = require("path");
const fs = require("fs");
const EventEmitter = require("events");
const initSqlJs = require("sql.js");
const XLSX = require("xlsx");

const isDev = !__dirname.includes("app.asar");
const DB_DIR = isDev
    ? path.join(__dirname, "..", "database")
    : path.join(path.dirname(process.execPath), "database");
const DB_PATH = path.join(DB_DIR, "billing.db");
const EXCEL_PATH = path.join(DB_DIR, "BillingExcel.xlsx");

if (!fs.existsSync(DB_DIR)) fs.mkdirSync(DB_DIR, { recursive: true });

let db = null;

// -----------------------------------------------------
// Excel auto-watch state — see startExcelWatcher() below.
// We track the Excel file's own last-modified time rather than a
// timing guard: after exportToExcel() writes it, we immediately
// re-read that timestamp, so the watcher can tell "this is a change
// I already know about" (ours) from "this is new" (a real hand-edit)
// no matter how many intermediate file-system events fired in between.
// -----------------------------------------------------
const excelEvents = new EventEmitter();
let lastKnownExcelMtimeMs = 0;

function statMtimeMs(p) {
    try {
        return fs.statSync(p).mtimeMs;
    } catch {
        return -1; // missing right now — e.g. Excel mid-save
    }
}

const SCHEMA_SQL = `
    CREATE TABLE customers (
        customer_id INTEGER PRIMARY KEY AUTOINCREMENT,
        name        TEXT NOT NULL UNIQUE,
        address     TEXT,
        city        TEXT,
        state       TEXT,
        gstin       TEXT,
        phone       TEXT,
        pincode     TEXT
    );

    CREATE TABLE invoices (
        invoice_id    INTEGER PRIMARY KEY AUTOINCREMENT,
        invoice_no    TEXT NOT NULL UNIQUE,
        invoice_date  TEXT NOT NULL,

        customer_name TEXT NOT NULL,
        address       TEXT,
        city          TEXT,
        state         TEXT,
        gstin         TEXT,
        phone         TEXT,
        pincode       TEXT,
        material      TEXT,
        shipped_to    TEXT,
        challan_no    TEXT,
        challan_date  TEXT,
        receipt_no    TEXT,
        receipt_date  TEXT,
        gp_bundles    REAL DEFAULT 0,

        items_json    TEXT NOT NULL DEFAULT '[]',
        subtotal      REAL NOT NULL DEFAULT 0,
        cgst_percent  REAL NOT NULL DEFAULT 0,
        sgst_percent  REAL NOT NULL DEFAULT 0,
        igst_percent  REAL NOT NULL DEFAULT 0,
        gst_amount    REAL NOT NULL DEFAULT 0,
        grand_total   REAL NOT NULL DEFAULT 0,
        created_at    TEXT NOT NULL DEFAULT (datetime('now'))
    );
`;

// -----------------------------------------------------
// SQL helpers (sql.js API)
// -----------------------------------------------------

function q(sql, params = []) {
    const stmt = db.prepare(sql);
    stmt.bind(params);
    const rows = [];
    while (stmt.step()) rows.push(stmt.getAsObject());
    stmt.free();
    return rows;
}

function run(sql, params = []) {
    db.run(sql, params);
}

function one(sql, params = []) {
    const rows = q(sql, params);
    return rows.length ? rows[0] : null;
}

function persist() {
    const bytes = db.export();
    fs.writeFileSync(DB_PATH, Buffer.from(bytes));
}

// Adds a column to an existing table if it isn't there yet — lets
// older billing.db files (created before a field like pincode
// existed) pick up new columns without losing any data.
function ensureColumn(table, column, sqlType) {
    const cols = q(`PRAGMA table_info(${table})`);
    const exists = cols.some(c => c.name === column);
    if (!exists) {
        run(`ALTER TABLE ${table} ADD COLUMN ${column} ${sqlType}`);
    }
}

function runMigrations() {
    ensureColumn("customers", "pincode", "TEXT");
    ensureColumn("invoices", "pincode", "TEXT");
    ensureColumn("invoices", "shipped_to", "TEXT");
}

// -----------------------------------------------------
// Bootstrap — the only async part. Everything else below
// is synchronous once `ready` has resolved.
// -----------------------------------------------------

async function init() {
    // require.resolve walks up node_modules the same way require() does,
    // so this finds sql-wasm.wasm correctly whether db.js sits at the
    // project root or nested in src/ — no hardcoded relative depth.
    const wasmPath = require.resolve("sql.js/dist/sql-wasm.wasm");
    const SQL = await initSqlJs({
        locateFile: () => wasmPath
    });

    if (fs.existsSync(DB_PATH)) {
        db = new SQL.Database(fs.readFileSync(DB_PATH));
        runMigrations();
        persist();
    } else {
        db = new SQL.Database();
        db.run(SCHEMA_SQL);
        persist();
    }

    startExcelWatcher();
}

const ready = init();

// -----------------------------------------------------
// Customers
// -----------------------------------------------------

function getOrCreateCustomer(c) {
    const existing = one(`SELECT customer_id FROM customers WHERE name = ?`, [c.name]);

    if (existing) {
        run(`UPDATE customers SET address=?, city=?, state=?, gstin=?, phone=?, pincode=? WHERE customer_id=?`,
            [c.address, c.city, c.state, c.gst, c.phone, c.pincode, existing.customer_id]);
        return existing.customer_id;
    }

    run(`INSERT INTO customers (name, address, city, state, gstin, phone, pincode) VALUES (?, ?, ?, ?, ?, ?, ?)`,
        [c.name, c.address, c.city, c.state, c.gst, c.phone, c.pincode]);
    return one(`SELECT last_insert_rowid() AS id`).id;
}

function customerNames() {
    return q(`SELECT name FROM customers ORDER BY name`).map(r => r.name);
}

function findCustomerByName(name) {
    return one(`SELECT * FROM customers WHERE name = ?`, [name]);
}

// Full customer records (not just names) — used by the standalone
// Customers tab so it can show a table of everyone on file.
function listCustomers() {
    return q(`SELECT * FROM customers ORDER BY name`);
}

// Adds a brand-new customer, or updates an existing one if the name
// already matches (same upsert rule as getOrCreateCustomer, which
// runs automatically when an invoice is saved). Used by the Customers
// tab so a customer can be added independently of any invoice.
// Persists to disk and rewrites the Excel mirror immediately, same as
// saveInvoice/deleteInvoice do.
function saveCustomer(c) {
    if (!c || !c.name || !c.name.trim()) {
        throw new Error("Customer name is required.");
    }

    const id = getOrCreateCustomer({
        name: c.name.trim(),
        address: c.address || "",
        city: c.city || "",
        state: c.state || "",
        gst: c.gst || "",
        phone: c.phone || "",
        pincode: c.pincode || ""
    });

    persist();
    exportToExcel();
    return id;
}

// -----------------------------------------------------
// Invoices
// -----------------------------------------------------

function getNextInvoiceNo() {
    const row = one(`SELECT seq FROM sqlite_sequence WHERE name = 'invoices'`);
    const n = row ? row.seq + 1 : 1;

    const today = new Date();

    let startYear, endYear;

    // Financial year starts on April 1
    if (today.getMonth() >= 3) { // April = 3
        startYear = today.getFullYear();
        endYear = startYear + 1;
    } else {
        startYear = today.getFullYear() - 1;
        endYear = today.getFullYear();
    }

    const fy = `${String(startYear)}-${String(endYear).slice(-2)}`;

    return `BL/${String(n).padStart(4, "0")}/${fy}`;
}
function saveInvoice(data) {
    getOrCreateCustomer(data.customer);
    const itemsJson = JSON.stringify(data.items || []);

    if (data.invoiceId) {
        run(`
            UPDATE invoices SET
                invoice_no=?, invoice_date=?,
                customer_name=?, address=?, city=?, state=?,
                gstin=?, phone=?, pincode=?, material=?, shipped_to=?, challan_no=?, challan_date=?, receipt_no=?, receipt_date=?, gp_bundles=?,
                items_json=?,
                subtotal=?, cgst_percent=?, sgst_percent=?, igst_percent=?, gst_amount=?, grand_total=?
            WHERE invoice_id=?
        `, [
            data.invoiceNo,
            data.invoiceDate,
            data.customer.name,
            data.customer.address,
            data.customer.city,
            data.customer.state,
            data.customer.gst,
            data.customer.phone,
            data.customer.pincode,
            data.material,
            data.shippedTo,
            data.challan,
            data.challanDate,
            data.receiptNo,
            data.receiptDate,
            data.gpBundles,
            itemsJson,
            data.subtotal,
            data.cgstPercent,
            data.sgstPercent,
            data.igstPercent,
            data.gstAmount,
            data.grandTotal,
            data.invoiceId
        ]);
        persist();
        exportToExcel();
        return data.invoiceId;
    }

    run(`
        INSERT INTO invoices (
            invoice_no, invoice_date,
            customer_name, address, city, state, gstin, phone, pincode,
            material, shipped_to, challan_no, challan_date, receipt_no, receipt_date, gp_bundles,
            items_json, subtotal, cgst_percent, sgst_percent, igst_percent, gst_amount, grand_total
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `, [
        data.invoiceNo, data.invoiceDate,
        data.customer.name, data.customer.address,
        data.customer.city, data.customer.state, data.customer.gst, data.customer.phone, data.customer.pincode,
        data.material, data.shippedTo, data.challan, data.challanDate, data.receiptNo, data.receiptDate, data.gpBundles,
        itemsJson,
        data.subtotal, data.cgstPercent, data.sgstPercent, data.igstPercent, data.gstAmount, data.grandTotal
    ]);

    const newId = one(`SELECT last_insert_rowid() AS id`).id;
    persist();
    exportToExcel();
    return newId;
}
// const challanNo = invoiceNo.match(/^R\/(.*?)\//)[1];

function listInvoices(term) {
    term = (term || "").trim();
    if (term) {
        const like = `%${term}%`;
        return q(`
            SELECT invoice_id, invoice_no, invoice_date, customer_name, grand_total
            FROM invoices WHERE invoice_no LIKE ? OR customer_name LIKE ?
            ORDER BY invoice_id DESC
        `, [like, like]);
    }
    return q(`
        SELECT invoice_id, invoice_no, invoice_date, customer_name, grand_total
        FROM invoices ORDER BY invoice_id DESC
    `);
}

function getInvoice(invoiceId) {
    const header = one(`SELECT * FROM invoices WHERE invoice_id=?`, [invoiceId]);
    if (!header) return null;
    return { header, items: JSON.parse(header.items_json || "[]") };
}

// Looks up an existing invoice's id by its invoice number — used to
// offer "replace the old bill" when saving hits the invoice_no
// UNIQUE constraint.
function findInvoiceIdByNo(invoiceNo) {
    const row = one(`SELECT invoice_id FROM invoices WHERE invoice_no = ?`, [invoiceNo]);
    return row ? row.invoice_id : null;
}

function deleteInvoice(invoiceId) {
    run(`DELETE FROM invoices WHERE invoice_id=?`, [invoiceId]);
    persist();
    exportToExcel();
}

function invoiceCount() {
    return one(`SELECT COUNT(*) AS n FROM invoices`).n;
}

// Distinct particulars this customer has been billed for before,
// most recent invoice first — used to populate the Particulars dropdown.
function getCustomerParticulars(customerName) {
    if (!customerName) return [];

    const rows = q(`
        SELECT items_json FROM invoices WHERE customer_name = ? ORDER BY invoice_id DESC LIMIT 50
    `, [customerName]);

    const seen = new Set();
    const list = [];
    rows.forEach(r => {
        JSON.parse(r.items_json || "[]").forEach(it => {
            const p = (it.particular || "").trim();
            if (p && !seen.has(p)) {
                seen.add(p);
                list.push(p);
            }
        });
    });
    return list;
}

// Distinct Shipped To addresses used on past invoices, most recent
// first — used to populate the Shipped To suggestion dropdown.
function getShippedToAddresses() {
    const rows = q(`
        SELECT shipped_to FROM invoices
        WHERE shipped_to IS NOT NULL AND TRIM(shipped_to) <> ''
        ORDER BY invoice_id DESC LIMIT 200
    `);

    const seen = new Set();
    const list = [];
    rows.forEach(r => {
        const s = (r.shipped_to || "").trim();
        if (s && !seen.has(s)) {
            seen.add(s);
            list.push(s);
        }
    });
    return list;
}

// -----------------------------------------------------
// Excel mirror — rewritten from the database on every
// save/delete, so it's always current.
// -----------------------------------------------------

function exportToExcel() {
    const invoices = q(`SELECT * FROM invoices ORDER BY invoice_id`);
    const customers = q(`SELECT * FROM customers ORDER BY name`);

    const bills = invoices.map(inv => {
        const items = JSON.parse(inv.items_json || "[]");
        const itemsSummary = items
            .map(it => `${it.particular} (qty ${it.qty})`)
            .join("; ");

        return {
            InvoiceNo: inv.invoice_no,
            Date: inv.invoice_date,
            CustomerName: inv.customer_name,
            Address: inv.address,
            City: inv.city,
            State: inv.state,
            GSTIN: inv.gstin,
            Phone: inv.phone,
            Pincode: inv.pincode,
            Material: inv.material,
            ShippedTo: inv.shipped_to,
            ChallanNo: inv.challan_no,
            ChallanDate: inv.challan_date,
            ReceiptNo: inv.receipt_no,
            ReceiptDate: inv.receipt_date,
            GPBundles: inv.gp_bundles,
            Items: itemsSummary,
            Subtotal: inv.subtotal,
            CGSTPercent: inv.cgst_percent,
            SGSTPercent: inv.sgst_percent,
            IGSTPercent: inv.igst_percent,
            GSTAmount: inv.gst_amount,
            GrandTotal: inv.grand_total
        };
    });

    const customerRows = customers.map(c => ({
        CustomerName: c.name,
        Address: c.address,
        City: c.city,
        State: c.state,
        GSTIN: c.gstin,
        Phone: c.phone,
        Pincode: c.pincode
    }));

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(bills), "Bills");
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(customerRows), "Customers");

    XLSX.writeFile(workbook, EXCEL_PATH);
    // Record the timestamp of the write we just made, so the watcher
    // recognizes it as "already known" rather than importing our own
    // export back in a pointless loop.
    lastKnownExcelMtimeMs = statMtimeMs(EXCEL_PATH);
}

// -----------------------------------------------------
// Excel import — reads BillingExcel.xlsx back in and
// applies edits. Deliberately limited in scope:
//
//   Customers sheet -> full upsert (adds new customers,
//                       updates existing ones by name — same
//                       rule saveCustomer already uses).
//
//   Bills sheet     -> HEADER FIELDS ONLY, and only on invoices
//                       that already exist (matched by InvoiceNo).
//                       Item lines, Subtotal, GSTAmount, and
//                       GrandTotal are never read from Excel:
//                       Subtotal stays whatever the database's
//                       item lines already total to, and
//                       GSTAmount/GrandTotal are recalculated
//                       from that plus the (possibly edited)
//                       GST% columns. The customer-snapshot
//                       columns on a bill (Address/City/etc.)
//                       are also left alone — edit those via the
//                       Customers sheet instead. Rows whose
//                       InvoiceNo doesn't match an existing
//                       invoice are skipped, not created — a new
//                       invoice needs real item lines, which this
//                       sheet doesn't carry in a safe, re-editable
//                       form (Items is just a flattened summary).
// -----------------------------------------------------

const EDITABLE_INVOICE_FIELDS = [
    // [Excel column name, db column, value kind]
    ["Date", "invoice_date", "date"],
    ["Material", "material", "text"],
    ["ShippedTo", "shipped_to", "text"],
    ["ChallanNo", "challan_no", "text"],
    ["ChallanDate", "challan_date", "date"],
    ["ReceiptNo", "receipt_no", "text"],
    ["ReceiptDate", "receipt_date", "date"],
    ["GPBundles", "gp_bundles", "number"],
    ["CGSTPercent", "cgst_percent", "number"],
    ["SGSTPercent", "sgst_percent", "number"],
    ["IGSTPercent", "igst_percent", "number"]
];

function dateToIso(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
}

// Accepts whatever a date cell might come back as after a round trip
// through Excel: the plain "YYYY-MM-DD" text this app originally wrote,
// a "DD-MM-YYYY"/"DD/MM/YYYY" string (if someone typed it by hand), a
// real JS Date (if cell-dates parsing kicked in), or an Excel date
// serial number (if the cell got reformatted as a Date in Excel).
// Returns "" for a blank cell (clears the field), or null if the value
// couldn't be understood at all (caller leaves the field untouched).
function normalizeExcelDate(value) {
    if (value === undefined || value === null || value === "") return "";

    if (value instanceof Date) return dateToIso(value);

    if (typeof value === "number") {
        const parsed = XLSX.SSF.parse_date_code(value);
        if (!parsed) return null;
        return dateToIso(new Date(parsed.y, parsed.m - 1, parsed.d));
    }

    const str = String(value).trim();
    if (!str) return "";
    if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;

    const dm = str.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
    if (dm) {
        const [, d, m, y] = dm;
        return `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
    }

    return null;
}

function importFromExcel(filePath) {
    filePath = filePath || EXCEL_PATH;
    if (!fs.existsSync(filePath)) {
        throw new Error(`Excel file not found: ${filePath}`);
    }

    const workbook = XLSX.readFile(filePath);
    const result = {
        customersAdded: 0,
        customersUpdated: 0,
        customersSkipped: [],
        invoicesUpdated: 0,
        invoicesSkipped: [],
        warnings: []
    };

    db.run("BEGIN TRANSACTION");
    try {
        // ---- Customers sheet: full upsert ----
        const custSheet = workbook.Sheets["Customers"];
        if (custSheet) {
            const rows = XLSX.utils.sheet_to_json(custSheet, { defval: "" });
            rows.forEach((row, i) => {
                const name = String(row.CustomerName || "").trim();
                if (!name) {
                    result.customersSkipped.push(`Customers row ${i + 2}: no CustomerName`);
                    return;
                }
                const existed = !!findCustomerByName(name);
                getOrCreateCustomer({
                    name,
                    address: String(row.Address || "").trim(),
                    city: String(row.City || "").trim(),
                    state: String(row.State || "").trim(),
                    gst: String(row.GSTIN || "").trim(),
                    phone: String(row.Phone || "").trim(),
                    pincode: String(row.Pincode || "").trim()
                });
                if (existed) result.customersUpdated++;
                else result.customersAdded++;
            });
        }

        // ---- Bills sheet: header fields only, existing invoices only ----
        const billsSheet = workbook.Sheets["Bills"];
        if (billsSheet) {
            const rows = XLSX.utils.sheet_to_json(billsSheet, { defval: "" });
            rows.forEach(row => {
                const invoiceNoVal = String(row.InvoiceNo || "").trim();
                if (!invoiceNoVal) return;

                const existing = one(`SELECT * FROM invoices WHERE invoice_no = ?`, [invoiceNoVal]);
                if (!existing) {
                    result.invoicesSkipped.push(invoiceNoVal);
                    return;
                }

                const updates = {};
                EDITABLE_INVOICE_FIELDS.forEach(([excelCol, dbCol, kind]) => {
                    if (!(excelCol in row)) return;
                    const raw = row[excelCol];

                    if (kind === "date") {
                        const iso = normalizeExcelDate(raw);
                        if (iso === null) {
                            result.warnings.push(
                                `${invoiceNoVal}: couldn't read "${excelCol}" value "${raw}" — left unchanged`
                            );
                            return;
                        }
                        updates[dbCol] = iso;
                    } else if (kind === "number") {
                        updates[dbCol] = parseFloat(raw) || 0;
                    } else {
                        updates[dbCol] = String(raw ?? "").trim();
                    }
                });

                if (Object.keys(updates).length === 0) return;

                // Subtotal is locked — it belongs to the item lines already
                // in the database, not to anything on this sheet.
                const subtotal = existing.subtotal;
                const cgst = "cgst_percent" in updates ? updates.cgst_percent : existing.cgst_percent;
                const sgst = "sgst_percent" in updates ? updates.sgst_percent : existing.sgst_percent;
                const igst = "igst_percent" in updates ? updates.igst_percent : existing.igst_percent;
                const gstAmount = igst > 0
                    ? subtotal * (igst / 100)
                    : subtotal * ((cgst + sgst) / 100);

                updates.gst_amount = gstAmount;
                updates.grand_total = subtotal + gstAmount;

                const setCols = Object.keys(updates);
                const setSql = setCols.map(c => `${c}=?`).join(", ");
                run(`UPDATE invoices SET ${setSql} WHERE invoice_id=?`,
                    [...setCols.map(c => updates[c]), existing.invoice_id]);

                result.invoicesUpdated++;
            });
        }

        db.run("COMMIT");
    } catch (err) {
        db.run("ROLLBACK");
        throw err;
    }

    persist();
    exportToExcel(); // refresh the mirror so it shows recalculated totals
    return result;
}

// -----------------------------------------------------
// Excel auto-watch — no button needed. Watches the database folder
// for changes to BillingExcel.xlsx and runs importFromExcel()
// automatically a moment after each save.
//
// A few things this deliberately accounts for:
//  - Excel doesn't do a single simple write when you hit Ctrl+S — on
//    Windows it typically writes a temp file (often under a different,
//    unpredictable name), deletes the original, then renames the temp
//    file into place. Trying to filter file-system events by filename
//    is fragile against that, so this doesn't try — ANY change inside
//    the database folder just triggers a check of EXCEL_PATH's own
//    last-modified time instead. Whatever name Excel used along the
//    way, the file that ends up sitting at EXCEL_PATH is what counts.
//  - fs.watch is debounced (waits ~1s after the last event) so that
//    burst of several events from one save collapses into one import,
//    and a slow poll (every 4s) runs alongside it as a backstop, in
//    case fs.watch itself doesn't fire reliably on a given machine
//    (this happens on some Windows setups with certain antivirus
//    tools or synced folders like OneDrive).
//  - The file can still be locked for a moment right after those
//    events fire (Excel finishing up, antivirus scanning it, or the
//    file briefly not existing mid delete-then-rename). Both the
//    mtime check and the import itself retry a few times with a short
//    backoff before giving up.
//  - exportToExcel() also writes this same file (every invoice/
//    customer save rewrites the mirror) and records the resulting
//    mtime immediately after — so the watcher recognizes that as
//    "already known" and skips it, instead of pointlessly re-importing
//    data that just came from the database in the first place.
// -----------------------------------------------------

let watchDebounceTimer = null;
let watcherStarted = false;

function startExcelWatcher() {
    if (watcherStarted) return;
    watcherStarted = true;

    lastKnownExcelMtimeMs = statMtimeMs(EXCEL_PATH);
    console.log("Watching for Excel edits at:", EXCEL_PATH);

    try {
        fs.watch(DB_DIR, () => {
            clearTimeout(watchDebounceTimer);
            watchDebounceTimer = setTimeout(() => checkExcelChanged(), 1000);
        });
    } catch (err) {
        excelEvents.emit("error", err);
    }

    // Backstop: some setups don't deliver fs.watch events reliably at
    // all. This costs one cheap stat() call every few seconds and
    // guarantees an edit is noticed even then.
    setInterval(() => checkExcelChanged(), 4000);
}

function isLikelyLockError(err) {
    const msg = String((err && err.message) || err);
    return /EBUSY|EPERM|locked|being used by another process/i.test(msg);
}

function checkExcelChanged(attempt = 0) {
    const mtime = statMtimeMs(EXCEL_PATH);

    if (mtime === -1) {
        // File briefly missing mid-save (delete-then-rename) — wait and retry.
        if (attempt < 5) setTimeout(() => checkExcelChanged(attempt + 1), 400);
        return;
    }
    if (mtime <= lastKnownExcelMtimeMs) return; // nothing new since we last looked

    // Record now, before attempting the import — so if the import
    // itself fails for a reason other than a lock, we don't retry the
    // same mtime forever on every subsequent poll tick.
    lastKnownExcelMtimeMs = mtime;
    runWatchedImport();
}

function runWatchedImport(attempt = 0) {
    try {
        const result = importFromExcel();
        excelEvents.emit("imported", result);
    } catch (err) {
        if (attempt < 5 && isLikelyLockError(err)) {
            setTimeout(() => runWatchedImport(attempt + 1), 500);
        } else {
            excelEvents.emit("error", err);
        }
    }
}

module.exports = {
    ready: () => ready,
    getNextInvoiceNo,
    saveInvoice,
    listInvoices,
    getInvoice,
    findInvoiceIdByNo,
    deleteInvoice,
    customerNames,
    findCustomerByName,
    listCustomers,
    saveCustomer,
    getCustomerParticulars,
    getShippedToAddresses,
    invoiceCount,
    exportToExcel,
    importFromExcel,
    onExcelImported: (cb) => excelEvents.on("imported", cb),
    onExcelImportError: (cb) => excelEvents.on("error", cb),
    EXCEL_PATH,
    DB_PATH
};