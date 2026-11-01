// =========================================================
// Bharat Litho Billing — app.js (simple offline edition)
// =========================================================

let db;
try {
    db = require("./db.js");
} catch (err) {
    console.error(err);
    document.body.innerHTML = `
        <div style="max-width:640px;margin:60px auto;padding:24px;
            font-family:Arial, sans-serif;border:2px solid #a3392f;
            border-radius:6px;background:#fff3f2;color:#1c2321;">
            <h2 style="margin-top:0;color:#a3392f;">The database couldn't load</h2>
            <p>The app's script stopped here, which is why nothing below the
               customer fields responds to typing or clicks.</p>
            <p>This almost always means the app's dependencies
               (<code>sql.js</code> and <code>xlsx</code>) aren't installed
               yet. From a terminal, in the app folder, run:</p>
            <pre style="background:#fff;border:1px solid #dcd8ce;padding:10px;
                border-radius:4px;overflow-x:auto;">npm install</pre>
            <p>then restart the app. If it still fails, open DevTools
               (Ctrl+Shift+I) and check the Console tab — the exact error
               below is:</p>
            <pre style="background:#fff;border:1px solid #dcd8ce;padding:10px;
                border-radius:4px;overflow-x:auto;white-space:pre-wrap;">${(err && err.message) || err}</pre>
        </div>
    `;
    throw err; // stop here — nothing below can safely run without db
}
const path = require("path");

const tableBody = document.getElementById("tableBody");
const subtotalBox = document.getElementById("pisubtotal");
const cgstPercentBox = document.getElementById("picgstPercent");
const sgstPercentBox = document.getElementById("pisgstPercent");
const igstPercentBox = document.getElementById("piigstPercent");
const gstAmountBox = document.getElementById("gstAmount");
const grandTotalBox = document.getElementById("grandTotal");
const invoiceDate = document.getElementById("invoiceDate");
const invoiceNo = document.getElementById("invoiceNo");
const customerNameInput = document.getElementById("customerName");
const challanDate = document.getElementById("challanDate");
const receiptDate = document.getElementById("receiptDate");
const challanBox = document.getElementById("challan");
const receiptNoBox = document.getElementById("receiptNo");
const shippedToBox = document.getElementById("shippedTo");

let currentInvoiceId = null;

// ---------- SHARED DIALOG (replaces window.alert / window.confirm) ----------
// Electron's native confirm()/alert() dialogs can leave the BrowserWindow
// unable to receive keyboard input once closed (a known Electron/Chromium
// focus quirk, most visible on Windows) — the window looks focused, but
// nothing you type registers until you physically click into a field
// again. That's the "click New, then can't type" bug. This in-page modal
// never touches OS-level focus, so it doesn't have that problem.
const dialogBackdrop = document.getElementById("dialogBackdrop");
const dialogTitle = document.getElementById("dialogTitle");
const dialogMessage = document.getElementById("dialogMessage");
const dialogOkBtn = document.getElementById("dialogOkBtn");
const dialogCancelBtn = document.getElementById("dialogCancelBtn");

function showDialog(message, { title = "Notice", showCancel = false } = {}) {
    return new Promise((resolve) => {
        dialogTitle.textContent = title;
        dialogMessage.textContent = message;
        dialogCancelBtn.style.display = showCancel ? "inline-block" : "none";
        dialogBackdrop.classList.add("active");

        function cleanup(result) {
            dialogBackdrop.classList.remove("active");
            dialogOkBtn.removeEventListener("click", onOk);
            dialogCancelBtn.removeEventListener("click", onCancel);
            resolve(result);
        }
        function onOk() { cleanup(true); }
        function onCancel() { cleanup(false); }

        dialogOkBtn.addEventListener("click", onOk);
        dialogCancelBtn.addEventListener("click", onCancel);
        dialogOkBtn.focus();
    });
}

// Drop-in replacements: appAlert("msg") / await appConfirm("msg")
function appAlert(message) {
    return showDialog(message, { showCancel: false });
}
function appConfirm(message) {
    return showDialog(message, { showCancel: true });
}
window.appAlert = appAlert;
window.appConfirm = appConfirm;

// ---------- START ----------

window.onload = async () => {
    try {
        setDbStatus("loading", "Loading database…");
        await db.ready();
        setDbStatus("ready", "Database ready — saved locally");
        loadDate();
        invoiceNo.value = db.getNextInvoiceNo();
        mirrorToChallan(invoiceNo.value);
        addRow();
        addRow();
        refreshCustomerSuggestions();
        refreshShippedToSuggestions();
        setExcelStatus();
    } catch (err) {
        console.error(err);
        setDbStatus("error", "Database failed to load: " + err.message);
    }
};

function setDbStatus(state, text) {
    document.getElementById("dbDot").className = "dot " + state;
    document.getElementById("dbStatus").textContent = text;
}

function setExcelStatus() {
    document.getElementById("excelStatus").textContent =
        "Excel mirror: " + db.EXCEL_PATH;
}

function loadDate() {
    const today = new Date();
    const y = today.getFullYear();
    const m = String(today.getMonth() + 1).padStart(2, "0");
    const d = String(today.getDate()).padStart(2, "0");
    invoiceDate.value = `${y}-${m}-${d}`;
    challanDate.value = `${y}-${m}-${d}`;
    receiptDate.value = `${y}-${m}-${d}`;
}

// ---------- CUSTOMER AUTOFILL ----------
// The invoice form only shows the customer's Name now — no address,
// city, GSTIN, etc. Those live on the customer record (added/edited
// via the Customers button) and are pulled in from the database at
// save/print time instead of being typed or shown here.

function refreshCustomerSuggestions() {
    const list = document.getElementById("customerNameList");
    list.innerHTML = db.customerNames()
        .map(name => `<option value="${escapeHtml(name)}">`)
        .join("");
}

// ---------- CHALLAN NO AUTO-COPY ----------
// Invoice No looks like "BL/0005/25-26" — as soon as the user types
// into Invoice No or Receipt No, just the number-between-the-slashes
// part ("0005") is mirrored live into Challan No, same as it's
// auto-generated for Invoice No itself. If the value doesn't have
// that slash pattern (e.g. Receipt No, which is free text), the raw
// value is used as-is. If the user then edits Challan No directly,
// mirroring stops for the rest of this invoice, so a deliberate
// change never gets silently overwritten. Resets on New/Save (see
// startNewInvoice()).
let challanAutoFilled = true;

challanBox.addEventListener("input", () => {
    challanAutoFilled = false;
});

function extractChallanNumber(value) {
    const match = String(value || "").match(/^[^/]*\/([^/]*)\//);
    return match ? match[1] : value;
}

function mirrorToChallan(value) {
    if (!challanAutoFilled) return;
    challanBox.value = extractChallanNumber(value);
}

invoiceNo.addEventListener("input", () => mirrorToChallan(invoiceNo.value));
receiptNoBox.addEventListener("input", () => mirrorToChallan(receiptNoBox.value));

// ---------- CHALLAN DATE AUTO-COPY ----------
// Same pattern as Challan No above: Challan Date mirrors Invoice Date
// live as soon as the user changes it, until the user edits Challan
// Date directly — then mirroring stops for the rest of this invoice,
// so a deliberate change never gets silently overwritten. Resets on
// New/Save (see startNewInvoice()).
let challanDateAutoFilled = true;

challanDate.addEventListener("input", () => {
    challanDateAutoFilled = false;
});

invoiceDate.addEventListener("input", () => {
    if (!challanDateAutoFilled) return;
    challanDate.value = invoiceDate.value;
});

customerNameInput.addEventListener("change", () => {
    refreshParticularsList(customerNameInput.value.trim());
});

customerNameInput.addEventListener("blur", () => {
    refreshParticularsList(customerNameInput.value.trim());
});

let currentParticularSuggestions = [];

function refreshParticularsList(customerName) {
    currentParticularSuggestions = db.getCustomerParticulars(customerName);
}

// Shows a small clickable list of this customer's past particulars
// under the given textarea (textareas don't support <datalist>).
function showParticularSuggestions(textarea) {
    hideParticularSuggestions();
    if (currentParticularSuggestions.length === 0) return;

    const box = document.createElement("div");
    box.className = "suggest-box";
    box.id = "activeSuggestBox";
    currentParticularSuggestions.forEach(p => {
        const opt = document.createElement("div");
        opt.className = "suggest-option";
        opt.textContent = p;
        opt.addEventListener("mousedown", (e) => {
            e.preventDefault();
            textarea.value = p;
            hideParticularSuggestions();
            calculateTotals();
        });
        box.appendChild(opt);
    });

    textarea.parentElement.appendChild(box);
}

function hideParticularSuggestions() {
    const existing = document.getElementById("activeSuggestBox");
    if (existing) existing.remove();
}

// ---------- SHIPPED TO SUGGESTIONS ----------
// Same idea as the Particulars suggestions above: previously used
// Shipped To addresses (across all invoices, not just this customer),
// most recent first, in a clickable dropdown under the textarea.

let currentShippedToSuggestions = [];

function refreshShippedToSuggestions() {
    currentShippedToSuggestions = db.getShippedToAddresses();
}

function showShippedToSuggestions(textarea) {
    hideShippedToSuggestions();
    if (currentShippedToSuggestions.length === 0) return;

    const box = document.createElement("div");
    box.className = "suggest-box";
    box.id = "activeShippedToSuggestBox";
    currentShippedToSuggestions.forEach(addr => {
        const opt = document.createElement("div");
        opt.className = "suggest-option";
        opt.textContent = addr;
        opt.addEventListener("mousedown", (e) => {
            e.preventDefault();
            textarea.value = addr;
            hideShippedToSuggestions();
        });
        box.appendChild(opt);
    });

    textarea.parentElement.appendChild(box);
}

function hideShippedToSuggestions() {
    const existing = document.getElementById("activeShippedToSuggestBox");
    if (existing) existing.remove();
}

shippedToBox.addEventListener("focus", () => showShippedToSuggestions(shippedToBox));
shippedToBox.addEventListener("blur", () => setTimeout(hideShippedToSuggestions, 150));

function escapeHtml(str) {
    return String(str)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
}

// ---------- CUSTOMERS TAB ----------
// A standalone modal for adding or editing customers directly,
// independent of any invoice. Saving here writes straight into the
// customers table, rewrites BillingExcel.xlsx, and refreshes the
// Customer Name dropdown on the invoice form.

document.getElementById("customersBtn").addEventListener("click", () => {
    document.getElementById("customerBackdrop").classList.add("active");
    document.getElementById("custName").focus();
});

document.getElementById("customerCloseBtn").addEventListener("click", () => {
    document.getElementById("customerBackdrop").classList.remove("active");
});

document.getElementById("custClearBtn").addEventListener("click", () => {
    document.getElementById("customerForm").reset();
    document.getElementById("custName").focus();
});

document.getElementById("customerForm").addEventListener("submit", async (e) => {
    e.preventDefault();

    const c = {
        name: document.getElementById("custName").value.trim(),
        address: document.getElementById("custAddress").value.trim(),
        city: document.getElementById("custCity").value.trim(),
        state: document.getElementById("custState").value.trim(),
        gst: document.getElementById("custGst").value.trim(),
        phone: document.getElementById("custPhone").value.trim(),
        pincode: document.getElementById("custPincode").value.trim()
    };

    if (!c.name) {
        await appAlert("Customer name is required.");
        document.getElementById("custName").focus();
        return;
    }

    try {
        db.saveCustomer(c);
        refreshCustomerSuggestions();
        document.getElementById("customerForm").reset();
        setDbStatus("ready", `Saved customer "${c.name}" — Excel updated`);
        document.getElementById("custName").focus();
    } catch (err) {
        console.error(err);
        await appAlert("Could not save customer: " + err.message);
    }
});

// ---------- EXCEL AUTO-IMPORT ----------
// No button — db.js watches BillingExcel.xlsx on disk and imports it
// automatically a moment after every save. Customers are fully
// editable there. On the Bills sheet, only header fields (dates,
// times, transport, material, challan/receipt info, GP bundles, GST%)
// are picked up, and only for invoices that already exist (matched by
// Invoice No) — see db.js importFromExcel() for exactly why Item
// lines/Subtotal/Grand Total are never read back from that sheet.

db.onExcelImported((result) => {
    const changed = result.customersAdded + result.customersUpdated + result.invoicesUpdated;
    const issues = result.invoicesSkipped.length + result.customersSkipped.length + result.warnings.length;

    if (changed === 0 && issues === 0) return; // nothing to say — likely our own export

    setDbStatus("ready", summarizeAutoImport(result, changed, issues));
    if (changed > 0) refreshCustomerSuggestions();

    if (issues > 0) {
        console.warn("Excel auto-import had issues:", result);
    }
});

db.onExcelImportError((err) => {
    console.error("Excel auto-import failed:", err);
    setDbStatus("error", "Couldn't read BillingExcel.xlsx: " + err.message);
});

function summarizeAutoImport(r, changed, issues) {
    const parts = [];
    if (r.customersAdded) parts.push(`${r.customersAdded} customer${r.customersAdded === 1 ? "" : "s"} added`);
    if (r.customersUpdated) parts.push(`${r.customersUpdated} customer${r.customersUpdated === 1 ? "" : "s"} updated`);
    if (r.invoicesUpdated) parts.push(`${r.invoicesUpdated} invoice${r.invoicesUpdated === 1 ? "" : "s"} updated`);

    let text = changed > 0
        ? "Excel import: " + parts.join(", ")
        : "Excel saved, but nothing matched";

    if (issues > 0) {
        text += ` — ${issues} row${issues === 1 ? "" : "s"} skipped, see console (Ctrl+Shift+I)`;
    }
    return text;
}

// ---------- MATERIAL SUPPLY ("Other" option) ----------
// The dropdown covers the common cases; picking "Other" reveals a
// plain text box so a one-off material description can still be
// typed in and printed, without needing to edit the dropdown itself.

const materialSelect = document.getElementById("material");
const materialOtherField = document.getElementById("materialOtherField");
const materialOtherInput = document.getElementById("materialOther");

materialSelect.addEventListener("change", () => {
    const isOther = materialSelect.value === "__other__";
    materialOtherField.style.display = isOther ? "block" : "none";
    if (isOther) materialOtherInput.focus();
    else materialOtherInput.value = "";
});

// Resolves the select + free-text pair down to the single string that
// gets saved/printed as the material.
function getMaterialValue() {
    if (materialSelect.value === "__other__") {
        return materialOtherInput.value.trim();
    }
    return materialSelect.value;
}

// ---------- ITEM ROWS ----------

document.getElementById("addRowBtn").addEventListener("click", () => addRow());

function addRow() {
    const row = document.createElement("tr");
    row.innerHTML = `
        <td><input class="qty" type="number" min="0"></td>
        <td class="particular-cell"><textarea class="particular" rows="2"></textarea></td>
        <td><input class="amount amount-field" type="text"></td>
        <td><button type="button" class="remove-row" title="Remove row">&times;</button></td>
    `;
    tableBody.appendChild(row);

    const qty = row.querySelector(".qty");
    const amount = row.querySelector(".amount");
    const particular = row.querySelector(".particular");

    // Amount is typed in by hand (the real calculation for this business
    // — paper sizes, ups, wastage — doesn't reduce to a simple formula),
    // so Qty is a reference field only and never overwrites it.
    amount.addEventListener("input", calculateTotals);
    particular.addEventListener("input", () => {
    particular.style.height = "auto";
    particular.style.height = particular.scrollHeight + "px";
});

    particular.addEventListener("focus", () => showParticularSuggestions(particular));
    particular.addEventListener("blur", () => setTimeout(hideParticularSuggestions, 150));

    row.querySelector(".remove-row").addEventListener("click", () => {
        row.remove();
        calculateTotals();
    });

    return row;
}

function calculateTotals() {
    let subtotal = 0;
    tableBody.querySelectorAll("tr").forEach(row => {
        subtotal += parseFloat(row.querySelector(".amount").value) || 0;
    });
    subtotalBox.value = subtotal.toFixed(2);
    calculateGST();
}

function calculateGST() {
    const subtotal = parseFloat(subtotalBox.value) || 0;
    const sgst = parseFloat(sgstPercentBox.value) || 0;
    const cgst = parseFloat(cgstPercentBox.value) || 0;
    const igst = parseFloat(igstPercentBox.value) || 0;

    const gstAmount = igst > 0
        ? subtotal * (igst / 100)
        : subtotal * ((cgst + sgst) / 100);

    gstAmountBox.value = gstAmount.toFixed(2);
    grandTotalBox.value = (subtotal + gstAmount).toFixed(2);
}

[sgstPercentBox, cgstPercentBox, igstPercentBox].forEach(box =>
    box.addEventListener("input", calculateGST)
);

function collectFilledItems() {
    const items = [];
    tableBody.querySelectorAll("tr").forEach(row => {
        const qty = parseFloat(row.querySelector(".qty").value) || 0;
        const particular = row.querySelector(".particular").value.trim();
        const amount = parseFloat(row.querySelector(".amount").value) || 0;
        if (qty === 0 && particular === "" && amount === 0) return;
        items.push({ qty, particular, amount });
    });
    return items;
}

// ---------- NEW ----------

document.getElementById("newBtn").addEventListener("click", async () => {
    const ok = await appConfirm("Start a new invoice? Unsaved changes will be lost.");
    if (!ok) return;
    startNewInvoice();
});

// ---------- SAVE ----------

// The invoice form only asks for a customer Name — everything else
// (address, city, state, GSTIN, phone, pincode) comes from the
// customer record on file, added/edited via the Customers button.
// A name typed here that doesn't match any saved customer just goes
// through with the extra fields blank (and gets created as a bare
// customer record on save).
function buildCustomerFromDb(name) {
    const match = name ? db.findCustomerByName(name) : null;
    return {
        name,
        address: match?.address || "",
        city: match?.city || "",
        state: match?.state || "",
        gst: match?.gstin || "",
        phone: match?.phone || "",
        pincode: match?.pincode || ""
    };
}

function buildInvoiceData() {
    return {
        invoiceId: currentInvoiceId,
        invoiceNo: invoiceNo.value.trim(),
        invoiceDate: invoiceDate.value,
        customer: buildCustomerFromDb(customerNameInput.value.trim()),
        material: getMaterialValue(),
        shippedTo: document.getElementById("shippedTo").value.trim(),
        challan: document.getElementById("challan").value.trim(),
        challanDate: document.getElementById("challanDate").value,
        receiptNo: document.getElementById("receiptNo").value.trim(),
        receiptDate: document.getElementById("receiptDate").value,
        gpBundles: parseFloat(document.getElementById("gpBundles").value) || 0,
        items: collectFilledItems(),
        subtotal: parseFloat(subtotalBox.value) || 0,
        cgstPercent: parseFloat(cgstPercentBox.value) || 0,
        sgstPercent: parseFloat(sgstPercentBox.value) || 0,
        igstPercent: parseFloat(igstPercentBox.value) || 0,
        gstAmount: parseFloat(gstAmountBox.value) || 0,
        grandTotal: parseFloat(grandTotalBox.value) || 0
    };
}
window.buildInvoiceData = buildInvoiceData;

document.getElementById("saveBtn").addEventListener("click", async () => {
    const data = buildInvoiceData();

    if (!data.customer.name) {
        await appAlert("Customer name is required.");
        customerNameInput.focus();
        return;
    }

    if (data.items.length === 0) {
        await appAlert("Add at least one item.");
        return;
    }

    try {
        saveInvoiceData(data);
    } catch (err) {
        console.error(err);
        if (String(err.message).includes("UNIQUE constraint failed: invoices.invoice_no")) {
            const replace = await appConfirm(
                `Invoice No ${data.invoiceNo} already exists. Do you want to replace the old bill with this one?`
            );
            if (replace) {
                const existingId = db.findInvoiceIdByNo(data.invoiceNo);
                if (existingId) {
                    data.invoiceId = existingId;
                    try {
                        saveInvoiceData(data);
                    } catch (err2) {
                        console.error(err2);
                        await appAlert("Could not replace invoice: " + err2.message);
                    }
                } else {
                    await appAlert("Couldn't find the existing invoice to replace. Change the Invoice No and save again.");
                }
            }
        } else {
            await appAlert("Could not save invoice: " + err.message);
        }
    }
});

// Saves (insert or, if data.invoiceId is set, update) and runs the
// same post-save refresh every successful save needs.
function saveInvoiceData(data) {
    db.saveInvoice(data);
    refreshCustomerSuggestions();
    refreshShippedToSuggestions();
    setDbStatus("ready", `Saved ${data.invoiceNo} — Excel updated`);
    startNewInvoice();
}

// Clears the form for the next invoice. Used both by the New button
// and automatically after every successful Save, so a fresh Save
// always creates a new invoice instead of quietly overwriting the
// one you just saved (that was the "can't add a new customer" bug).
function startNewInvoice() {
    currentInvoiceId = null;
    challanAutoFilled = true;
    challanDateAutoFilled = true;
    document.querySelectorAll(".card input, .card textarea, .card select").forEach(el => el.value = "");
    materialOtherField.style.display = "none";
    tableBody.innerHTML = "";
    addRow(); addRow(); addRow();
    cgstPercentBox.value = 0;
    sgstPercentBox.value = 0;
    igstPercentBox.value = 0;
    document.getElementById("gpBundles").value = 0;
    currentParticularSuggestions = [];
    hideParticularSuggestions();
    loadDate();
    invoiceNo.value = db.getNextInvoiceNo();
    mirrorToChallan(invoiceNo.value);
    calculateTotals();
    customerNameInput.focus();
}

// ---------- INITIAL ----------

calculateTotals();