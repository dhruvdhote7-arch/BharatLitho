const ITEMS_TOP_MM = 104;
const ROW_GAP_MM = 1.5; // extra breathing room between item rows
const BOTTOM_LIMIT_MM = 250; // stay above the item area's own rule (see print.css)
const LINE_HEIGHT_MM = 4.6; // one line of 11pt/10pt text — matches --line-height-mm
const MM_TO_PX = 96 / 25.4; // browsers always render CSS mm at 96 CSS px/inch

// `require` only exists in an Electron renderer with nodeIntegration on
// (or no contextIsolation). If that's not this window's setup, calling
// require() throws a ReferenceError immediately — which used to happen
// here, before the click listener below was ever attached, so Print did
// nothing at all. Guard it, and fall back to the browser's own print
// dialog if ipcRenderer isn't available.
let ipcRenderer = null;
try {
    ({ ipcRenderer } = require('electron'));
} catch (e) {
    console.warn("Electron ipcRenderer not available — Print will use window.print() instead.");
}

// Returns the current form data if it's valid to print, or null
// (after alerting) if something's missing. Uses the in-page dialog
// (window.appAlert, defined in app.js) instead of window.alert() —
// Electron's native alert() can leave the window unable to receive
// keyboard input afterward, which is the same bug the New button had.
async function getPrintableInvoiceData() {
    const data = window.buildInvoiceData();

    if (!data.customer.name) {
        await window.appAlert("Enter a customer name before printing.");
        return null;
    }
    if (data.items.length === 0) {
        await window.appAlert("Add at least one item before printing.");
        return null;
    }
    return data;
}

// print.html also loads this file but has no #printBtn on it, so guard
// against that instead of throwing on getElementById(...).addEventListener.
const printBtn = document.getElementById("printBtn");
if (printBtn) {
    printBtn.addEventListener("click", async () => {
        const data = await getPrintableInvoiceData();
        if (!data) return;

        populatePrintView(data);

        if (ipcRenderer) {
            ipcRenderer.send("print-invoice");
        } else {
            window.print();
        }
    });
}

// ---------------------------------------------------------------
// Every field on this stationery is its own fixed, individually
// positioned element (see print.css) — Material Supply, Consignee
// block, Invoice No/Date, Subtotal, GST breakdown, Rounded Off,
// Challan, Shipped To, Receipt, Amount-in-Words, and the Grand
// Total. Only the item rows are dynamic, since row count varies —
// they're positioned by a running cursor below.
// ---------------------------------------------------------------

function populatePrintView(data) {
    // Note: this is the CUSTOMER's GSTIN (consignee box, top-left).
    // The seller's own GSTIN/PAN block is already printed on the
    // stationery itself and is never touched here.

    document.getElementById("piConsigneeName").textContent = data.customer.name || "";
    document.getElementById("piConsigneeAddress").textContent = data.customer.address || "";

    // City, State — Pincode (pincode printed right after the state name,
    // on the same line, so it doesn't need a field of its own).
    const cityState = [data.customer.city, data.customer.state].filter(Boolean).join(", ");
    document.getElementById("piConsigneeCityState").textContent =
        data.customer.pincode
            ? (cityState ? cityState + " - " + data.customer.pincode : data.customer.pincode)
            : cityState;

    document.getElementById("piConsigneeGstin").textContent =
        data.customer.gst ? "GSTIN: " + data.customer.gst : "";

    document.getElementById("piInvoiceNo").innerHTML =
        "Tax Invoice No:<br>" + escapePrint(data.invoiceNo);

    document.getElementById("piIssueDateTime").textContent =
        "Date: " + formatDateForPrint(data.invoiceDate);

    const container = document.getElementById("piItemsContainer");
    container.innerHTML = "";
    let cursorMm = ITEMS_TOP_MM;

    // ---- Material Supply — its own fixed field (see #piMaterial in
    // print.css), not part of the item cursor. ----
    setText("piMaterial", data.material || "");

    // ---- item rows — Qty / Description / Amount only. The Total
    // column stays blank here; it's reserved for the running totals
    // below, same as the sample invoice. ----
    data.items.forEach(item => {
        const rowHeightMm = addPrintRow(container, cursorMm, item.qty, item.particular, formatMoney(item.amount));
        cursorMm += rowHeightMm + ROW_GAP_MM;
    });

    if (cursorMm > BOTTOM_LIMIT_MM) {
        console.warn(`Item rows (${cursorMm.toFixed(1)}mm) ran past the item area — consider fewer/shorter item lines, or move --items-bottom-limit further down in print.css.`);
    }

    // ---- Subtotal — now its own fixed field too (see #piSubtotal in
    // print.css), independent of the item-row cursor. ----
    setText("piSubtotal", formatMoney(data.subtotal));

    // ---- CGST / SGST / IGST breakdown — each its own fixed field now
    // (see #piCgstLabel/Value etc. in print.css), so positions can be
    // tuned independently instead of via the shared cursor. ----
    setText("piCgstLabel", "");
    setText("piCgstValue", "");
    setText("piSgstLabel", "");
    setText("piSgstValue", "");
    setText("piIgstLabel", "");
    setText("piIgstValue", "");

    if (data.igstPercent > 0) {
        setText("piIgstLabel", `IGST @ ${trimPercent(data.igstPercent)}%:`);
        setText("piIgstValue", formatMoney(data.gstAmount));
    } else {
        if (data.cgstPercent > 0) {
            const cgstAmt = data.subtotal * (data.cgstPercent / 100);
            setText("piCgstLabel", `CGST @ ${trimPercent(data.cgstPercent)}%:`);
            setText("piCgstValue", formatMoney(cgstAmt));
        }
        if (data.sgstPercent > 0) {
            const sgstAmt = data.subtotal * (data.sgstPercent / 100);
            setText("piSgstLabel", `SGST @ ${trimPercent(data.sgstPercent)}%:`);
            setText("piSgstValue", formatMoney(sgstAmt));
        }
    }

    // Grand total is rounded to the nearest rupee on the printed sheet.
    const grandTotalRounded = Math.round(data.grandTotal);
    const roundOffAmount = grandTotalRounded - data.grandTotal;
    const roundOffText = (roundOffAmount >= 0 ? "+" : "") + roundOffAmount.toFixed(2);

    // ---- Total (subtotal + GST, before rounding) — no "Total" label,
    // just a rule line above the figure. Fixed field, see
    // #piTotalAfterGst in print.css. ----
    setText("piTotalAfterGst", formatMoney(data.grandTotal));

    // ---- Rounded Off — the paisa adjustment applied to reach the
    // final rounded figure (e.g. "+0.35"), not the total itself.
    // Fixed field, see #piRoundOffLabel/Value in print.css. ----
    setText("piRoundOffLabel", "Rounded Off:");
    setText("piRoundOffValue", roundOffText);

    // ---- Shipped To / Receipt — plain lines, only printed when
    // there's something to say, except Receipt which always prints a
    // pre-formatted spot (falls back to "____" for the receipt no.) ----
    // Carrier name is a print-only field (#printCarrierName) — read directly
    // from the DOM here, never added to buildInvoiceData(), so it's never
    // saved to the database. Left blank, it's simply omitted from the line.
    const carrierName = (document.getElementById("printCarrierName")?.value || "").trim();

    // ---- Challan — its own fixed field now (see #piChallan in
    // print.css), only printed when there's something to say. ----
    setText("piChallan", "");
    if (data.challan) {
        const challanLine = `Challan No.${data.challan}${data.challanDate ? " dt " + formatDateForPrint(data.challanDate) : ""}`
            + (carrierName ? ` & ${carrierName}` : "");
        setText("piChallan", challanLine);
    }

    // ---- Shipped To — its own fixed field (see #piShippedTo in
    // print.css), only printed when there's something to say. ----
    setText("piShippedTo", data.shippedTo ? "Shipped To: " + data.shippedTo : "");

    // ---- Receipt — its own fixed field (see #piReceipt in
    // print.css). Only printed when the Carrying Name box has
    // something in it — left blank, the whole Receipt line is
    // skipped instead of falling back to "____". ----
    setText("piReceipt", "");
    if (carrierName) {
        const receiptNoText = data.receiptNo || "____";
        let receiptLineText = `Receipt No.${receiptNoText}`;
        if (data.receiptDate) {
            receiptLineText += ` dt.${formatDateForPrint(data.receiptDate)}`;
        }
        if (data.gpBundles) {
            receiptLineText += ` for ${data.gpBundles} GP Bundles`;
        }
        setText("piReceipt", receiptLineText);
    }

    // ---- Amount in Words — its own fixed field (see #piAmountWords
    // in print.css). ----
    setText("piAmountWords", `Rupees ${numberToWords(grandTotalRounded)} Only`);

    // ---- Grand total — the one fixed field, inside the pre-printed
    // TOTAL box at the bottom of the page ----
    document.getElementById("piGrandTotal").textContent = formatMoney(grandTotalRounded);
}

function addPrintRow(container, topMm, qty, desc, amount) {
    const row = document.createElement("div");
    row.className = "pi-row";
    row.style.top = topMm + "mm";
    row.innerHTML = `
        <div class="pi-qty">${escapePrint(qty)}</div>
        <div class="pi-desc">${escapePrint(desc)}</div>
        <div class="pi-amount">${escapePrint(amount)}</div>
    `;
    container.appendChild(row);

    return measureDescHeightMm(desc);
}

// A row that sits in the Total column, with an optional label in the
// Description column on the same line (e.g. "SGST @ 9%" / "1234.56").
// options.rule draws a rule above the total value, for the Subtotal
// and final Total (Rounded Off) lines.
function addTotalRow(container, topMm, label, value, options = {}) {
    const row = document.createElement("div");
    row.className = "pi-row";
    row.style.top = topMm + "mm";
    row.innerHTML = `
        ${label ? `<div class="pi-desc">${escapePrint(label)}</div>` : ""}
        <div class="pi-total${options.rule ? " pi-total-rule" : ""}">${escapePrint(value)}</div>
    `;
    container.appendChild(row);

    return LINE_HEIGHT_MM;
}

let measureEl = null;
function getMeasureEl() {
    if (measureEl) return measureEl;
    measureEl = document.createElement("div");
    measureEl.style.position = "fixed";
    measureEl.style.visibility = "hidden";
    measureEl.style.left = "-9999px";
    measureEl.style.top = "0";
    measureEl.style.width = "var(--col-desc-width)";
    measureEl.style.fontFamily = "Arial, Helvetica, sans-serif";
    measureEl.style.fontSize = "11pt";
    measureEl.style.lineHeight = "1.3";
    measureEl.style.whiteSpace = "pre-wrap";
    document.body.appendChild(measureEl);
    return measureEl;
}

function measureDescHeightMm(text) {
    const el = getMeasureEl();
    el.textContent = text || "";
    return pxToMm(el.getBoundingClientRect().height);
}

function pxToMm(px) {
    return px / MM_TO_PX;
}

function trimPercent(n) {
    return Number(n) % 1 === 0 ? Number(n).toFixed(0) : Number(n).toString();
}

function formatMoney(n) {
    const num = Number(n || 0);
    const isNegative = num < 0;
    const fixed = Math.abs(num).toFixed(2);
    const [intPart, decPart] = fixed.split(".");
 
    // Indian grouping: last 3 digits, then pairs of 2 from there
    // (e.g. 1234567.89 -> "12,34,567.89", not "1,234,567.89").
    let lastThree = intPart.substring(intPart.length - 3);
    let otherNumbers = intPart.substring(0, intPart.length - 3);
    if (otherNumbers !== "") {
        lastThree = "," + lastThree;
    }
    const formattedInt = otherNumbers.replace(/\B(?=(\d{2})+(?!\d))/g, ",") + lastThree;
 
    return (isNegative ? "-" : "") + formattedInt + "." + decPart;
}
function formatDateForPrint(iso) {
    if (!iso) return "";
    const [y, m, d] = iso.split("-");
    return `${d}-${m}-${y}`;
}

function escapePrint(str) {
    return String(str ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
}

// Sets the text of one of the individualized fixed fields (e.g.
// #piCgstValue, #piChallan) — no-op if the element isn't found, so
// this stays safe even before the matching div is added to the HTML.
function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}

// ---------- Indian-format amount in words ----------

function numberToWords(amount) {
    amount = Math.round((Number(amount) || 0) * 100) / 100;
    const rupees = Math.floor(amount);
    const paise = Math.round((amount - rupees) * 100);

    let words = rupeesToWords(rupees);
    if (paise > 0) {
        words += " and " + rupeesToWords(paise) + " Paise";
    }
    return words;
}

function rupeesToWords(num) {
    if (num === 0) return "Zero";

    const ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
        "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen",
        "Eighteen", "Nineteen"];
    const tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];

    function twoDigits(n) {
        if (n < 20) return ones[n];
        return tens[Math.floor(n / 10)] + (n % 10 ? " " + ones[n % 10] : "");
    }

    function threeDigits(n) {
        if (n < 100) return twoDigits(n);
        return ones[Math.floor(n / 100)] + " Hundred" + (n % 100 ? " and " + twoDigits(n % 100) : "");
    }

    // Indian grouping: crore, lakh, thousand, hundred
    const crore = Math.floor(num / 10000000);
    num %= 10000000;
    const lakh = Math.floor(num / 100000);
    num %= 100000;
    const thousand = Math.floor(num / 1000);
    num %= 1000;
    const hundred = num;

    const parts = [];
    if (crore) parts.push(threeDigits(crore) + " Crore");
    if (lakh) parts.push(threeDigits(lakh) + " Lakh");
    if (thousand) parts.push(threeDigits(thousand) + " Thousand");
    if (hundred) parts.push(threeDigits(hundred));

    return parts.join(" ");
}