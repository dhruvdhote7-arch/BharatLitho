const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');

function createWindow() {
    const win = new BrowserWindow({
        width: 1200,
        height: 800,
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false
        }
    });

    win.loadFile('src/index.html');
}

// print.js (the renderer) sends this once it has populated the
// #printInvoice overlay with the current invoice's data. The
// stationery already supplies its own margins (see @page in
// print.css), so we print with no extra Chromium margins or
// header/footer, and let the person pick their printer/tray as usual.
ipcMain.on('print-invoice', (event) => {
    const win = BrowserWindow.fromWebContents(event.sender);
    if (!win) return;

    win.webContents.print(
        {
            silent: false,
            printBackground: true,
            margins: { marginType: 'none' }
        },
        (success, errorType) => {
            if (!success && errorType) {
                console.error('Print failed:', errorType);
            }
        }
    );
});

app.whenReady().then(() => {
    createWindow();

    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    });
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});