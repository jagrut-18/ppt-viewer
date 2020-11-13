const { app, BrowserWindow, dialog, ipcMain } = require("electron");

function createWindow() {
  // Create the browser window.
  const win = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      nodeIntegration: true,
      devTools: false,
    },
    frame: false,
    titleBarStyle: "hidden",
    transparent: true,
  });
  win.setMenuBarVisibility(false);
  // and load the index.html of the app.
  win.loadFile("index.html");

  // Open the DevTools.
  // win.webContents.openDevTools();
  return win;
}
var win = null;
// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
app.whenReady().then(function () {
  win = createWindow();
});
// Quit when all windows are closed.
app.on("window-all-closed", () => {
  // On macOS it is common for applications and their menu bar
  // to stay active until the user quits explicitly with Cmd + Q
  if (process.platform !== "darwin") {
    app.quit();
  }
});
app.on("ready", () => {});
app.on("activate", () => {
  // On macOS it's common to re-create a window in the app when the
  // dock icon is clicked and there are no other windows open.
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

// In this file you can include the rest of your app's specific main process
// code. You can also put them in separate files and require them here.

// ipcMain.on("resize", (e, width, height) => {
//   console.log(width, height);
//   win.setSize(width, height, true);
//   win.center();
// });

ipcMain.on("open-dialog", (event) => {
  dialog
    .showOpenDialog({
      properties: ["openDirectory"],
    })
    .then((e) => {
      event.sender.send("got-folder", e.filePaths[0]);
    })
    .catch((err) => {
      event.sender.send("got-folder", "E-r-r-o-r : " + err);
    });
});

ipcMain.on("close", (event) => {
  win.close();
});
ipcMain.on("maximize", (event) => {
  win.maximize();
});
ipcMain.on("original", (event) => {
  win.unmaximize();
});
ipcMain.on("minimize", (event) => {
  win.minimize();
});
