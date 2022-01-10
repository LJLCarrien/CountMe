import console from 'console'
import { app, BrowserWindow, ipcMain } from 'electron'
import path from 'path'
import '../renderer/store'

const fs = require("fs")

/**
 * Set `__static` path to static files in production
 * https://simulatedgreg.gitbooks.io/electron-vue/content/en/using-static-assets.html
 */
if (process.env.NODE_ENV !== 'development') {
  global.__static = require('path').join(__dirname, '/static').replace(/\\/g, '\\\\')
}

let mainWindow
const winURL = process.env.NODE_ENV === 'development'
  ? `http://localhost:9080`
  : `file://${__dirname}/index.html`

let jsonInfoObj

app.on('ready', onReady)
function onReady() {
  createWindow();
  readFile();
  writeFile();
}

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

app.on('activate', () => {
  if (mainWindow === null) {
    createWindow()
  }
})

function createWindow() {

  /**
   * Initial window options
   */
  mainWindow = new BrowserWindow({
    height: 563,
    useContentSize: true,
    width: 1000
  })

  mainWindow.loadURL(winURL)

  mainWindow.on('closed', () => {
    mainWindow = null
  })


}

ipcMain.on('pageFinishInit', (event, arg) => {
  //这里是渲染进程发送来的消息
  console.log('----------------------pageFinishInit', arg);
  if (mainWindow != null) {
    mainWindow.webContents.send('initTitle', jsonInfoObj['title'], jsonInfoObj['secondTitle']);
  }

});

/**
 * Auto Updater
 *
 * Uncomment the following code below and install `electron-updater` to
 * support auto updating. Code Signing with a valid certificate is required.
 * https://simulatedgreg.gitbooks.io/electron-vue/content/en/using-electron-builder.html#auto-updating
 */

/*
import { autoUpdater } from 'electron-updater'

autoUpdater.on('update-downloaded', () => {
  autoUpdater.quitAndInstall()
})

app.on('ready', () => {
  if (process.env.NODE_ENV === 'production') autoUpdater.checkForUpdates()
})
 */


function readFile() {
  const filepath = path.join(__dirname, '../../../../config.json').replace(/\\/g, "\/");
  const exists = fs.existsSync(filepath);
  console.log(filepath);
  if (exists) {
    jsonInfoObj = JSON.parse(fs.readFileSync(filepath));
  }
  else {
    console.log(filepath, '文件不存在');
  }
}
ipcMain.on('onSaveJson', (event, arg) => {
  writeFile(arg)
});
function writeFile(jsonObj) {

  // convert JSON object to string
  const data = JSON.stringify(jsonObj);

  // write JSON string to a file
  fs.writeFile('user.json', data, (err) => {
    if (err) {
      throw err;
    }
    console.log("JSON data is saved.");
  });
}
