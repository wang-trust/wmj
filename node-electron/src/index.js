// electron-main/index.ts
import { app, BrowserWindow } from "electron";
import path from "path";

import { dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

import { ipcMain } from "./indexEvent.js";


const createWindow = () => {
    const win = new BrowserWindow({
        webPreferences: {
            // contextIsolation: false, // 是否开启隔离上下文
            //   nodeIntegration: true, // 渲染进程使用Node API
            preload: path.join(__dirname, "./preload.js"), // 需要引用js文件
        },
    })

    win.loadFile(path.join(__dirname, 'index.html'));
    // win.webContents.openDevTools();

    
}


app.whenReady().then(() => {
    createWindow() // 创建窗口
    app.on("activate", () => {
        if (BrowserWindow.getAllWindows().length === 0) createWindow()
    })
})

// 关闭窗口
app.on("window-all-closed", () => {
    if (process.platform !== "darwin") {
        app.quit()
    }
})
