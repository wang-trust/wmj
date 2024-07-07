import { ipcMain } from "electron";
import { dialog } from "electron";
import fs from "fs";
import path from "path";
import child_process from "child_process";


const logpath = 'result.log'


ipcMain.handle('e::getCheckPath', (e) => {
    return new Promise((resolve, reject) => {
        let newDialog = dialog.showOpenDialog({
            buttonLabel: 'Please Select',
            title: '请选择文件或文件夹',
            properties: ['openDirectory', 'multiSelections'],
            // filters: [
            //     { 'name': '代码文件', extensions: ['js', 'json']},
            //     { 'name': '图片文件', extensions: ['jpg', 'png']},
            // ]
        });
        resolve(newDialog);
    });
});


ipcMain.on('e::runCheckPath', (event, filepath, level) => {
    // core function


    // log
    child_process.exec('start notepad ' + logpath);

});


ipcMain.handle('e::getWritePath', (e) => {
    return new Promise((resolve, reject) => {
        let newDialog = dialog.showOpenDialog({
            buttonLabel: 'Please Select',
            title: '请选择文件或文件夹',
            properties: ['openDirectory', 'multiSelections'],
            // filters: [
            //     { 'name': '代码文件', extensions: ['js', 'json']},
            //     { 'name': '图片文件', extensions: ['jpg', 'png']},
            // ]
        });
        resolve(newDialog);
    });
});

ipcMain.on('e::runWritePath', (event, filepath, level) => {
    // core function
    // console.log('e::runWritePath');

    // log
    child_process.exec('start notepad ' + logpath);

});



// word
function readWord(filepath){
    console.log(`readWord filepath = [${filepath}]`);

}


// not use for this program
function recursionDicPath(filepath, fileDic){
    let substat = fs.statSync(filepath);

    if(substat.isDirectory()){
        let subfilelist = fs.readdirSync(filepath);
        
        if(fileDic[filepath] === undefined){
            fileDic[filepath] = [];
        }
        for (let i of subfilelist) {
            recursionDicPath(path.join(filepath, i), fileDic[filepath]);
        }
    }
    else {
        fileDic.push(filepath);
        return;
    }
}

function recursionPath(filepath, fileArray) {
    let substat = fs.statSync(filepath);

    if (substat.isDirectory()) {
        let subfilelist = fs.readdirSync(filepath);
        for (let i of subfilelist) {
            recursionPath(path.join(filepath, i), fileArray);
        }
    }
    else {
        fileArray.push(filepath);
        return;
    }
}


export {
    ipcMain
}


