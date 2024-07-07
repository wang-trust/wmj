
window.onload = () => {
    console.log('page load over...');

    setClickEvent();
};


var econJs = {
    'fileCheckPath': '',
    'checkLevel': 1,
    'fileWritePath':'',
    'writeLevel': 1
}

function setClickEvent() {
    document.querySelector('#check-path').onclick = getCheckPath;
    document.querySelector('#select-check').onchange = setCheckLevel;
    document.querySelector('#file-path-check').oninput = setCheckPath;
    document.querySelector('#run-check').onclick = runCheckPath;
    
    document.querySelector('#write-path').onclick = getWritePath;
    document.querySelector('#select-write').onchange = setWriteLevel;
    document.querySelector('#file-path-write').oninput = setWritePath;
    document.querySelector('#run-write').onclick = runWritePath;

}

function getCheckPath() {
    let checkText = document.querySelector('#file-path-check');
    let ret = window.electronAPI.getCheckPath();
    ret.then((data) => {
        if (data.canceled === false) {
            console.log(data.filePaths[0]);
            checkText.value = data.filePaths[0];
            econJs.fileCheckPath = data.filePaths[0];
        }
    }).catch((data) => {
        checkText.value = 'path error!'
    });

}

function setCheckLevel() {
    let v1 = document.querySelector('#select-check');
    econJs.checkLevel = v1.value;
}

function setCheckPath() {
    let v1 = document.querySelector('#file-path-check');
    econJs.fileCheckPath = v1.value;
}

function runCheckPath() {
    console.log('runCheckPath');
    if (econJs.fileCheckPath) {
        window.electronAPI.runCheckPath(econJs.fileCheckPath, econJs.checkLevel);
    }

}


function getWritePath(){
    let checkText = document.querySelector('#file-path-write');
    let ret = window.electronAPI.getWritePath();
    ret.then((data) => {
        if (data.canceled === false) {
            console.log(data.filePaths[0]);
            checkText.value = data.filePaths[0];
            econJs.fileWritePath = data.filePaths[0];
        }
    }).catch((data) => {
        checkText.value = 'path error!'
    });
}

function setWriteLevel() {
    let v1 = document.querySelector('#select-write');
    econJs.writeLevel = v1.value;
}

function setWritePath() {
    let v1 = document.querySelector('#file-path-write');
    econJs.fileWritePath = v1.value;
}

function runWritePath() {
    // console.log('runWritePath');
    if (econJs.fileWritePath) {
        window.electronAPI.runWritePath(econJs.fileWritePath, econJs.writeLevel);
    }

}



function getPath() {
    console.log('path...');
    window.electronAPI.preGetPath('path');

}


