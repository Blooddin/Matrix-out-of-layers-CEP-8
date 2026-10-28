/* 1) Create an instance of CSInterface. */
var csInterface = new CSInterface();

function evalScriptPromise(script) {
    return new Promise((resolve, reject) => {
        
        csInterface.evalScript(script, (result) => {
            if (result === undefined || result === null) {
                reject(new Error("Скрипт вернул undefined: " + script));
            } else {
                resolve(result);
            }
        });
    });
}


const paragraph = document.getElementById("myParagraph"); 
var makeMatrixButton = document.getElementById("make-matrix");

makeMatrixButton.addEventListener("click", makeMatrix);

async function makeMatrix() {
  try {
        paragraph.textContent = "Выполнение...";
        var count = await evalScriptPromise('getSelectedLayers();');
        paragraph.textContent = count;
    } catch (error) {
        console.error("Ошибка:", error);
        paragraph.textContent = "Ошибка: " + error.message;
    }

}


