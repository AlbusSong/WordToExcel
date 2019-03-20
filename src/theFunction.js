const fs = require("fs");
const admZip = require('adm-zip');
const Excel = require('excel-class');
const path = require('path');
const electron = require("electron");
const {dialog} = electron.remote;

let docFilePath;

function importDocFile() {
    var options = {
        defaultPath: '~',
        filters: [
            { name: '', extensions: ['doc', 'docx', 'DOC', 'DOCX'] }
          ],
        properties: ['openFile']
    }
    dialog.showOpenDialog(options, (fileNames) => {
         // fileNames is an array that contains all the selected
        if(fileNames === undefined){
            console.log("No file selected");
            return;
        } else {
            docFilePath = fileNames[0];
            console.log("importDocFile: " + docFilePath);
            extractDataFromWordFile(docFilePath);
            //   document.getElementById("message").innerHTML = "已选择文件：" + fileNames[0] + "<br><br>" + "正在导出数据请稍候...";
            //   var filepath = fileNames[0];
            }
            // start_process(filepath);
        }
    );
}

function extractDataFromWordFile(thePath) {
    console.log('aaaaaaaaadsdddd');
    const zip = new admZip(thePath);
    let contentXml = zip.readAsText("word/document.xml");
    let str = "";
    let tmpStr = "";
    const filePath = __dirname + "/2.txt";

    var resultList = [];

    //正则匹配出对应的<w:p>里面的内容,方法是先匹配<w:p>,再匹配里面的<w:t>,将匹配到的加起来即可
    //注意？表示非贪婪模式(尽可能少匹配字符)，否则只能匹配到一个<w:p></w:p>    
    var matchedWP = contentXml.match(/<w:p.*?>.*?<\/w:p>/gi);
    if (matchedWP) 
    {
        matchedWP.forEach(function(wpItem) 
        {
            //注意这里<w:t>的匹配，有可能是<w:t xml:space="preserve">这种格式，需要特殊处理
            var matchedWT = wpItem.match(/(<w:t>.*?<\/w:t>)|(<w:t\s.[^>]*?>.*?<\/w:t>)/gi);
            var textContent = '';
            if(matchedWT) 
            {
                matchedWT.forEach(function(wtItem) 
                {
                    //如果不是<w:t xml:space="preserve">格式
                    if( wtItem.indexOf('xml:space') === -1) 
                    {
                       textContent += wtItem.slice(5,-6);
                    } else 
                    {
                        textContent += wtItem.slice(26,-6);
                    }
                });

                resultList.push(textContent)
            }
        });
    }
    console.log(resultList);
}

function oldExtractDataFromWordFile() {
    console.log('aaaaaaaaadsdddd');
    const zip = new admZip(__dirname + '/test.docx');
    let contentXml = zip.readAsText("word/document.xml");
    let str = "";
    let tmpStr = "";
    const filePath = __dirname + "/2.txt";
    let arrOfItem = contentXml.match(/<w:t>[\s\S]*?<\/w:t>/ig);
    console.log("arrOfItem: \n" + arrOfItem);
    for (let i = 0; i < arrOfItem.length; i++) {
        let item = arrOfItem[i];
        console.log("\n" + item)
        tmpStr = item.slice(5,-6);
        console.log("\n tmpStr: " + tmpStr);
        str = str + tmpStr + "\n";   
    }
    console.log("finalStr: " + str);
    fs.writeFile(filePath, str, (error) => {
        if(error) {
            throw error; 
        }
    });
}

function exportExcel() {
    let fileNameIndex = docFilePath.lastIndexOf("/") + 1;
    let fileNameOfDocFile = docFilePath.substr(fileNameIndex);
    console.log("fileNameOfDocFile: " + fileNameOfDocFile);
    const options = {
        title: fileNameOfDocFile,
        filters: [
          { name: "", extensions: ['xlsx'] }
        ],
        properties: ['openFile']
      }
      dialog.showSaveDialog(options, function (filename) {
        // event.sender.send('saved-file', filename)
        console.log(filename);
        exportDataToExcel(filename);
      })
}


function exportDataToExcel(excelPath) {
    console.log("bbbbbbbbbbbbbbbb");
    let excel = new Excel(excelPath)
    excel.writeSheet('Sheet1', ['name','age','country\ncococo'], [
        {
            name: 'Jane\n\njjjjj',
            age: 19,
            country: 'China'
        },
        {
            name: 'Maria',
            age: 20,
            country: 'America'
        }
    ]).then(()=>{
        //do other things
        console.log("Exported");
    });
}

function readText() {
    const content = fs.readFileSync(__dirname + "/a.txt", "utf8");
    console.log("the Content: " + content);  
}



// Useful electron links
// https://www.jianshu.com/p/57d910008612
// https://www.cnblogs.com/buzhiqianduan/p/7620099.html

// Useful word excel data handling links
// https://www.jb51.net/article/107802.htm
// https://www.jb51.net/article/145571.htm
// https://www.cnblogs.com/xiashan17/p/6214817.html
// https://github.com/laoqiren/excel-class/blob/master/CN.md
// https://cnodejs.org/topic/5846c5914c17b38d35436412
// https://github.com/sail-sail/ejsExcel
// https://www.jianshu.com/p/48dc84f391c0