const fs = require("fs");
const admZip = require('adm-zip');
const Excel = require('excel-class');
const path = require('path');
const electron = require("electron");
const {dialog} = electron.remote;
// var contensHandler = require("HandleContents");

let docFilePath;
var resultList = [];

var tableTitle;
var indexOfTableTitle = 0;
function processRawData(rawData) {
    if (rawData.length == 0) {
        return [];
    }

    tableTitle = rawData[0];
    console.log("tableTitle: " + tableTitle);

    var groupedData = [];
    while (indexOfTableTitle >= 0) {        
        let nextIndexOfTableTitle = rawData.indexOf(tableTitle, indexOfTableTitle+1);
        console.log("nextIndexOfTableTitle: ", nextIndexOfTableTitle);                
        let group = [];
        for (let i = indexOfTableTitle; i < (nextIndexOfTableTitle > indexOfTableTitle ? nextIndexOfTableTitle : rawData.length); i++) {
            let item = rawData[i];
            // console.log();
            group.push(item);
        }
        console.log("theGroup: \n", group);

        let arrangedKeysAndValues = arrangeTableKeysAndValues(group);
        // console.log("arrangedKeysAndValues: \n", arrangedKeysAndValues);
        
        if (nextIndexOfTableTitle >= 0) {
            indexOfTableTitle = nextIndexOfTableTitle;
        } else {
            break;
        }        
    }    
}

var tableKeys = ["工单编号", "来电时间", "热线号码", "受理单位", "来电人", "来电号码", "联系方式", "来电人地址", "问题分类", "工单分类", "发生地址", "被反映单位", "标题", "主要内容", "派单人员", "派单时间", "处理意见", "截止时间", "处理时限", "承办单位", "处理情况"];
function arrangeTableKeysAndValues (flattedGroup) {
    var rst = [];
    if (flattedGroup.length == 0) {
        return rst;
    }

    let tableValues = [];
    for (let i = 0; i < tableKeys.length - 1; i++) {
        let currentKey = tableKeys[i];
        let nextKey = tableKeys[i+1];        

        // let currentKeyIndex = flattedGroup.indexOf(currentKey);
        let currentKeyIndex = flattedGroup.findIndex(function(value, index, arr) {
            return (value.indexOf(currentKey) > -1);
        });
        let nextKeyIndex = flattedGroup.findIndex(function(value, index, arr) {
            return (value.indexOf(nextKey) > -1);
        });

        console.log("currentKey: ", currentKey, "   ", currentKeyIndex);

        if (i == 13) {
            // 如果是“主要内容”，要单独处理
            let mainInfoString = flattedGroup[currentKeyIndex];
            tableValues.push(mainInfoString.substring(5));
        } else {
            if (currentKeyIndex + 2 == nextKeyIndex) {
                tableValues.push(flattedGroup[currentKeyIndex+1]);
            } else {
                tableValues.push("EMPTY");
            }
        }        

        // 如果是最后两个key
        // if (i == tableKeys.length - 2) {
        //     if (flattedGroup.length > (tableKeys.length - 1)) {
        //         tableValues.push(flattedGroup[nextKeyIndex+1]);
        //     } else {

        //     }
        // }

        // if (currentKeyIndex < 0 || nextKeyIndex < 0) {
        //     tableValues.push("EMPTY");
        //     continue;
        // }        
    }

    console.log("tableValues: ", tableValues);

    rst.push(tableKeys);
    rst.push(tableValues);

    return rst;
}

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
    // contensHandler.handleContents();
    processRawData(resultList);
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
          { name: fileNameOfDocFile, extensions: ['xlsx'] }
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
    let excelTitles = [];
    let excelContents = [];
    
    var theDict = {};
    var tmpKey;
    for (let i = 1; i < resultList.length; i++) {
        if (i % 2 == 1) {
            tmpKey = resultList[i];
            excelTitles.push(tmpKey);
        } else {
            // let theKey = excelTitles[i/2];
            var tmpV = resultList[i];
            theDict[tmpKey] = tmpV;     
            console.log("theDict: " + tmpKey + "\n" + tmpV + "\n");
            // excelContents.push(tmpDict);
        }
    }

    console.log("excelTitles: " + excelTitles);
    // console.log("excelContents: " + excelContents);

    excel.writeSheet(resultList[0], excelTitles, [theDict]).then(()=>{
        //do other things
        console.log("Exported");
    });

    // excel.writeSheet(resultList[0], ['name','age','countrydcococo'], [
    //     {
    //         name: 'Jane\n\njjjjj',
    //         age: 19,
    //         country: 'China'
    //     },
    //     {
    //         name: 'Maria',
    //         age: 20,
    //         country: 'America'
    //     }
    // ]).then(()=>{
    //     //do other things
    //     console.log("Exported");
    // });
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