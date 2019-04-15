const fs = require("fs");
const admZip = require('adm-zip');
const Excel = require('excel-class');
const path = require('path');
const electron = require("electron");
const {dialog} = electron.remote;
// var contensHandler = require("HandleContents");

let docFilePath;
var resultList = [];
var tableValuesList = [];
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
        // console.log("nextIndexOfTableTitle: ", nextIndexOfTableTitle);                
        let group = [];
        for (let i = indexOfTableTitle; i < (nextIndexOfTableTitle > indexOfTableTitle ? nextIndexOfTableTitle : rawData.length); i++) {
            let item = rawData[i];
            // console.log();
            group.push(item);
        }
        // console.log("theGroup: \n", group);

        let arrangedKeysAndValues = arrangeTableKeysAndValues(group);
        // console.log("arrangedKeysAndValues: \n", arrangedKeysAndValues);
        tableValuesList.push(arrangedKeysAndValues);

        if (nextIndexOfTableTitle >= 0) {
            indexOfTableTitle = nextIndexOfTableTitle;
        } else {
            break;
        }        
    }

    // console.log("resultList: \n", resultList);
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

        // console.log("currentKey: ", currentKey, "   ", currentKeyIndex);

        if (i == 13) {
            // 如果是“主要内容”，要单独处理
            let mainInfoString = flattedGroup[currentKeyIndex];
            tableValues.push(mainInfoString.substring(5));
        } else {
            if (currentKeyIndex + 2 == nextKeyIndex) {
                tableValues.push(flattedGroup[currentKeyIndex+1]);
            } else {
                tableValues.push(" ");
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

    // console.log("tableValues: ", tableValues);
    // rst.push(tableKeys);
    // rst.push(tableValues);

    return tableValues;
}

// var tableKeys = ["工单编号0", "来电时间1", "热线号码2", "受理单位3", "来电人4", "来电号码5", "联系方式6", "来电人地址7", "问题分类8", "工单分类9", "发生地址10", "被反映单位11", "标题12", "主要内容13", "派单人员14", "派单时间15", "处理意见16", "截止时间17", "处理时限18", "承办单位19", "处理情况20"];
var excelTitles = ["外网编号", "业务类别", "登记时间", "姓名", "问题发生地省", "问题发生地市", "问题发生地区县", "联系电话 注：咨询账号（注册手机号码）", "纳税人识别号", "外网坐席", "二级分类", "反映内容", "被举报人姓名/被投诉单位或个人/被检举人姓名", "被举报人所属单位/被检举人住所", "是否保密", "紧急程度", "流转方向", "接收人"];
function generateFormattedDictBy(tValues) {
    console.log("tValues: \n", tValues);
    var theDict = {};

    for (let i = 0; i < excelTitles.length; i++) {
        let key = excelTitles[i];
        if (i == 0) {
            // 外网编号
            let value = tValues[0];
            let year = value.substr(2, 4);
            let tradeNumber = value.substring(9);
            let outerNetworkNumber = year + tradeNumber;
            theDict[key] = outerNetworkNumber;
        } else if (i == 1) {
            // 业务类别
            theDict[key] = "0305";
        } else if (i == 2) {
            // 登记时间
            theDict[key] = tValues[15];
        } else if (i == 3) {
            // 姓名
            theDict[key] = tValues[4];
        } else if (i <= 6) {
            // 问题发生地省、市、县
            theDict[key] = "";
        } else if (i == 7) {
            // 联系电话 注：咨询账号（注册手机号码）
            theDict[key] = tValues[5];
        } else if (i == 8) {
            // 纳税人识别号
            theDict[key] = "";
        } else if (i == 9) {
            // 外网坐席
            theDict[key] = tValues[14];
        } else if (i == 10) {
            // 二级分类
            theDict[key] = "11";
        } else if (i == 11) {
            // 反映内容
            let feedbackContent = "外网编号：";
            feedbackContent += tValues[0];
            feedbackContent += "\n来电人姓名：";
            feedbackContent += tValues[4];
            feedbackContent += "\n联系电话：";
            feedbackContent += tValues[5];
            feedbackContent += "\n主要内容：";
            feedbackContent += tValues[12];
            feedbackContent += "\n反映内容：";
            feedbackContent += tValues[13];

            theDict[key] = feedbackContent;
        } else if (i == 12) {
            // 被举报人姓名/被投诉单位或个人/被检举人姓名
            theDict[key] = "见举报内容";
        } else if (i == 13) {
            // 被举报人所属单位/被检举人住所
            theDict[key] = "见举报内容";
        } else if (i == 14) {
            // 是否保密
            theDict[key] = "";
        } else if (i == 15) {
            // 紧急程度
            theDict[key] = "";
        } else if (i == 16) {
            // 流转方向
            theDict[key] = "12345工单";
        } else if (i == 17) {
            // 接收人
            theDict[key] = "";
        }
    }

    console.log("generateFormattedDictBy: \n", theDict);
    return theDict;
}

// 导出excel文件的核心函数
function exportDataToExcel(excelPath) {    
    let excel = new Excel(excelPath);
    let excelContents = [];

    for (let i = 0; i < tableValuesList.length; i++) {
        let dict = generateFormattedDictBy(tableValuesList[i]);
        excelContents.push(dict);
    }

    excel.writeSheet(tableTitle, excelTitles, excelContents).then(()=>{
        //do other things
        console.log("Exported");
        // tail work
        docFilePath = "";
        tableTitle = "";
        indexOfTableTitle = 0;
        resultList = [];
        tableValuesList = [];

        dialog.showMessageBox({
            title :'导出成功', type :'info', message : '导出成功'
          });
        return;
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

function importDocFile() {
    var options = {
        defaultPath: '~',
        filters: [
            { name: '', extensions: ['docx', 'DOCX'] }
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

function exportExcel() {
    if (tableValuesList.length == 0) {
        dialog.showMessageBox({
            title :'请先导入docx文件', type :'info', message : '请先导入docx文件'
          });
        return;
    }

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

// Github AccessToken: 8869651861eff11725e11859b84a87857ecc868d 