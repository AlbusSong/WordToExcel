const fs = require("fs");
var admZip = require('adm-zip');

function readText() {
    const content = fs.readFileSync(__dirname + "/a.txt", "utf8");
    console.log("the Content: " + content);  
}

function extractDataFromWordFile() {
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
    // fs.writeFile(filePath, str, (err) => {
    //     //将./2.txt替换为你要输出的文件路径
    //     if(err) throw err; 
    // }
}

// function readText(file){
//     const fs = require('fs');
//     return fs.readFileSync(file, 'utf8');
// }