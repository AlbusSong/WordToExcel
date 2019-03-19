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



// Useful electron links
// https://www.jianshu.com/p/57d910008612
// https://www.cnblogs.com/buzhiqianduan/p/7620099.html

// Useful word excel data handling links
// https://www.jb51.net/article/107802.htm
// https://www.jb51.net/article/145571.htm
// https://www.cnblogs.com/xiashan17/p/6214817.html