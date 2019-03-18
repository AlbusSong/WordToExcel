// var admZip = require('adm-zip');
// const zip = new admZip('test.docx');
const fs = require("fs");

// document.querySelector('#btnLoad').addEventListener('click', () => {
//     readText();
// })

function readText(){
    const content = fs.readFileSync(__dirname + "/a.txt", "utf8");
    console.log("the Content: " + content);
    // const content = fs.readFile(__dirname + "/a.txt", "utf8", (err, data) =>  {
    //     console.log(err);
    //     console.log("the Content: " + data);
    // });    
}

function extractDataFromWordFile(){
    console.log('aaaaaaaaadsdddd');
}

// function readText(file){
//     const fs = require('fs');
//     return fs.readFileSync(file, 'utf8');
// }