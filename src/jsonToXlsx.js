
let xlsx = require('node-xlsx')
let fs = require('fs')
//首个输入参数
let param1 = process.argv[2];
let param2 = process.argv[3];
//解析json格式的数据表-单个表
function parseJsonData(jsonData){
    if(!jsonData){
        console.log('error: data is empty')
        return []
    }
    let rowsData=[]
    let keyRow = []
    for(let key in jsonData[0]){
        keyRow.push(key)
    }
    rowsData.push(keyRow)
    for(let i in jsonData){
        let row=[]
        for(let k in jsonData[i]){
            row.push(jsonData[i][k])
        }
        rowsData.push(row)
    }
    return [rowsData]
}
//表数据生成
function buildXlsxSheets(sheetsData){
    let sheets=[]
    for (let i in sheetsData){
        data = sheetsData[i]
        sheets.push({name:'sheet'+i,data:data})
    }
    return sheets
}
function writeToFile(fileName,sheets){
    let xlsxFile = xlsx.build(sheets)
    fs.writeFileSync(fileName,xlsxFile,'binary')
}

function jsonToXlsx(filePath,fileName){
    let jsonData = JSON.parse(fs.readFileSync(filePath));
    let sheetsData = parseJsonData(jsonData)
    let xlsxSheets = buildXlsxSheets(sheetsData)
    writeToFile(fileName,xlsxSheets)
}

function main(){
    let jsonFilePath,xlsxFilePath;
    if(param1==='--help'||param1==='-help'){
        console.log(`Usage:
        $ node jsonToXlsx jsonFilePath xlsxFileName`)
        return;
    }
    if(!param1){
        console.log('Error warning: please input json filepath !');
        return;
    }
    jsonFilePath = param1;
    if(!param2){
        console.log(param1)
        if(/\.json$/.test(param1)){
            xlsxFilePath = param1.replace(/.json$/,'.xlsx');
        }else{
            xlsxFilePath = param1+'.xlsx';
        }
        console.log('Warning: you did not input xlsx filename, auto named to'+xlsxFilePath);
    }else{
        xlsxFilePath = param2;
    }
    jsonToXlsx(jsonFilePath,xlsxFilePath)
}
main();