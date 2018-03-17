let xlsx = require('node-xlsx')
let fs = require('fs');
// let xlsxPath = `${__dirname}/demo.xlsx`
//首个输入参数
let param = process.argv[2]
let xlsxPath = '';
//解析Excel,
function parseExcel(workSheets,keyRow,parseType,typesRow)
{
    let sheets = []
    for (let i = 0; i < workSheets.length; i++) 
    {
         let sheetData = workSheets[i].data
         let objArray  = []
         let typeArray =  parseType===true?sheetData[typesRow]:[]
         let keyNum = keyRow? keyRow:0
         let keyArray =  sheetData[keyNum]
        for (let j = 1; j < sheetData.length ; j++)
        {
             let row = sheetData[j]
             if(row.length == 0) continue
             let jsonRow = makeObjsWithKeysValues(keyArray,row,typeArray)
             objArray.push(jsonRow)
        }
        if(objArray.length >0) sheets.push({name:workSheets[i].name,data:objArray})
    }
    return sheets
}

//转换数据类型
function makeObjsWithKeysValues(keys,values,typeArray)
{
     let obj = {};
    for (let i = 0; i < values.length; i++) 
    {
        //字母
        obj[keys[i]] = parseType(values[i],typeArray[i]);  
    }
    return obj;
}
function parseType(value,type)
{
    if(value == null || value =="null") return ""
    if(type =="int") return Math.floor(value)
    if(type =="Number") return value
    if(type =="String") return value
    return value 
}
//写文件
function writeSheetToFile(fileName,data)
{  
  fs.writeFile(fileName,data,'utf-8',complete);
  function complete(err)
  {
      if(!err)
      {
          console.log(`file ${fileName} build success..`)
      }   
  } 
}
function xlsxToJson(filePath){
    console.log('start parsing xlsx...')
    let workSheets = xlsx.parse(filePath)
    let sheets = parseExcel(workSheets)
    for(let i in sheets){
        let sheet = sheets[i]
        writeSheetToFile(sheet.name+".json",JSON.stringify(sheet.data))
    }
    console.log("finished")
}
function main(){
    if(param==='--help'||param==='-help'){
        console.log(`Usage:
        $ node xlsxToJson xlsxFilePath`)
        return;
    }
    if(!param){
        console.log('error warning: please input xlsx filepath !');
        return;
    }
    xlsxPath=param;
    xlsxToJson(xlsxPath)
}
main();