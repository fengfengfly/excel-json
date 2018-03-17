# <center>excel-json</center>
####profile:
主要功能:
解析excel文件为json格式文件, 转换 json文件为excel文件.
Parse excel file to json file, convert json file to excel file.
>依赖的node框架:[node-xlsx](https://github.com/mgcrea/node-xlsx)
####run:
1. 首先`yarn install` 安装依赖.
2. cd 到 src 目录下:
3. excel转json:
   - 执行命令: `node xlsxToJson 目标xlsx文件的路径` 
   - excel文件每个sheet工作表转换为一个json文件, 文件名默认为sheet表名;
   - excel文件第一行对应的单元格值默认为转换json数据的key值;
4. json转excel:
   - 执行命令: `node jsonToXlsx 目标json文件的路径 转换后xlsx指定路径(可省略)` 
   - json文件格式需满足:整个json文件为一个数组对象, 数组每一个元素将会转换为excel文件的每行数据;
   - json对象的key值转换为excel表第一行的数据;