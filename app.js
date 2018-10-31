let data = require('./data.json');
let dot = require('dot-object');
let excelExport = require('./exporter');
let alpha = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
let headers = {};
let map = {};
let index = 0;
let formatedData = []
let exportFile = 'out.xlsx';

for( let name in dot.dot(data[0]) )
{
    map[alpha[index]] = name;
    headers[alpha[index]+"1"] = name;
    index++
}

for( let i = 0; i < data.length; i++ )
{
    formatedData.push( dot.dot(data[i]) );
}

 excelExport.singleWorkbook({
    headers: headers,
    letterDataMap: map,
    sheetName: 'test'
}, formatedData, exportFile, function () {
    console.log('export callback', arguments);
});