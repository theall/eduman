var xlsx = require('node-xlsx');
var fs = require('fs');

var excel = {
    'save': function(data) {
        let sheetData = [];
        sheetData.push(data['heads']);
        let cellData = data['data'];
        for(let i=0;i<cellData.length;i++)
            sheetData.push(cellData[i]);
        let xlsData = [
            {
                name: 'sheet1',
                data: sheetData
            }
        ];
        var buffer = xlsx.build(xlsData);
        return buffer;
    },
    
    'load': function(path) {
        return {}
    }
};

module.exports = excel;
