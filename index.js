const path = require('path');
const xlsx = require('xlsx');
const xlsjs = require('xlsjs');
const cvcsv = require('csv');

const options = {
  input: "example.xls"
};

const cvjson = function(csv, callback) {
  var record = []
  var header = []

  cvcsv()
    .from.string(csv)
    .transform(function(row){
      row.unshift(row.pop());
      return row;
    })
    .on('record', function(row, index){
      if(index === 0) {
        header = row;
      }else{
        var obj = {};
        header.forEach(function(column, index) {
          var key = column.trim();
          obj[key] = row[index].trim();
        })
        record.push(obj);
      }
    })
    .on('end', function(count){
      callback(record);
    })
    .on('error', function(error){
      console.error(error.message);
    });
}


const ext = path.extname(options.input);
const xls = (ext === ".xlsx") ? xlsx : xlsjs;
const workbook = xls.readFile(options.input);
const sheet = workbook.SheetNames[0];
const ws = workbook.Sheets[sheet];

const csv = xlsx.utils.make_csv(ws);
cvjson(csv, function(dictionary){
  console.log(dictionary);
});