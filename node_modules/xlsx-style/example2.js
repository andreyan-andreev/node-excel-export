var XLSX = require('/Users/pieter/Dev/js-xlsx');
var workbook = XLSX.readFile('/Users/pieter/Exp/js-xlsx/lab/date/date.xlsx', {type: "xlsx", cellDates: true, cellStyles: true});
console.log(workbook.Sheets);