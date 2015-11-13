
'use strict';

const excel = require('./lib/excel');
module.exports = {
  buildExport: function(params) {
    if( ! (params instanceof Array)) throw 'buildExport expects an array';

    let sheets = [];
    params.forEach(function(sheet, index) {
      let specification = sheet.specification;
      let dataset = sheet.data;
      let sheet_name = sheet.name || 'Sheet' + index+1;

      if( ! specification || ! dataset) throw 'missing specification or dataset.';

      //build the header row
      let header = [];
      for (let col in specification) {
        header.push({
          value: specification[col].displayName,
          style: (specification[col].headerStyle) ? specification[col].headerStyle : undefined
        });
      }
      let data = [header]; //Inject the header at 0

      dataset.forEach(record => {
        let row = [];
        for (let col in specification) {
          let cell_value = record[col];

          if(specification[col].cellFormat && typeof specification[col].cellFormat == 'function') {
            cell_value = specification[col].cellFormat(cell_value);
          }

          if(specification[col].cellStyle) {
            cell_value = {
              value: cell_value,
              style: specification[col].cellStyle
            }
          }
          row.push(cell_value) // Push new cell to the row
        }
        data.push(row); // Push new row to the sheet
      });

      sheets.push({
        name: sheet_name,
        data: data
      });

    });

    return excel.build(sheets);

  }
}
