'use strict';

const XLSX = require('xlsx-style');

function datenum(v, date1904) {
  if(date1904) v += 1462;
  let epoch = Date.parse(v);
  return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}

function sheet_from_array_of_arrays(data, merges) {
  let ws = {};
  let range = {s: {c:10000000, r:10000000}, e: {c:0, r:0}};
  for(let R = 0; R !== data.length; ++R) {
    for(let C = 0; C !== data[R].length; ++C) {
      if(range.s.r > R) range.s.r = R;
      if(range.s.c > C) range.s.c = C;
      if(range.e.r < R) range.e.r = R;
      if(range.e.c < C) range.e.c = C;

      let cell;
      if(data[R][C] && typeof data[R][C] === 'object' && data[R][C].style && !(data[R][C] instanceof Date)) {
        cell = {
          v: data[R][C].value,
          s: data[R][C].style
        };
      } else {
        cell = {v: data[R][C] };
      }

      if(cell.v === null) continue;
      let cell_ref = XLSX.utils.encode_cell({c:C,r:R});

      if(typeof cell.v === 'number') {
        cell.t = 'n';
        if(data[R][C] && typeof data[R][C] === 'object'
          && data[R][C].style && typeof data[R][C].style === 'object'
          && data[R][C].style.numFmt) {
          cell.z = data[R][C].style.numFmt;
        }
      }
      else if(typeof cell.v === 'boolean') cell.t = 'b';
      else if(cell.v instanceof Date) {
        cell.t = 'n';
        if(data[R][C] && typeof data[R][C] === 'object'
           && data[R][C].style && typeof data[R][C].style === 'object'
           && data[R][C].style.numFmt) {
          cell.z = data[R][C].style.numFmt;
        } else {
          cell.z = XLSX.SSF._table[14];
        }
        cell.v = datenum(cell.v);
      }
      else cell.t = 's';

      ws[cell_ref] = cell;
    }
  }

  if (merges) {
    if (!ws['!merges']) ws['!merges'] = [];
    merges.forEach(function (merge) {
        ws['!merges'].push({
          s: {
            r: merge.start.row - 1,
            c: merge.start.column - 1
          },
          e: {
            r: merge.end.row - 1,
            c: merge.end.column - 1
          }
        });
    });
  }

  if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
  return ws;
}

function Workbook() {
  if(!(this instanceof Workbook)) return new Workbook();
  this.SheetNames = [];
  this.Sheets = {};
}

module.exports = {
  parse: function(mixed, options) {
    let ws;
    if(typeof mixed === 'string') ws = XLSX.readFile(mixed, options);
    else ws = XLSX.read(mixed, options);
    return _.map(ws.Sheets, function(sheet, name) {
      return {name: name, data: XLSX.utils.sheet_to_json(sheet, {header: 1, raw: true})};
    });
  },
  build: function(array) {
    let defaults = {
      bookType:'xlsx',
      bookSST: false,
      type:'binary'
    };
    let wb = new Workbook();
    array.forEach(function(worksheet) {
      let name = worksheet.name || 'Sheet';
      let data = sheet_from_array_of_arrays(worksheet.data || [], worksheet.merge);
      wb.SheetNames.push(name);
      wb.Sheets[name] = data;

      if(worksheet.config.cols) {
        wb.Sheets[name]['!cols'] = worksheet.config.cols
      }

    });

    let data = XLSX.write(wb, defaults);
    if(!data) return false;
    let buffer = new Buffer(data, 'binary');
    return buffer;

  }
};
