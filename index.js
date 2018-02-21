"use strict";

const excel = require('./lib/excel')

let buildExport = params => {
  if (!(params instanceof Array)) throw 'buildExport expects an array'

  let sheets = []
  params.forEach(function (sheet, index) {
    let specification = sheet.specification
    let dataset = sheet.data
    let sheet_name = sheet.name || 'Sheet' + (index + 1)
    let data = []
    let merges = sheet.merges
    let config = {
      cols: []
    }

    if (!specification || !dataset) throw 'missing specification or dataset.'

    if (sheet.heading) {
      sheet.heading.forEach(function (row) {
        data.push(row)
      })
    }

    //build the header row
    let header = []
    for (let col in specification) {
      header.push({
        value: specification[col].displayName,
        style: specification[col].headerStyle || ''
      })

      if (specification[col].width) {
        if (Number.isInteger(specification[col].width)) {
          config.cols.push({ wpx: specification[col].width })
        } else if (Number.isInteger(parseInt(specification[col].width))) {
          config.cols.push({ wch: specification[col].width })
        } else {
          throw 'Provide column width as a number'
        }
      } else {
        config.cols.push({})
      }

    }
    data.push(header) //Inject the header at 0

    dataset.forEach(record => {
      let row = []
      for (let col in specification) {
        let cell_value = record[col]

        if (specification[col].cellFormat && typeof specification[col].cellFormat == 'function') {
          cell_value = specification[col].cellFormat(record[col], record)
        }

        if (specification[col].cellStyle && typeof specification[col].cellStyle == 'function') {
          cell_value = {
            value: cell_value,
            style: specification[col].cellStyle(record[col], record)
          }
        } else if (specification[col].cellStyle) {
          cell_value = {
            value: cell_value,
            style: specification[col].cellStyle
          }
        }
        row.push(cell_value) // Push new cell to the row
      }
      data.push(row) // Push new row to the sheet
    })

    sheets.push({
      name: sheet_name,
      data: data,
      merge: merges,
      config: config
    })

  })

  return excel.build(sheets)

}

module.exports = {
  buildExport
}
