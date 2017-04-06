### Node.JS Excel-Export

Nice little module that is assisting when creating excel exports from datasets. It takes normal array-of-objects dataset plus a json report specification and builds excel(.xlsx) file. It supports styling and re-formating of the data on the fly. Check the example usage for more information.

### Installation
```bash
npm install node-excel-export
```

### Usage

```javascript
var excel = require('node-excel-export');

// You can define styles as json object
// More info: https://github.com/protobi/js-xlsx#cell-styles
var styles = {
  headerDark: {
    fill: {
      fgColor: {
        rgb: 'FF000000'
      }
    },
    font: {
      color: {
        rgb: 'FFFFFFFF'
      },
      sz: 14,
      bold: true,
      underline: true
    }
  },
  cellPink: {
    fill: {
      fgColor: {
        rgb: 'FFFFCCFF'
      }
    }
  },
  cellGreen: {
    fill: {
      fgColor: {
        rgb: 'FF00FF00'
      }
    }
  }
};

//Array of objects representing heading rows (very top)
let heading = [
  [{value: 'a1', style: styles.headerDark}, {value: 'b1', style: styles.headerDark}, {value: 'c1', style: styles.headerDark}],
  ['a2', 'b2', 'c2'] // <-- It can be only values
];

//Here you specify the export structure
var specification = {
  customer_name: { // <- the key should match the actual data key
    displayName: 'Customer', // <- Here you specify the column header
    headerStyle: styles.headerDark, // <- Header style
    cellStyle: function(value, row) { // <- style renderer function
      // if the status is 1 then color in green else color in red
      // Notice how we use another cell value to style the current one
      return (row.status_id == 1) ? styles.cellGreen : {fill: {fgColor: {rgb: 'FFFF0000'}}}; // <- Inline cell style is possible 
    },
    width: 120 // <- width in pixels
  },
  status_id: {
    displayName: 'Status',
    headerStyle: styles.headerDark,
    cellFormat: function(value, row) { // <- Renderer function, you can access also any row.property
      return (value == 1) ? 'Active' : 'Inactive';
    },
    width: '10' // <- width in chars (when the number is passed as string)
  },
  note: {
    displayName: 'Description',
    headerStyle: styles.headerDark,
    cellStyle: styles.cellPink, // <- Cell style
    width: 220 // <- width in pixels
  }
}

// The data set should have the following shape (Array of Objects)
// The order of the keys is irrelevant, it is also irrelevant if the
// dataset contains more fields as the report is build based on the
// specification provided above. But you should have all the fields
// that are listed in the report specification
var dataset = [
  {customer_name: 'IBM', status_id: 1, note: 'some note', misc: 'not shown'},
  {customer_name: 'HP', status_id: 0, note: 'some note'},
  {customer_name: 'MS', status_id: 0, note: 'some note', misc: 'not shown'}
]

// Create the excel report.
// This function will return Buffer
var report = excel.buildExport(
  [ // <- Notice that this is an array. Pass multiple sheets to create multi sheet report
    {
      name: 'Sheet name', // <- Specify sheet name (optional)
      heading: heading, // <- Raw heading array (optional)
      specification: specification, // <- Report specification
      data: dataset // <-- Report data
    }
  ]
);

// You can then return this straight
res.attachment('report.xlsx'); // This is sails.js specific (in general you need to set headers)
return res.send(report);

// OR you can save this buffer to the disk by creating a file.

```
