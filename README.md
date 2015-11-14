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
  }
};

//Here you specify the export structure
var specification = {
  customer_name: { // <- the key should match the actual data key
    displayName: 'Customer', // <- Here you specify the column header
    headerStyle: styles.headerDark // <- Header style
  },
  status_id: {
    displayName: 'Status',
    headerStyle: styles.headerDark,
    cellFormat: function(value) { // <- Renderer function
      return (value == 1) ? 'Active' : 'Inactive';
    }
  },
  note: {
    displayName: 'Description',
    headerStyle: styles.headerDark,
    cellStyle: styles.cellPink // <- Cell style [todo: allow function]
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
      name: 'Sheet name', // <- Specify sheet name
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
