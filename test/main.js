var tap = require('tap')
var XLSX = require('xlsx-style');
var expectedBodyOutput = require('./expectedBodyValues.js')

require('../index.js')

var workbook = XLSX.readFileSync('./Leann\ Yan.xlsx')
var workbookValues = workbook.Sheets.Main

// NOTE: these tests don't test as dep into the objects as they could but
// it should be close enough for now
tap.equal(JSON.stringify(expectedBodyOutput), JSON.stringify(workbookValues))
tap.equal(Object.keys(workbook.Sheets)[0], 'Main')

