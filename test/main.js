var tap = require('tap')
var fs = require('fs')
var XLSX = require('xlsx-style');
var expectedBodyOutput = require('./expectedBodyValues.js')
var expectedStylingOutput = require('./expectedStyling.js')

var converter = require('../index.js')
var workbook = XLSX.readFileSync('./Leann\ Yan.xlsx');
var workbookValues = workbook.Sheets.Main

tap.equal(JSON.stringify(expectedBodyOutput), JSON.stringify(workbookValues))
tap.equal(Object.keys(workbook.Sheets)[0], 'Main')

