var XLSX = require('xlsx-style');
var workbook = XLSX.readFile('./Test.xlsx', {cellStyles:true});
var libreconv = require('libreconv').convert;
var path = require('path');




var sheet = workbook.Sheets.Sheet1

var headerRow = Object.keys(sheet).reduce(function(memo, key) {
  if(key.slice(1) == 1) {
    memo[key] = sheet[key]
  }

  return memo
}, {})

var charArray = "ABCDEFGHI".split('')
var currentRowNumber = 2

var sortedArtistMap = Object.keys(sheet).reduce(function(memo, key) {
  if (key.slice(1) == 1 ) return memo

  if (key[0] == 'A') {
    var artistDetails = charArray.map(function(letter) {
      var cellLocation = letter + currentRowNumber

      return sheet[cellLocation]
    })

    var artistName = sheet[key].v

    if(!(artistName in memo)) {
      var remappedArtistDetailsRow = remapArtistDetails(2, charArray, artistDetails)
      memo[artistName] = appendObject(remappedArtistDetailsRow, headerRow)
      memo[artistName][artistRowCounter] = 1
    } else {
      var artistRowCounter = ++memo[artistName][artistRowCounter]
      var remappedArtistDetailsRow = remapArtistDetails(artistRowCounter, charArray, artistDetails)
      appendObject(memo[artistName], remappedArtistDetailsRow)
    }
    currentRowNumber++
  }

  return memo
}, {})
// [ 'Leann Yan', 'Stacey Test', 'Mary Jane'  ]
// console.info(Object.keys(sortedArtistMap['Stacey Test']).length) //63 - includes header
// console.info(Object.keys(sortedArtistMap['Mary Jane']).length) //63 - includes header
// console.info(Object.keys(sortedArtistMap['Leann Yan'])) //63 - includes header
// console.log(sortedArtistMap['Leann Yan'].I9)
// console.log(sortedArtistMap) //63 - includes header


function convertFileToPDF(filePath, outputFormat, opts = {}) {
  var opts = {
    output: './convertedFiles/',
    format: 'pdf'
  }

  libreconv(path.join(__dirname, filePath), outputFormat, opts)
}

function main() {
  var completedFiles = generateExcelFile(sortedArtistMap)

  completedFiles.forEach(function(file){
    convertFileToPDF(file, 'pdf')
  })
}

main()


function generateExcelFile(sortedArtistMap) {
  var completedFiles = []
  for (var artist in sortedArtistMap ) {
    // TODO: dynamically define
    var range = "A1:I136";

    var workBookBody = sortedArtistMap[artist]

    var artistCommissionTotal = calculateArtistCommissionTotal(workBookBody)
    // if(artist == 'Leann Yan') console.log(workBookBody)
    var lastRowNumber = findLastRowNumberOnColumn(workBookBody, 'H')

    var commissionTotalCellLocation = 'H' + (lastRowNumber + 1)
    var totalTitleCellLocation = 'A' + (lastRowNumber + 1)

    workBookBody[commissionTotalCellLocation] = generateCellMetaData(artistCommissionTotal)
    workBookBody[totalTitleCellLocation] = generateCellMetaData(artist + ' Total')

    workBookBody['!ref'] = range
    // workBookBody['!cols'] = sheet['!cols']

    workBookBody['!printHeader'] = [1,1]





    // ###########
    // # OPTIONS #
    // ###########

    workBookBody['!pageSetup'] = {orientation: 'landscape'}

    var wchColumnWidths = [
      {wch: 5},
      {wch: 5},
      {wch: 15},
      {wch: 7},
      {wch: 10},
      {wch: 30},
      {wch: 5},
      {wch: 10},
      {wch: 5}
    ]
    workBookBody['!cols'] = wchColumnWidths

    var workbook = {
      "SheetNames": [
        "Main"
      ],
      "Sheets": {
        'Main': workBookBody
      }
    }

    var fileName = artist + '.xlsx'
    XLSX.writeFile(workbook, fileName);
    completedFiles.push(fileName)
  }

  return completedFiles
}

function generateCellMetaData(cellValue) {
  var cellMetaTypes = {
    number: {
      t: 'n',
      v: cellValue,
      s: {
        numFmt: '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-',
        font: {
          bold: true,
          sz: '10',
          color: { theme: '1', rgb: 'FFFFFF'  },
          name: 'Calibri' },
        border: {}
      },
      w: ' 10.50 '
    },
    string: {
      t: 's',
      v: cellValue,
      r: "<t>"+ cellValue + "</t>",
      h: "''" + cellValue + "''",
      s: {
        numFmt: 'General',
        font: {
          bold: true,
          sz: '10',
          color: { theme: '1', rgb: 'FFFFFF'  },
          name: 'Calibri'
        },
        border: {}
      },
      w: 'Online'
    }
  }

  var type = typeof cellValue

  return cellMetaTypes[type]
}


function findLastRowNumberOnColumn(workBookBody, column) {
  var result =  Object.keys(workBookBody).reduce(function(lastRowNumber, cell) {
    if(cell[0] == column){
      var currentRowNumber = parseInt(cell.slice(1))
      lastRowNumber = (currentRowNumber > lastRowNumber) ? currentRowNumber : lastRowNumber
    }
    return lastRowNumber
  }, 0)
  console.log(result)

  return parseInt(result)
}

function checkKeys(obj) {
  return Object.keys(obj)
}

function calculateArtistCommissionTotal(workBook) {
  return  Object.keys(workBook).reduce(function(memo, cell) {
    var rowNumber = cell.slice(1)

    if(cell[0] == 'H' && rowNumber != 1)
      memo += workBook[cell].v

    return memo
  }, 0)
}

function appendObject(target, source) {
  return Object.keys(source).reduce(function(memo, cellLocation) {
    memo[cellLocation] = source[cellLocation]

    return memo
  }, target)
}

function remapArtistDetails (currentRowNumber, charArray, artistDetails) {
  var characterPointer = 0
  return artistDetails.reduce(function(memo, cellDetails){
    var newCellLocation = charArray[characterPointer] + currentRowNumber
    memo[newCellLocation] = cellDetails
    characterPointer++

    return memo
  }, {})
  characterPointer = 0
}
