'use strict'

var XLSX = require('xlsx-style')
var libreconv = require('libreconv').convert
var path = require('path')
var charArray = 'ABCDEFGHI'.split('')
var currentRowNumber = 2

var workbook = XLSX.readFile('./target.xlsx', { cellStyles: true })
var sheet = workbook.Sheets.Sheet1

var headerRow = getHeaderRow(sheet)
var range = sheet['!ref']

var sortedArtistMap = sortByArtistName(sheet)

function sortByArtistName (sheet) {
  var artistRowCounter = {}
  return Object.keys(sheet).reduce(function (memo, cellNumber) {
    // skip header
    if (parseInt(cellNumber.slice(1)) === 1) return memo

    if (cellNumber[0] === 'A') {
      console.log('sorting new row')
      // collect all of the cells in the row
      var artistDetails = charArray.map(function (letter) {
        var cellLocation = letter + currentRowNumber

        return sheet[cellLocation]
      })

      var artistName = sheet[cellNumber].v

      // sort rows by artist
      if (!(artistName in memo)) {
        var remappedArtistDetailsRow = remapArtistDetails(2, charArray, artistDetails)

        memo[artistName] = appendObject(remappedArtistDetailsRow, headerRow)
        artistRowCounter.artistName = 2
      } else {
        ++artistRowCounter.artistName

        remappedArtistDetailsRow = remapArtistDetails(artistRowCounter.artistName, charArray, artistDetails)
        appendObject(memo[artistName], remappedArtistDetailsRow)

        if (artistRowCounter.artistName % 36 === 0) {
          var remappedHeaderRow = remapHeaderRow(headerRow, ++artistRowCounter.artistName)

          appendObject(memo[artistName], remappedHeaderRow)
        }
      }
      currentRowNumber++
    }

    return memo
  }, {})
}

function remapHeaderRow (header, currentArtistRow) {
  return Object.keys(header).reduce(function (newHeader, cell, index) {
    var column = charArray[index]
    var newCellLocation = column + currentArtistRow

    newHeader[newCellLocation] = headerRow[cell]

    return newHeader
  }, {})
}

function main () {
  var completedFiles = generateExcelFile(sortedArtistMap, range)


  completedFiles.forEach(function (file) {
    var opts = {
      output: './convertedPDFFiles/',
      format: 'pdf'
    }

    convertFileToPDF(file, 'pdf', opts)
  })
}

main()

// #############
// # UTILITIES #
// #############

function generateExcelFile (sortedArtistMap, range) {
  // total column widths: ~91
  var wchColumnWidths = [
    {wch: 10,
      wpx: 40
    },
    {wch: 10,
      wpx: 40
    },
    {wch: 15,
      wpx: 60
    },
    {wch: 7,
      wpx: 60
    },
    {wch: 17,
      wpx: 70
    },
    {wch: 40,
      wpx: 250
    },
    {wch: 20,
      wpx: 70
    },
    {wch: 15,
      wpx: 60
    },
    {wch: 10,
      wpx: 80
    }
  ]

  return Object.keys(sortedArtistMap).reduce(function (completedFiles, artist) {
    console.log('Generating excile file for ' + artist)

    var workBookBody = sortedArtistMap[artist]

    var artistCommissionTotal = getColumnSum(workBookBody, 'H')
    var lastRowNumber = findLastRowNumberOnColumn(workBookBody, 'H')

    var commissionTotalCellLocation = 'H' + (lastRowNumber + 1)
    var totalTitleCellLocation = 'A' + (lastRowNumber + 1)

    workBookBody[commissionTotalCellLocation] = generateCellMetaData(artistCommissionTotal)
    workBookBody[totalTitleCellLocation] = generateCellMetaData(artist + ' Total')

    // # CONFIGS
    workBookBody['!ref'] = range
    workBookBody['!printHeader'] = [1, 1]
    workBookBody['!pageSetup'] = {orientation: 'landscape'}

    workBookBody['!cols'] = wchColumnWidths

    var workbook = {
      'SheetNames': [
        'Main'
      ],
      'Sheets': {
        'Main': workBookBody
      }
    }

    var fileName = artist + '.xlsx'

    XLSX.writeFile(workbook, fileName)
    completedFiles.push(fileName)

    return completedFiles
  }, [])
}

function generateCellMetaData (cellValue) {
  var cellMetaTypes = {
    number: {
      t: 'n',
      v: cellValue,
      s: {
        numFmt: '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-',
        font: {
          bold: true,
          sz: '10',
          color: { theme: '1', rgb: 'FFFFFF' },
          name: 'Calibri'
        },
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
          color: { theme: '1', rgb: 'FFFFFF' },
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

function findLastRowNumberOnColumn (workBookBody, column) {
  var result = Object.keys(workBookBody).reduce(function (lastRowNumber, cell) {
    if (cell[0] === column) {
      var currentRowNumber = parseInt(cell.slice(1))
      lastRowNumber = (currentRowNumber > lastRowNumber) ? currentRowNumber : lastRowNumber
    }
    return lastRowNumber
  }, 0)

  return parseInt(result)
}

function getColumnSum (workBook, columnLetter) {
  return Object.keys(workBook).reduce(function (memo, cell) {
    var rowNumber = cell.slice(1)

    if (cell[0] === columnLetter && rowNumber !== 1 && (!isNaN(workBook[cell].v))) {
      memo += workBook[cell].v
    }

    return memo
  }, 0)
}

function appendObject (target, source) {
  return Object.keys(source).reduce(function (memo, cellLocation) {
    memo[cellLocation] = source[cellLocation]

    return memo
  }, target)
}

function convertFileToPDF (filePath, outputFormat, opts = {}) {
  libreconv(path.join(__dirname, filePath), outputFormat, opts)
}

function remapArtistDetails (currentRowNumber, _, artistDetails) {
  var characterPointer = 0

  return artistDetails.reduce(function (memo, cellDetails) {
    var newCellLocation = charArray[characterPointer] + currentRowNumber

    memo[newCellLocation] = cellDetails
    characterPointer++

    return memo
  }, {})
}

function getHeaderRow (sheet) {
  var header = {}

  for (var cell in sheet) {
    if (cell === '!ref') {
      continue
    } else if (parseInt(cell.slice(1)) === 1) {
      header[cell] = sheet[cell]
    } else {
      break
    }
  }

  return header
}
