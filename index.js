var XLSX = require('xlsx-style');
var workbook = XLSX.readFile('csv-dump/Test.xlsx', {cellStyles:true});

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
// console.info(Object.keys(sortedArtistMap['Leann Yan']).length) //63 - includes header
// console.log(sortedArtistMap) //63 - includes header


generateExcelFile(sortedArtistMap)
function generateExcelFile(sortedArtistMap) {

  for (var artist in sortedArtistMap ) {
    // TODO: dynamically define
      var range = "A1:I18";

    // var data = appendObject(sortedArtistMap[artist], range)
    sortedArtistMap[artist]['!ref'] = range

    var workbook = {

      "SheetNames": [
        "Main"

      ],
      "Sheets": {
        'Main': sortedArtistMap[artist]
      }
    }
    // if(artist == 'Stacey Test') console.log(workbook['Sheets']['Main'])

      XLSX.writeFile(workbook, artist + '.xlsx');
  }
}

function checkKeys(obj) {
  return Object.keys(obj)
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
