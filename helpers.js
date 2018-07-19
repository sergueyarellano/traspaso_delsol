const moment = require('moment')
const fs = require('fs')
const path = require('path')

module.exports = {
  copyColIntoSheet,
  formatDates,
  formatIVAs,
  formatBoolean,
  formatFamilyCode,
  writeFile,
  getSheets,
  generateFAMJSON,
  formatFamilies
}

function writeFile (sheet, WB, wsName, targetFile) {
  sheet.target.spliceRows(1, 1) // remove first header row
  WB.removeWorksheet(wsName) // do not need this worksheet for target file
  WB.xlsx.writeFile(targetFile)
    .then(function () {
      console.log('done')
    })
}

function getSheets (WB, targetWS, sourceWS) {
  WB.addWorksheet(targetWS)
  const base = WB.getWorksheet(sourceWS)
  const target = WB.getWorksheet(targetWS)
  return {base, target}
}

function copyColIntoSheet (sheet, colBase, colTarget, formatFn) {
  const values = sheet.base.getColumn(colBase).values
  sheet.target.getColumn(colTarget).values = formatFn ? formatFn(values) : values
  return sheet
}

function formatDates (dates) {
  return dates.map((date) => {
    return typeof date !== 'string' ? moment(date).format('DD/MM/YYYY') : date
  })
}

function formatIVAs (IVAs) {
  return IVAs.reduce((acc, val) => acc.concat(formatIVAType(val)), [])
}
function formatIVAType (ivaType) {
  switch (ivaType) {
    case 'N':
      return 1
    case 'R':
      return 2
    case 'S':
      return 3
    default:
      return 0
  }
}

function formatBoolean (booleans) {
  return booleans.map((bool) => bool ? 1 : 0)
}

function formatFamilyCode (codes) {
  const re = /([0-9]{1})(?:[0-9A-Za-z])([0-9A-Za-zñÑ]{2,2})/

  return codes.map((code) => {
    code = code.replace(/\s/g, '')
    const match = code.match(re)
    return match ? match[1] + match[2] : trimTo3Chars(code)
  })
}

function formatFamilies (codes) {
  const famRelations = require('./tmp/fam.json')
  return codes.map((code) => {
    return famRelations[code]
  })
}

function trimTo3Chars (str) {
  const length = 3
  return str.substring(0, length)
}

function generateFAMJSON (sheet) {
  const oldCodes = sheet.base.getColumn('A').values
  const newCodes = sheet.target.getColumn('A').values

  const output = oldCodes.reduce((acc, oldCode, i) => {
    if (oldCode !== 'Código de familia') {
      acc[oldCode] = newCodes[i]
    }
    return acc
  }, {})

  const content = JSON.stringify(output)

  fs.writeFile(path.resolve(__dirname, 'tmp/fam.json'), content, 'utf8', function (err) {
    if (err) {
      return console.log(err)
    }

    console.log('The file was saved!')
  })
  return sheet
}
