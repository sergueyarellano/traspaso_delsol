const moment = require('moment')

module.exports = {
  copyColIntoSheet,
  formatDates,
  formatIVAs,
  formatBoolean,
  formatFamilyCode,
  writeFile,
  getSheets
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
    return match ? match[1] + match[2] : code
  })
}

function getDuplicates (list) {
  list.forEach(function (item, i) {
    if (list.indexOf(item) !== i) {
      console.log('duplicate item' + item + ' at position ' + i)
    }
  })
}
