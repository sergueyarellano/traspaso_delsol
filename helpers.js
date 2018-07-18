const moment = require('moment')

module.exports = {
  copyColIntoSheet,
  formatDates,
  formatIVAs,
  formatBoolean
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

function formatFamilyCode (code) {
}
