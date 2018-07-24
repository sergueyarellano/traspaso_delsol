const moment = require('moment')
const fs = require('fs')
const path = require('path')
const clientTypes = require('./tmp/cli.type.json')

module.exports = {
  copyColIntoSheet,
  formatDates,
  formatIVAs,
  formatBoolean,
  formatFamilyCode,
  writeFile,
  getSheets,
  generateFAMJSON,
  generateClientCodesJSON,
  formatFamilies,
  writeFilePure,
  formatTypeClient,
  filterEmail,
  formatRoute,
  formatTimeTable,
  formatStatus,
  formatTlf
}

function writeFile (sheet, WB, wsName, targetFile) {
  sheet.target.spliceRows(1, 1) // remove first header row
  WB.removeWorksheet(wsName) // do not need this worksheet for target file
  WB.xlsx.writeFile(targetFile)
    .then(function () {
      console.log('done')
    })
}
function writeFilePure (sheet, WB, wsName, targetFile) {
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
function formatTypeClient (values) {
  return values.map((type) => clientTypes[type])
}

function formatStatus (values) {
  // being 3 dado de baja
  // and 1 habitual
  return values.map((isObsolete) => isObsolete ? 3 : 1)
}

function formatRoute (values) {
  const routes = {
    '0': 'XXX',
    '1': 'LUN',
    '2': 'MAR',
    '3': 'MIE',
    '4': 'JUE',
    '5': 'VIE',
    '6': 'XXX',
    '7': 'XXX'
  }
  return values.map((route) => routes[route])
}

function filterEmail (values) {
  const re = /(.*)@(.*)([.])(.*)/
  return values.map((posibleEmail) => {
    const isValid = re.test(posibleEmail)
    return isValid ? posibleEmail : ''
  })
}

function formatTimeTable (values) {
  const closeDays = {
    '0': 'Sin asignar',
    '1': 'Lunes',
    '2': 'Martes',
    '3': 'Miércoles',
    '4': 'Jueves',
    '5': 'Viernes',
    '6': 'Sábado',
    '7': 'Domingo'
  }
  return values.map((closeDay) => `Cierre: ${closeDays[closeDay]}`)
}

function formatDates (dates) {
  return dates.map((date) => {
    return typeof date !== 'string' ? moment(date).format('DD/MM/YYYY') : date
  })
}

function formatTlf (tlfs) {
  return tlfs.map((tlf) => {
    const rawtlf = tlf.replace(/[-_ ]/g, '')
    return `${rawtlf.substring(0, 3)} ${rawtlf.substring(3, 6)} ${rawtlf.substring(6)}`
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
  let newCodes = []

  return codes.map((code) => {
    code = code.replace(/\s/g, '')
    const match = code.match(re)
    let newCode = match ? match[1] + match[2] : match

    newCode = generateCode(newCode, newCodes)

    match && newCodes.push(newCode)
    return match ? newCode : trimTo3Chars(code)
  })
}

function generateCode (code, list) {
  if (!isDuplicated(code, list)) {
    return code
  } else {
    let splitted = code.split('')
    let unicodeTail = splitted[2].codePointAt(0)
    let newTail = String.fromCharCode(unicodeTail + 1)
    let newCode = splitted[0] + splitted[1] + newTail
    return generateCode(newCode, list)
  }
}

function isDuplicated (str, list) {
  return list.some((val) => {
    return val === str
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

function generateClientCodesJSON (sheet) {
  const oldCodes = sheet.base.getColumn('DQ').values
  const newCodes = sheet.base.getColumn('A').values

  const output = oldCodes.reduce((acc, oldCode, i) => {
    acc[oldCode] = newCodes[i]
    return acc
  }, {})

  const content = JSON.stringify(output)

  fs.writeFile(path.resolve(__dirname, 'tmp/code.cli.eq.json'), content, 'utf8', function (err) {
    if (err) {
      return console.log(err)
    }

    console.log('The file was saved!')
  })
  return sheet
}
