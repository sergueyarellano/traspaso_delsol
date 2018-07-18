const Excel = require('exceljs')
const path = require('path')
const moment = require('moment')
const WB = new Excel.Workbook()
const options = {
  dateFormats: ['DD/MM/YYYY']
}
WB.xlsx.readFile(path.resolve(__dirname, 'xls/Artículos.xlsx'), options)
  .then(() => {
    // set up worksheets
    WB.addWorksheet('art')
    const base = WB.getWorksheet('Artículos')
    const target = WB.getWorksheet('art')
    return {base, target}
  })
  .then((sheet) => copyColIntoSheet(sheet, 'A', 'A')) // codigo
  .then((sheet) => copyColIntoSheet(sheet, 'Z', 'B')) // codigo barras
  .then((sheet) => copyColIntoSheet(sheet, 'C', 'F')) // descripcion
  .then((sheet) => copyColIntoSheet(sheet, 'Q', 'I')) // Proveedor habitual
  .then((sheet) => copyColIntoSheet(sheet, 'J', 'J', formatIVAs)) // Tipo de IVA
  .then((sheet) => copyColIntoSheet(sheet, 'G', 'K')) // Precio de costo
  .then((sheet) => copyColIntoSheet(sheet, 'U', 'O', formatDates)) // Fecha alta
  .then((sheet) => copyColIntoSheet(sheet, 'P', 'AM', formatBoolean)) // Tratar stock
  .then((sheet) => copyColIntoSheet(sheet, 'V', 'AP', formatDates)) // Fecha ultima modificacion
  .then((sheet) => {
    sheet.target.spliceRows(1, 1) // remove first header row
    WB.removeWorksheet('Artículos') // do not need this worksheet for target file
    WB.xlsx.writeFile(path.resolve(__dirname, 'xls/ART.xlsx'))
      .then(function () {
        console.log('done')
      })
  })
  .catch((e) => console.error(e))

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
