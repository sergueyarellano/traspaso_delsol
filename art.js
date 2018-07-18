const Excel = require('exceljs')
const path = require('path')
const {copyColIntoSheet, formatIVAs, formatDates, formatBoolean} = require('./helpers')
const WB = new Excel.Workbook()

WB.xlsx.readFile(path.resolve(__dirname, 'xls/Artículos.xlsx'))
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
