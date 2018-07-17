const Excel = require('exceljs')
const path = require('path')
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
  .then((sheet) => {
    // Tipo de iva
    // Need a little formatting because the source col has N, R, S, Z for iva type
    const col = sheet.base.getColumn('J').values
    const formattedCol = col.reduce((acc, val) => acc.concat(formatIVAType(val)), [])
    sheet.target.getColumn('J').values = formattedCol
    return sheet
  })
  .then((sheet) => {
    sheet.target.spliceRows(1, 1) // remove first header row
    WB.removeWorksheet('Artículos') // do not need this worksheet for target file
    WB.xlsx.writeFile(path.resolve(__dirname, 'xls/ART.xlsx'))
      .then(function () {
        console.log('done')
      })
  })
  .catch((e) => console.error(e))

function copyColIntoSheet (sheet, colBase, colTarget) {
  const col = sheet.base.getColumn(colBase)
  sheet.target.getColumn(colTarget).values = col.values
  return sheet
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
