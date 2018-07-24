const Excel = require('exceljs')
const path = require('path')
const codeEq = require('./tmp/code.cli.eq.json')
const {getSheets, writeFilePure} = require('./helpers')
const WB = new Excel.Workbook()
const WS_NAME = 'Clientes'
const WS_TARGET_NAME = 'dir'
const SOURCE_FILE = path.resolve(__dirname, 'xls/Clientes.xlsx')
const TARGET_FILE = path.resolve(__dirname, 'xls/traspaso/DIR.xlsx')

WB.xlsx.readFile(SOURCE_FILE)
  .then(() => getSheets(WB, WS_TARGET_NAME, WS_NAME))
  .then((sheet) => {
    const re = /(\d{1,5})([a-zA-Z]*)/

    sheet.base.eachRow((row, rowNumber) => {
      let codigo = row.getCell('A').value
      codigo = codeEq[codigo]
      let direccion = row.getCell('AM').value
      let cp = row.getCell('AN').value
      let poblacion = row.getCell('AO').value
      let hasToBeModified = re.test(poblacion)

      if (hasToBeModified) {
        let match = poblacion.match(re)
        poblacion = match[2]
        cp = match[1]
      }

      if (rowNumber !== 1) {
        sheet.target.addRow([codigo, 1, '', direccion, poblacion, cp, 'Madrid', '', '', 0, 13])
      }
    })

    return sheet
  })

  .then((sheet) => writeFilePure(sheet, WB, WS_NAME, TARGET_FILE))
  .catch((e) => console.error(e))
