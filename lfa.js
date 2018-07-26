const Excel = require('exceljs')
const path = require('path')
const {getSheets, writeFilePure} = require('./helpers')
const WB = new Excel.Workbook()
const WS_NAME = 'Detalle_Factura_cliente'
const WS_TARGET_NAME = 'lta'
const SOURCE_FILE = path.resolve(__dirname, 'xls/Detalle Factura cliente.xlsx')
const TARGET_FILE = path.resolve(__dirname, 'xls/traspaso/LFA.xlsx')

WB.xlsx.readFile(SOURCE_FILE)
  .then(() => getSheets(WB, WS_TARGET_NAME, WS_NAME))
  .then((sheet) => {
    sheet.base.eachRow((row, rowNumber) => {
    
      let serie = row.getCell('W').value
      let numeroDoc = row.getCell('B').value
      let codArt = row.getCell('C').value
      let descripcion = row.getCell('E').value
      let cantidad = row.getCell('F').value
      let precioArt = row.getCell('H').value
      let total = Number(precioArt) * Number(cantidad)
      let tipoIVA = row.getCell('I').value
      let ivaIncluido = 0
    })

    return sheet
  })

  .then((sheet) => writeFilePure(sheet, WB, WS_NAME, TARGET_FILE))
  .catch((e) => console.error(e))
