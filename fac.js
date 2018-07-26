const Excel = require('exceljs')
const path = require('path')
const {copyColIntoSheet, getSheets, writeFile, formatSerie, formatClientCode} = require('./helpers')
const WB = new Excel.Workbook()
const WS_NAME = 'Facturas_clientes'
const WS_TARGET_NAME = 'fac'
const SOURCE_FILE = path.resolve(__dirname, 'xls/Facturas clientes.xlsx')
const TARGET_FILE = path.resolve(__dirname, 'xls/traspaso/FAC.xlsx')

WB.xlsx.readFile(SOURCE_FILE)
  .then(() => getSheets(WB, WS_TARGET_NAME, WS_NAME))
  .then((sheet) => copyColIntoSheet(sheet, 'BR', 'A', formatSerie)) // serie
  .then((sheet) => copyColIntoSheet(sheet, 'A', 'B')) // numero documento
  .then((sheet) => copyColIntoSheet(sheet, 'C', 'D')) // fecha
  .then((sheet) => copyColIntoSheet(sheet, 'L', 'F')) // cod almacen
  .then((sheet) => copyColIntoSheet(sheet, 'B', 'I', formatClientCode)) // cod cliente
  .then((sheet) => copyColIntoSheet(sheet, 'G', 'J')) // nombre cliente
  .then((sheet) => copyColIntoSheet(sheet, 'H', 'K')) // domicilio
  .then((sheet) => copyColIntoSheet(sheet, 'I', 'L')) // poblacion
  .then((sheet) => copyColIntoSheet(sheet, 'J', 'M')) // cp
  .then((sheet) => copyColIntoSheet(sheet, 'AC', 'AT')) // bi1
  .then((sheet) => copyColIntoSheet(sheet, 'AH', 'AU')) // bi2
  .then((sheet) => copyColIntoSheet(sheet, 'AM', 'AV')) // bi3
  .then((sheet) => copyColIntoSheet(sheet, 'AD', 'AW')) // porcentaje iva1
  .then((sheet) => copyColIntoSheet(sheet, 'AI', 'AX')) // porcentaje iva2
  .then((sheet) => copyColIntoSheet(sheet, 'AN', 'AY')) // porcentaje iva3
  .then((sheet) => copyColIntoSheet(sheet, 'AE', 'AZ')) // cuota iva1
  .then((sheet) => copyColIntoSheet(sheet, 'AJ', 'BA')) // cuota iva2
  .then((sheet) => copyColIntoSheet(sheet, 'AO', 'BB')) // cuota iva3
  .then((sheet) => copyColIntoSheet(sheet, 'AR', 'BK')) // total
  .then((sheet) => writeFile(sheet, WB, WS_NAME, TARGET_FILE))
  .catch((e) => console.error(e))