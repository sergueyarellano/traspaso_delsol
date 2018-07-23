const Excel = require('exceljs')
const path = require('path')
const {copyColIntoSheet, getSheets, writeFile} = require('./helpers')
const WB = new Excel.Workbook()
const WS_NAME = 'Clientes'
const WS_TARGET_NAME = 'cli'
const SOURCE_FILE = path.resolve(__dirname, 'xls/Clientes.xlsx')
const TARGET_FILE = path.resolve(__dirname, 'xls/traspaso/CLI.xlsx')

WB.xlsx.readFile(SOURCE_FILE)
  .then(() => getSheets(WB, WS_TARGET_NAME, WS_NAME))
  .then((sheet) => copyColIntoSheet(sheet, 'A', 'A')) // Codigo
  .then((sheet) => copyColIntoSheet(sheet, 'G', 'C')) // nif
  .then((sheet) => copyColIntoSheet(sheet, 'C', 'D')) // nombre fiscal
  .then((sheet) => copyColIntoSheet(sheet, 'B', 'E')) // nombre comercial
  .then((sheet) => copyColIntoSheet(sheet, 'D', 'F')) // domicilio
  .then((sheet) => copyColIntoSheet(sheet, 'F', 'G')) // poblacion
  .then((sheet) => copyColIntoSheet(sheet, 'E', 'H')) // cp
  .then((sheet) => copyColIntoSheet(sheet, 'I', 'K')) // telefono
  .then((sheet) => copyColIntoSheet(sheet, 'K', 'L')) // FAX
  .then((sheet) => copyColIntoSheet(sheet, 'J', 'M')) // movil
  .then((sheet) => copyColIntoSheet(sheet, 'AE', 'N')) // persona de contacto
  .then((sheet) => copyColIntoSheet(sheet, 'P', 'X')) // tarifa
  .then((sheet) => copyColIntoSheet(sheet, 'O', 'AB')) // tipo de cliente
  .then((sheet) => copyColIntoSheet(sheet, 'BH', 'AN')) // fecha alta
  .then((sheet) => copyColIntoSheet(sheet, 'BC', 'AP')) // email
  .then((sheet) => copyColIntoSheet(sheet, 'AD', 'AT')) // observaciones
  .then((sheet) => copyColIntoSheet(sheet, 'AF', 'BR')) // ruta
  .then((sheet) => copyColIntoSheet(sheet, 'AH', 'CU')) // horario
  .then((sheet) => copyColIntoSheet(sheet, 'N', 'DP')) // estado
  .then((sheet) => writeFile(sheet, WB, WS_NAME, TARGET_FILE))
  .catch((e) => console.error(e))
