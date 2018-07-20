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
  .then((sheet) => copyColIntoSheet(sheet, 'A', 'A'))
  .then((sheet) => writeFile(sheet, WB, WS_NAME, TARGET_FILE))
  .catch((e) => console.error(e))
