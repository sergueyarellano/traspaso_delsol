const Excel = require('exceljs')
const path = require('path')
const {copyColIntoSheet, formatIVAs, formatDates, formatBoolean, getSheets, writeFile} = require('./helpers')
const WB = new Excel.Workbook()
const WS_NAME = 'Artículos'
const WS_TARGET_NAME = 'art'
const SOURCE_FILE = path.resolve(__dirname, 'xls/Artículos.xlsx')
const TARGET_FILE = path.resolve(__dirname, 'xls/ART.xlsx')

WB.xlsx.readFile(SOURCE_FILE)
  .then(() => getSheets(WB, WS_TARGET_NAME, WS_NAME))
  .then((sheet) => copyColIntoSheet(sheet, 'A', 'A')) // codigo
  .then((sheet) => copyColIntoSheet(sheet, 'Z', 'B')) // codigo barras
  .then((sheet) => copyColIntoSheet(sheet, 'C', 'F')) // descripcion
  .then((sheet) => copyColIntoSheet(sheet, 'Q', 'I')) // Proveedor habitual
  .then((sheet) => copyColIntoSheet(sheet, 'J', 'J', formatIVAs)) // Tipo de IVA
  .then((sheet) => copyColIntoSheet(sheet, 'G', 'K')) // Precio de costo
  .then((sheet) => copyColIntoSheet(sheet, 'U', 'O', formatDates)) // Fecha alta
  .then((sheet) => copyColIntoSheet(sheet, 'P', 'AM', formatBoolean)) // Tratar stock
  .then((sheet) => copyColIntoSheet(sheet, 'V', 'AP', formatDates)) // Fecha ultima modificacion
  .then((sheet) => writeFile(sheet, WB, WS_NAME, TARGET_FILE))
  .catch((e) => console.error(e))
