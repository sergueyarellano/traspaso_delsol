const Excel = require('exceljs')
const path = require('path')
const {formatFamilyCode, copyColIntoSheet, writeFile, getSheets, generateFAMJSON} = require('./helpers')
const WB = new Excel.Workbook()
const WS_NAME = 'Familias_de_artículos'
const WS_TARGET_NAME = 'fam'
const SOURCE_FILE = path.resolve(__dirname, './xls/Familias de artículos.xlsx')
const TARGET_FILE = path.resolve(__dirname, './xls/traspaso/FAM.xlsx')

WB.xlsx.readFile(SOURCE_FILE)
  .then(() => getSheets(WB, WS_TARGET_NAME, WS_NAME))
  .then((sheet) => copyColIntoSheet(sheet, 'A', 'A', formatFamilyCode))
  .then((sheet) => copyColIntoSheet(sheet, 'B', 'B'))
  .then((sheet) => generateFAMJSON(sheet))
  .then((sheet) => writeFile(sheet, WB, WS_NAME, TARGET_FILE))
  .catch((e) => console.error(e))

  // TODO: Crear un json con las equivalencias entre familias antiguas y nuevas para poder reemplazar en ART.xlsx
