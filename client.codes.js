const Excel = require('exceljs')
const path = require('path')
const {getSheets, generateClientCodesJSON} = require('./helpers')
const WB = new Excel.Workbook()
const WS_NAME = 'cli'
const WS_TARGET_NAME = 'cli2'
const SOURCE_FILE = path.resolve(__dirname, 'xls/traspaso/CLI.xlsx')

WB.xlsx.readFile(SOURCE_FILE)
  .then(() => getSheets(WB, WS_TARGET_NAME, WS_NAME))
  .then((sheet) => {
    generateClientCodesJSON(sheet)
  })
  .catch((e) => console.error(e))
