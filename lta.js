const Excel = require('exceljs')
const path = require('path')
const {getSheets, writeFilePure} = require('./helpers')
const WB = new Excel.Workbook()
const WS_NAME = 'Artículos'
const WS_TARGET_NAME = 'lta'
const SOURCE_FILE = path.resolve(__dirname, 'xls/Artículos.xlsx')
const TARGET_FILE = path.resolve(__dirname, 'xls/traspaso/LTA.xlsx')

WB.xlsx.readFile(SOURCE_FILE)
  .then(() => getSheets(WB, WS_TARGET_NAME, WS_NAME))
  .then((sheet) => {
    sheet.base.eachRow((row, rowNumber) => {
      let artCode = row.getCell('A').value
      let tar1 = row.getCell('D').value
      let tar2 = row.getCell('E').value
      let tar3 = row.getCell('F').value
      let tar4 = row.getCell('AN').value
      let tar5 = row.getCell('AO').value

      if (rowNumber !== 1) {
        sheet.target.addRow([1, artCode, 0, tar1])
        sheet.target.addRow([2, artCode, 0, tar2])
        sheet.target.addRow([3, artCode, 0, tar3])
        sheet.target.addRow([4, artCode, 0, tar4])
        sheet.target.addRow([5, artCode, 0, tar5])
      }
    })

    return sheet
  })

  .then((sheet) => writeFilePure(sheet, WB, WS_NAME, TARGET_FILE))
  .catch((e) => console.error(e))
