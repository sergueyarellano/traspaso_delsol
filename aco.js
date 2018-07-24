const Excel = require('exceljs')
const path = require('path')

const {getSheets, writeFilePure} = require('./helpers')
const WB = new Excel.Workbook()
const WS_NAME = 'cli'
const WS_TARGET_NAME = 'aco'
const SOURCE_FILE = path.resolve(__dirname, 'xls/traspaso/CLI.xlsx')
const TARGET_FILE = path.resolve(__dirname, 'xls/traspaso/ACO.xlsx')

function * genId () {
  var id = 1
  while (true) {
    yield id++
  }
}

WB.xlsx.readFile(SOURCE_FILE)
  .then(() => getSheets(WB, WS_TARGET_NAME, WS_NAME))
  .then((sheet) => {
    let id = genId()

    sheet.base.eachRow((row, rowNumber) => {
      let re = /[cCnN]{1,1}$/
      let status = row.getCell('L').value
      const eq = {
        'c': 2,
        'C': 2,
        'n': 1,
        'N': 1
      }

      if (re.test(status) && status) {
        let code = id.next().value
        let clientCode = row.getCell('A').value
        let fecha = '01/11/2017'
        let fechaFin = '01/12/2017'
        let typeAction = eq[status]

        sheet.target.addRow([code, clientCode, fecha, '12:00', 0, 0, 2, typeAction, '', '', fechaFin])
      }
    })

    return sheet
  })

  .then((sheet) => writeFilePure(sheet, WB, WS_NAME, TARGET_FILE))
  .catch((e) => console.error(e))
