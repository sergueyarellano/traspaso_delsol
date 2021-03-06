const Excel = require('exceljs')
const path = require('path')
const {copyColIntoSheet, getSheets, writeFile, formatTypeClient, formatDates, filterEmail, formatRoute, formatTimeTable, formatStatus, formatTlf} = require('./helpers')
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
  .then((sheet) => copyColIntoSheet(sheet, 'I', 'K', formatTlf)) // telefono
  .then((sheet) => copyColIntoSheet(sheet, 'K', 'L')) // FAX
  .then((sheet) => copyColIntoSheet(sheet, 'J', 'M', formatTlf)) // movil
  .then((sheet) => copyColIntoSheet(sheet, 'AE', 'N')) // persona de contacto
  .then((sheet) => copyColIntoSheet(sheet, 'P', 'X')) // tarifa
  .then((sheet) => copyColIntoSheet(sheet, 'O', 'AB', formatTypeClient)) // tipo de cliente
  .then((sheet) => copyColIntoSheet(sheet, 'BH', 'AN', formatDates)) // fecha alta
  .then((sheet) => copyColIntoSheet(sheet, 'BC', 'AP', filterEmail)) // email
  .then((sheet) => copyColIntoSheet(sheet, 'AD', 'AT')) // observaciones
  .then((sheet) => copyColIntoSheet(sheet, 'AH', 'AU', formatTimeTable)) // horario
  .then((sheet) => copyColIntoSheet(sheet, 'AF', 'BR', formatRoute)) // ruta
  .then((sheet) => copyColIntoSheet(sheet, 'N', 'DP', formatStatus)) // estado
  .then((sheet) => {
    const re = /(\d{1,5})([a-zA-Z]*)/
    sheet.target.eachRow((row, rowNumber) => {
      let poblacion = row.getCell('G').value
      let hasToBeModified = re.test(poblacion)

      if (hasToBeModified) {
        let newPob = poblacion.match(re)[2]
        let cp = poblacion.match(re)[1]
        row.getCell('H').value = cp
        row.getCell('G').value = newPob
      }

      row.getCell('I').value = 'Madrid'
      row.getCell('J').value = 'España'
    })

    return sheet
  })
  .then((sheet) => writeFile(sheet, WB, WS_NAME, TARGET_FILE))
  .catch((e) => console.error(e))
