const Excel = require('exceljs')
const path = require('path')
const WB = new Excel.Workbook()

WB.xlsx.readFile(path.resolve(__dirname, 'xls/Artículos.xlsx'))
    .then(() => {
        WB.addWorksheet('art')
        const base = WB.getWorksheet('Artículos')
        const target = WB.getWorksheet('art')
        return {base, target}
    })
    .then((sheet) => {
        const codCol = sheet.base.getColumn('A')
        sheet.target.getColumn('A').values = codCol.values
        
        return sheet
    })
    .then((sheet) => {
        const codCol = sheet.base.getColumn('Z')
    })
    .then(() => {
        WB.removeWorksheet('Artículos')
        WB.xlsx.writeFile(path.resolve(__dirname, 'xls/ART.xlsx'))
            .then(function() {
                
               console.log('done')
            });
    })

  