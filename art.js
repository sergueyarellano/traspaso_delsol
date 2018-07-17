const Excel = require('exceljs')
const path = require('path')
const 
const WB = new Excel.Workbook()

WB.xlsx.readFile(path.resolve(__dirname, 'xls/Artículos.xlsx'))
    .then(() => {WS: WB.getWorksheet('Artículos')})
    .then(() => WB.addWorksheet('art'))
    .then((WS) => {
        const codCol = WS.getColumn('A')
        
        
        const WStarget = WB.getWorksheet('art');
        WStarget.getColumn('A').values = codCol.values
        
        return WS
    })
    .then((WS) => {
        const codCol = WS.getColumn('Z')
    })
    .then(() => {
        WB.removeWorksheet('Artículos')
        WB.xlsx.writeFile(path.resolve(__dirname, 'xls/ART.xlsx'))
            .then(function() {
                
               console.log('done')
            });
    })

  