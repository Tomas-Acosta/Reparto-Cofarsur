function ocultarCeldasR1() {
  console.log('R1 ocultar celdas')
  let spreadsheet = SpreadsheetApp.getActive();
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[5];
  let totalFarmacia = spreadsheet.getRange('B3').getValue();
  let totalEspacios = parseInt(totalFarmacia[17]) + 2;
  let celdaDesde = 'A' + (5 + totalEspacios);

  spreadsheet.getRange('A5').activate();
  console.log(totalEspacios)

  if ((5 + totalEspacios) <= 18) {
    sheet.unhideRow(sheet.getRange('A5:A33'));
    sheet.hideRow(sheet.getRange('A19:A33'));
    console.log('R1 <18 ' + celdaDesde);
  }
  if ((5 + totalEspacios) >= 19 && (5 + totalEspacios) <= 27) {
    sheet.unhideRow(sheet.getRange('A5:A34'));
    sheet.hideRow(sheet.getRange(celdaDesde + ':A33'))
    console.log('R1 >19 ' + celdaDesde);
  }
  if ((5 + totalEspacios) > 27) {
    sheet.showRows(5, 28);
    console.log('R1 >27 ' + celdaDesde);
  }
}
/* esta funcion sirve para ocultar las filas respecto a la cantidad de celdas ocupadas, y se divide en 3 casos */
function ocultarCeldasR2() {
  console.log('R2 ocultar celdas')
  let spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('R2'), true);
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[6];
  let totalFarmacia = spreadsheet.getRange('B3').getValue();// tomamos el dato de cuantas farmacias tenemos en la planilla  
  let totalEspacios = parseInt(totalFarmacia[17]);  // lo parseamos como entero al dato
  let celdaDesde = 'A' + (5 + totalEspacios);

  spreadsheet.getRange('A5').activate();
  console.log(totalEspacios)

  if ((5 + totalEspacios) <= 18) {
    sheet.unhideRow(sheet.getRange('A5:A33'));
    sheet.hideRow(sheet.getRange('A19:A33'));
    console.log('R2 <18 ' + celdaDesde)
  }

  if ((5 + totalEspacios) >= 19 && (5 + totalEspacios) <= 27) {
    sheet.unhideRow(sheet.getRange('A5:A34'));
    sheet.hideRow(sheet.getRange(celdaDesde + ':A33'))
    console.log('R2 >19 ' + celdaDesde)
  }

  if ((5 + totalEspacios) > 27) {
    sheet.unhideRow(sheet.getRange("A5:A33"));
    console.log('R2 >27 ' + celdaDesde)
  }
}
