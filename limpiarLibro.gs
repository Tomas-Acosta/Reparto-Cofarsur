/* Funcion para Limpiar el Libro completo */
function limpiado() {
  let spreadsheet = SpreadsheetApp.getActive();
  let sheetR1 = SpreadsheetApp.getActiveSpreadsheet().getSheets()[6];
  let sheetR2 = SpreadsheetApp.getActiveSpreadsheet().getSheets()[7];

  spreadsheet.getRange('BORRADOR!A:G').clear();
  spreadsheet.getRange('PRE-PLANILLA!C4:E').clear();
  spreadsheet.getRange('C4:E').merge()
    .breakApart();
  spreadsheet.getRange('PRE-PLANILLA!A150').copyTo(spreadsheet.getRange('PRE-PLANILLA!A4:A149'), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false)
  spreadsheet.getRange('PRE-PLANILLA!B4:C52').setBorder(null, null, null, null, true, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  spreadsheet.getRange('R1!A5:D34').clear({ contentsOnly: true, skipFilteredRows: true });
  spreadsheet.getRange('R1!A5:C34').merge().breakApart();
  spreadsheet.getRange('R1!A5:E34').setBorder(false, false, false, false, false, false);
  spreadsheet.getRange('R2!A5:D34').clear({ contentsOnly: true, skipFilteredRows: true });
  spreadsheet.getRange('R2!A5:C34').merge().breakApart();
  spreadsheet.getRange('R2!A5:E34').setBorder(false, false, false, false, false, false);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('BORRADOR'), true);
  spreadsheet.getRange('A1').activate();
};
