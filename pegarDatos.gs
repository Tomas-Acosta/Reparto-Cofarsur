function pegado() {
  var spreadsheet = SpreadsheetApp.getActive();

  if (spreadsheet.getRange('BORRADOR!A1') === null) {
    Logger('Borrador en blanco, no se ejecutará la funcion pegado')
  } else {
    /* START COPY DATES TO 'BORRADOR' */

    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('PRE-PLANILLA'), true);
    spreadsheet.getRange('BORRADOR!A:B').copyTo(spreadsheet.getRange('C4'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    spreadsheet.getRange('BORRADOR!F:F').copyTo(spreadsheet.getRange('E4'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

    /* END COPY DATES */

    /* STAR FILTER DATES */

    spreadsheet.setCurrentCell(spreadsheet.getRange('A52'));
    spreadsheet.getRange('A52').copyTo(spreadsheet.getRange('A4:A52'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    spreadsheet.getRange('B4').activate();
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.getRange('B:B').activate();

    /* Filtra para la ZONA R1 */
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('PRE-PLANILLA'), true);
    var criteria = SpreadsheetApp.newFilterCriteria()
      .setHiddenValues(['', 'R2', 'R3'])
      .build();
    spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(1, criteria);
    spreadsheet.getRange('B3:C3').activate();
    var currentCell = spreadsheet.getCurrentCell();
    spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
    currentCell.activateAsCurrentCell();

    /* Pegado de datos filtrados de PREPLANILLA a R1 */
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('R1'), true);
    spreadsheet.getRange('\'PRE-PLANILLA\'!C3:E33').copyTo(spreadsheet.getRange('A4'), SpreadsheetApp.CopyPasteType.PASTE_NO_BORDERS, false);
    spreadsheet.getRange('A4:D34').setBackground('BACKGROUND')
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
    spreadsheet.getRange('A4:D34').setVerticalAlignment('middle')
      .setBackground('BACKGROUND');
    correccionesPlanillas();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('PRE-PLANILLA'), true);

    /* restaura los filtros de PREPLANILLA */
    spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(1, criteria);
    criteria = SpreadsheetApp.newFilterCriteria()
      .build();
    spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(1, criteria);

    /* Filtra para la ZONA R2 */
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('PRE-PLANILLA'), true);
    var criteria = SpreadsheetApp.newFilterCriteria()
      .setHiddenValues(['', 'R1', 'R3'])
      .build();
    spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(1, criteria);
    spreadsheet.getRange('B3:C3').activate();
    var currentCell = spreadsheet.getCurrentCell();
    spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
    currentCell.activateAsCurrentCell();

    /* Pegado de datos filtrados de PREPLANILLA a R2 */
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('R2'), true);
    spreadsheet.getRange('\'PRE-PLANILLA\'!C3:E33').copyTo(spreadsheet.getRange('A4'), SpreadsheetApp.CopyPasteType.PASTE_NO_BORDERS, false);
    spreadsheet.getRange('A4:D34').setBackground('BACKGROUND')
      .setVerticalAlignment('middle');
    correccionesPlanillas();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('PRE-PLANILLA'), true);

    spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(1, criteria);
    criteria = SpreadsheetApp.newFilterCriteria()
      .build();
    spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(1, criteria);

    /* BORDERS R1 AND R2 */
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('R1'), true);
    spreadsheet.getRange('A5:E34').activate().setBorder(false, false, false, false, false, false);
    spreadsheet.getRange('C4:D4').activate().mergeAcross();

    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('R2'), true);
    spreadsheet.getRange('A5:E34').setBorder(false, false, false, false, false, false);
    spreadsheet.getRange('C4:D4').activate().mergeAcross();

    /* Encuadramos los datos de cada cliente con sus respectivos remitos y bultos */

    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('R1'), true);
    ocultarCeldasR1();
    encuadrarClientes();
    correccionesPlanillas();

    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('R2'),);
    ocultarCeldasR2();
    encuadrarClientes();
    correccionesPlanillas();

    /* Volvemos a la hoja BORRADOR */
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('BORRADOR'), true);

    Utilities.sleep(500)
  }
};
