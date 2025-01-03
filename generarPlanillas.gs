/* EJECUCION DE LA FUNCION PRINCIPAL // EJECUTE OF THE PRINCIPAL FUNCTION */
function generarPlanillas() {
  let sheet = SpreadsheetApp.getActive();
  let pdfName = '';
  let spreadsheet = SpreadsheetApp.getActive();

  /* PANTALLA DE CARGA */
  let html = HtmlService.createHtmlOutput('<div style="padding-top:75.000%;position:relative;"></div>').setHeight(500);
  SpreadsheetApp.getUi().showModelessDialog(html, ' ');
  Utilities.sleep(500) // TimeOut que ayuda a que no se sature la ejecucion de la funcion

  /* FUNCION PRINCIPAL */
  pegado(); // funcion principal
  /* PANTALLA CHECK */
  
  Utilities.sleep(1000) // TimeOut que ayuda a que no se sature la ejecucion de la funcion
  html = HtmlService.createHtmlOutput('<div class="gif-container"><div class="overlay" style="position: absolute; top: 0; left: 0; width: 100%; height: 100%; background: transparent;"></div></div>').setHeight(500); // pantalla de check en gif
  SpreadsheetApp.getUi().showModelessDialog(html, ' ');
  createPDF(ssId, sheet, pdfName);
}
