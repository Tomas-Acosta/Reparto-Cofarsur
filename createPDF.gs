/*  */
let date = SpreadsheetApp.getActiveSpreadsheet().getRange('PRE-PLANILLA!B1').getValue().toLocaleString().split(', ');
/* Generamos las variables para ponerle nombre y fechas a nuestros archivos pdf */
let pdfNameR1 = 'R1 - ' + date[0]; // Cambia esto al nombre que desees para el PDF
let pdfNameR2 = 'R2 - ' + date[0]; // Cambia esto al nombre que desees para el PDF

date = date[0].split('/');
let OUTPUT_FOLDER_NAME = date[2]; // Cambia esto al nombre de tu carpeta
const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
/* Constituimos los id de las hojas R1 y R2 */
const sheetR1 = SpreadsheetApp.getActive().getSheetByName('R1').getSheetId();
const sheetR2 = SpreadsheetApp.getActive().getSheetByName('R2').getSheetId();



//const pdfFile = createPDF(ssId, sheet, pdfName);

function getFolderByName_(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    // La carpeta no existe, puedes crearla si lo deseas
    return DriveApp.createFolder(folderName);
  }
}

function createPDF(ssId, sheet, pdfName) {
  const url1 = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?format=pdf&size=7&fzr=true&portrait=true&fitw=true&gridlines=false&printtitle=false&top_margin=0.7&bottom_margin=0.25&left_margin=0.25&right_margin=0.25&sheetnames=false&pagenum=UNDEFINED&attachment=true&gid=" + sheetR1;

  const url2 = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?format=pdf&size=7&fzr=true&portrait=true&fitw=true&gridlines=false&printtitle=false&top_margin=0.7&bottom_margin=0.25&left_margin=0.25&right_margin=0.25&sheetnames=false&pagenum=UNDEFINED&attachment=true&gid=" + sheetR2;

  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const blob1 = UrlFetchApp.fetch(url1, params).getBlob().setName(pdfNameR1 + '.pdf');
  const blob2 = UrlFetchApp.fetch(url2, params).getBlob().setName(pdfNameR2 + '.pdf');

  // Obtiene la carpeta en Drive donde se almacenan los PDF.
  const folder = getFolderByName_(OUTPUT_FOLDER_NAME);

  const pdfFileR1 = folder.createFile(blob1);
  const pdfFileR2 = folder.createFile(blob2);

  // Crea una URL para descargar y abrir el PDF.
  const downloadUrlR1 = pdfFileR1.getDownloadUrl();
  const downloadUrlR2 = pdfFileR2.getDownloadUrl();

  const conbinedUrlR1 = downloadUrlR1 + downloadUrlR2;
  const conbinedUrlR2 = downloadUrlR2 + downloadUrlR1;

  // Crea un enlace con estilo de botón en una sola línea de HTML
  var htmlOutput = HtmlService.createHtmlOutput(
    '<div style="display: grid; justify-content: center;"><div style="padding:5px"><a style="display:inline-block;padding:10px 20px;background-color:#096176;color:#fff;text-decoration:none;border-radius:5px;" href="' + conbinedUrlR1 + '"target="_blank">Descargar Zona R1</a></div><div style="padding:10px"></div><div style="padding:5px"><a style="display:inline-block;padding:10px 20px;background-color:#096176;color:#fff;text-decoration:none;border-radius:5px;" href="' + conbinedUrlR2 + '"target="_blank">Descargar Zona R2</a></div></div>'
  ).setHeight(150).setWidth(300);

  // Abre la interfaz web en una ventana modal.
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Planillas listas');

  return pdfFileR1 & pdfFileR2;
}
