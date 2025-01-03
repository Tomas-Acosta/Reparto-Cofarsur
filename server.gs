function doGet() {
  const html = HtmlService.createHtmlOutputFromFile('index')
  return html
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
