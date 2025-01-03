function correccionesPlanillas() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  /* Parche estético para las hojas R1 y R2 */
  spreadsheet.getRange('B3').setFormula('=CONCATENATE("TOTAL FARMACIAS: ";COUNTIF($A$5:$A$34;">A");"  /  __")').setTextStyle(SpreadsheetApp.newTextStyle().setBold(true).setFontSize(12).build())
  spreadsheet.getRange('D3').setFormula('=CONCATENATE("TOTAL DE BULTOS: ";SUBTOTAL(9;$C$5:$C$34);"  /  __")').setTextStyle(SpreadsheetApp.newTextStyle().setBold(true).setFontSize(12).build());
  spreadsheet.getRange('A34:E').clear();
  spreadsheet.getRange('E4').setValue('FIRMA FARMACIA');

  // Damos estilo a Farmacias, Remitos, Bultos por remito y Bultos totales respectivamente.
  spreadsheet.getRange('A5:A34').setTextStyle(SpreadsheetApp.newTextStyle().setBold(true).setFontSize(11).build());
  spreadsheet.getRange('B5:B34').setTextStyle(SpreadsheetApp.newTextStyle().setFontSize(9).build());
  spreadsheet.getRange('C5:C34').setTextStyle(SpreadsheetApp.newTextStyle().setFontSize(9).build());
  spreadsheet.getRange('D5:D34').setTextStyle(SpreadsheetApp.newTextStyle().setBold(true).setFontSize(18).build());
}

/* Funcion de prueba para encuadrar clientes */
function encuadrarClientes() {
  let spreadsheet = SpreadsheetApp.getActive();
  let row = 5  // variable que indica la fila, va a servir de acumulador, comienza en 5 por la celda A5, hasta
  let range = 'A' + row // variable para formar el rango compuesto entre A (columna) y row (fila - acumulador)

  while (row < 33) {
    if (spreadsheet.getRange(range).getValue() > "") {
      spreadsheet.getRange(range).activate();
      spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
      spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
      spreadsheet.getActiveRange().setBorder(true, false, true, false, true, false);
      /* este apartado es para sumar la cantidad de bultos por remito e insertarlos como un total de bultos en la columna D de su respectiva farmacia */
      let valores = spreadsheet.getActiveRange().getValues();
      let sumaColumnaC = 0;
      if (valores.length > 1) {
        for (i = 0; i < valores.length; i++) {
          let valoresColumnaC = valores[i][2]; // buscamos los valores "i" que se encuentren en la columna C, es decir [2] del array en las ordenadas
          sumaColumnaC += parseInt(valoresColumnaC); // lo sumamos parseandolo como un numero entero
        }
        spreadsheet.getRange("D" + row).setValue(sumaColumnaC); // seleccionamos el rango de D + fila en cuestion para insertar el valor que sumamos
      } else {
        spreadsheet.getRange("C" + row).copyTo(spreadsheet.getRange("D" + row), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false); // en caso de que los valores lo anterior no funciones, solo copiamos el valor de C y lo pegamos en D pertenecientes a la misma fila (esto aplica si hay un solo remito, es decir una sola fila, copia y pega el valor de bulto)

        spreadsheet.getRange('D' + row).set
      }
    }
    /* con esto componemos el siguiente rango ej: pasa de A5 a A6 (A + row) siendo row inicial = 5 */
    row = row + 1
    range = 'A' + row
  }
}
