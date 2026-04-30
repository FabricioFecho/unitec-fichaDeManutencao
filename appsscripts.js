function doPost(e) {

  var sheet = SpreadsheetApp.openById("COLE_O_ID_DA_PLANILHA_AQUI").getSheetByName("Página1");

  var dados = JSON.parse(e.postData.contents);

  var linha = [];

  for (var key in dados) {
    linha.push(dados[key]);
  }

  sheet.appendRow(linha);

  return ContentService
    .createTextOutput("OK")
    .setMimeType(ContentService.MimeType.TEXT);
}