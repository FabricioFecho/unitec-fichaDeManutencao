function doPost(e) {

  var sheet = SpreadsheetApp
    .openById("11tG2w7AUZLxlnI8AQ9lDQ8xj61d50iGi9TrCiDRtEW4")
    .getSheetByName("UNITEC");

  var dados = JSON.parse(e.postData.contents);

  var assinaturaUrl = "";

  // ===== SALVAR ASSINATURA =====
  if (dados.assinatura && dados.assinatura.includes("base64")) {
    try {

      var base64 = dados.assinatura.split(",")[1];

      var blob = Utilities.newBlob(
        Utilities.base64Decode(base64),
        "image/png",
        "assinatura_" + new Date().getTime() + ".png"
      );

      var pasta = DriveApp.getFolderById("19x36SyUNs06G1Bi27ziHLPQYh35acCY2");

      var arquivo = pasta.createFile(blob);

      arquivo.setSharing(
        DriveApp.Access.ANYONE,
        DriveApp.Permission.VIEW
      );

      assinaturaUrl = arquivo.getUrl();

    } catch (err) {
      assinaturaUrl = "ERRO: " + err.message;
    }
  }

  // REMOVE assinatura do objeto (pra não mandar base64 gigante)
  delete dados.assinatura;

  var linha = [];

  for (var key in dados) {
    linha.push(dados[key]);
  }

  // adiciona o link no final
  linha.push(assinaturaUrl);

  sheet.appendRow(linha);

  return ContentService
    .createTextOutput("OK")
    .setMimeType(ContentService.MimeType.TEXT);