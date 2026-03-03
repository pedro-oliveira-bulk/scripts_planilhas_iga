/**
 * Serve o HTML para a planilha cliente.
 * O arquivo "upload.html" deve estar no projeto da Biblioteca.
 */
function servirInterface() {
  return HtmlService.createHtmlOutputFromFile("upload")
    .setTitle("Registro de Presença")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Função ponte que recebe os dados do HTML e processa
 * @param {Spreadsheet} ss - Objeto da planilha cliente
 * @param {Object} config - Configurações (Pastas, Abas) da cliente
 */
/**function processarEnvioWeb(ss, config, lancheBase64, selfieBase64, listaBase64) {
  Logger.log("Iniciando processamento via Biblioteca para a planilha: " + ss.getName());

  if (!lancheBase64 || !selfieBase64 || !listaBase64) {
    throw new Error("Todas as fotos devem ser enviadas.");
  }

  // Busca a data na aba da cliente para nomear os arquivos
  const lancar = ss.getSheetByName(config.ABAS.LANCAR);
  const dataAula = lancar.getRange("B4").getValue();
  const dataStr = dataAula instanceof Date ? 
    Utilities.formatDate(dataAula, ss.getSpreadsheetTimeZone(), "dd-MM-yyyy") : "data";

  // 1. Salva as imagens usando o salvarImagem que já está no Codigo.gs da biblioteca
  const linkLanche = salvarImagem(lancheBase64, "lanche", dataStr, config);
  const linkSelfie = salvarImagem(selfieBase64, "selfie", dataStr, config);
  const linkLista  = salvarImagem(listaBase64,  "lista",  dataStr, config);

  // 2. Chama a função de salvar presença da biblioteca passando a planilha cliente
  const links = { lanche: linkLanche, selfie: linkSelfie, lista: linkLista };
  salvarPresenca(ss, config, links);

  return "ok";
}*/
