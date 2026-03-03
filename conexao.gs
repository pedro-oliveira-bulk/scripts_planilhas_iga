// CONFIGURAÇÕES INDIVIDUAIS DO CLIENTE
// Altere os IDs das pastas conforme a necessidade da planilha específica.
const CONFIG = {
  ABAS: {
    P1: "Pagina 1",
    CORR: "Correções Manuais",
    LANCAR: "Lancar_Presenca",
    DADOS: "Presenca_Dados"
  },
  CORES: {
    A: "#cfe2f3", // Azul (Falta Abonada)
    D: "#ea4335", // Vermelho Escuro (Desistente)
    F: "#f4cccc", // Vermelho (Falta)
    P: "#b7e1cd"  // Verde (Presença)
  },
  PASTAS_DRIVE: {
    lanche: "1IjoJD8Fadqyc4Rf8X1CCj306SaPH1WQ4",
    selfie: "1NjCJqxK8YyInrNsO8xd3zqrSaZMNa6o5",
    lista:  "1U6QJ4y2aJBawx7z4ST3wapuWLZZ2BMF_"
  }
};
function autorizarAcessoAoDrive() {
  try {
    const root = DriveApp.getRootFolder();
    console.log("Pasta raiz acessada: " + root.getName());
    
    const drive = DriveApp;
    const metodo = "getFolderById";
    
    drive[metodo](CONFIG.PASTAS_DRIVE.lanche);
    
    SpreadsheetApp.getUi().alert("Autorização concluída com sucesso!");
  } catch (e) {
    SpreadsheetApp.getUi().alert("Clique em 'Revisar Permissões' na janela que apareceu para liberar o Drive.");
    console.error("Erro detalhado: " + e.message);
  }
}
function forcarSincronizacaoGeral() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ModeloSheets.atualizarPagina1ComPresencas(ss, CONFIG);
  ModeloSheets.espelharCorrecoesManuais(ss, CONFIG);
  SpreadsheetApp.getUi().alert("Sincronização forçada concluída!");
}
// SERVIÇO WEB (WEB APP)
function doGet(e) {
  return ModeloSheets.servirInterface();
}

// SALVAMENTO WEB
function salvarPresencaComFotos(lanche, selfie, lista) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ModeloSheets.processarEnvioWeb(ss, CONFIG, lanche, selfie, lista);
}

// MENU E FUNÇÕES
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Sistema de Presenças")
    .addItem("Salvar Correções Manuais", "rodarSalvarCorrecoes")
    .addToUi();
}

function processamentoDiarioCompletoPresencas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ModeloSheets.executarProcessoCompleto(ss, CONFIG);
}

function rodarSalvarCorrecoes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ModeloSheets.salvarCorrecoesManuais(ss, CONFIG);
}
