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
    // Força o prompt de autorização de forma genérica
    const root = DriveApp.getRootFolder();
    console.log("Pasta raiz acessada: " + root.getName());
    
    // Usamos esta forma para o Google não travar a execução antes de pedir permissão
    const drive = DriveApp;
    const metodo = "getFolderById";
    
    // Tenta acessar as pastas do seu CONFIG
    drive[metodo](CONFIG.PASTAS_DRIVE.lanche);
    
    SpreadsheetApp.getUi().alert("Autorização concluída com sucesso!");
  } catch (e) {
    // Se o erro for de falta de permissão, o Google abrirá a janela automaticamente
    SpreadsheetApp.getUi().alert("Clique em 'Revisar Permissões' na janela que apareceu para liberar o Drive.");
    console.error("Erro detalhado: " + e.message);
  }
}
function forcarSincronizacaoGeral() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Isso ignora as travas e força a atualização
  ModeloSheets.atualizarPagina1ComPresencas(ss, CONFIG);
  ModeloSheets.espelharCorrecoesManuais(ss, CONFIG);
  SpreadsheetApp.getUi().alert("Sincronização forçada concluída!");
}
// SERVIÇO WEB (WEB APP)
// Esta função faz a ponte para carregar o HTML que está na Biblioteca.
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
