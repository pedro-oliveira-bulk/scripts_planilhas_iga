/** @OnlyCurrentDoc */
// VARIÁVEL GLOBAL para controle de concorrência dentro da Lib
var EM_PROCESSAMENTO = false;

// 1. FUNÇÃO PRINCIPAL
function executarProcessoCompleto(ss, config) {
  if (EM_PROCESSAMENTO) return;
  EM_PROCESSAMENTO = true;
  try {
    validarStatusGeral(ss, config);
    SpreadsheetApp.flush();

    aplicarRegraDesistenciaAutomatica(ss, config);
    SpreadsheetApp.flush();

    espelharCorrecoesManuais(ss, config);
    SpreadsheetApp.flush();

    carregarAlunosNaPresenca(ss, config);
    
    console.log("Processo concluído.");
  } finally{
    EM_PROCESSAMENTO = false;
  }
}

// 2. LANÇAMENTO DE PRESENÇAS
function salvarPresenca(ss, config, linksImagens) {
  if (EM_PROCESSAMENTO) return;
  EM_PROCESSAMENTO = true;
  
  const lancar = ss.getSheetByName(config.ABAS.LANCAR);
  const dadosSheet = ss.getSheetByName(config.ABAS.DADOS);

  try {
    const dataAula = lancar.getRange("B4").getValue();
    const professor = lancar.getRange("B5").getValue();

    if (!dataAula || !professor) throw new Error("Preencha data e professor!");

    const dadosForm = lancar.getRange("A13:C60").getValues();
    const presentes = [], ausentes = [];

    dadosForm.forEach(([id, nome, check]) => {
      if (!id) return;
      check === true ? presentes.push(String(id).trim()) : ausentes.push(String(id).trim());
    });

    const { lanche, selfie, lista } = linksImagens;

    // Registra na aba de log
    dadosSheet.appendRow([
      dataAula, 
      professor, 
      presentes.join(", "), 
      ausentes.join(", "), 
      lanche, 
      selfie, 
      lista, 
      new Date()
    ]);
    
    SpreadsheetApp.flush(); 
    
    atualizarPagina1ComPresencas(ss, config);
    espelharCorrecoesManuais(ss, config);
    aplicarRegraDesistenciaAutomatica(ss, config);

    return "Presença e fotos registradas com sucesso!";

  } catch (e) {
    return "Erro ao salvar: " + e.message;
  } finally {
    EM_PROCESSAMENTO = false;
  }
}

// 3. ATUALIZA P1
function atualizarPagina1ComPresencas(ss, config) {
  const p1 = ss.getSheetByName(config.ABAS.P1);
  const dadosSheet = ss.getSheetByName(config.ABAS.DADOS);
  const lancar = ss.getSheetByName(config.ABAS.LANCAR);

  if (!p1 || !dadosSheet || !lancar) return;

  // ================================
  // 1. CAPTURA E NORMALIZA DATA
  // ================================
  let dataAulaRaw = lancar.getRange("B4").getValue();

  let dataAula = dataAulaRaw instanceof Date
    ? dataAulaRaw
    : new Date(dataAulaRaw);

  if (isNaN(dataAula)) {
    console.log("Data inválida em B4.");
    return;
  }

  // Remove horário
  dataAula = new Date(
    dataAula.getFullYear(),
    dataAula.getMonth(),
    dataAula.getDate()
  );

  // ================================
  // 2. PEGA ÚLTIMO REGISTRO DA ABA DADOS
  // ================================
  const registrosDados = dadosSheet.getDataRange().getValues();
  if (registrosDados.length < 2) return;

  const ultimaLinhaDados = registrosDados[registrosDados.length - 1];
  const presentes = String(ultimaLinhaDados[2] || "")
    .split(",")
    .map(id => id.trim())
    .filter(id => id !== "");

  // ================================
  // 3. LOCALIZA OU CRIA COLUNA DA DATA
  // ================================
  const lastCol = p1.getLastColumn();
  const headers = p1.getRange(1, 1, 1, lastCol).getValues()[0];

  let colIndex = -1;

  for (let j = 4; j < headers.length; j++) {
    if (headers[j] instanceof Date) {
      const headerDate = new Date(
        headers[j].getFullYear(),
        headers[j].getMonth(),
        headers[j].getDate()
      );

      if (headerDate.getTime() === dataAula.getTime()) {
        colIndex = j;
        break;
      }
    }
  }

  // Se não existir, cria nova coluna
  if (colIndex === -1) {
    colIndex = Math.max(4, lastCol);
    p1.getRange(1, colIndex + 1)
      .setValue(dataAula)
      .setNumberFormat("dd/MM/yyyy");
  }

  // ================================
  // 4. MARCA PRESENÇA
  // ================================
  const ultimaLinhaP1 = p1.getLastRow();
  if (ultimaLinhaP1 < 2) return;

  const idsP1 = p1.getRange(2, 1, ultimaLinhaP1 - 1, 1).getValues();

  const resultados = [];
  const cores = [];

  for (let i = 0; i < idsP1.length; i++) {
    const idAtual = String(idsP1[i][0]).trim();
    const isPresente = presentes.includes(idAtual);

    resultados.push([isPresente ? "P" : "F"]);
    cores.push([isPresente ? config.CORES.P : config.CORES.F]);
  }

  p1.getRange(2, colIndex + 1, resultados.length, 1)
    .setValues(resultados)
    .setBackgrounds(cores);
}

// 4. REGRA DE DESISTÊNCIA
function aplicarRegraDesistenciaAutomatica(ss, config) {
  const sheet = ss.getSheetByName(config.ABAS.P1);
  const range = sheet.getDataRange();
  const dados = range.getValues();
  const cores = range.getBackgrounds();
  let mudou = false;

  for (let i = 1; i < dados.length; i++) {
    let jaDesistiu = false;
    for (let j = 4; j < dados[0].length; j++) {
      if (jaDesistiu) {
        if (dados[i][j] !== "D") {
          dados[i][j] = "D";
          cores[i][j] = config.CORES.D;
          mudou = true;
        }
      } else if (String(dados[i][j]).toUpperCase() === "D") {
        jaDesistiu = true;
      }
    }
  }

  if (mudou) {
    range.setValues(dados);
    range.setBackgrounds(cores);
    validarStatusGeral(ss, config);
  }
}

// 5. SINCRONIZA P1 PARA CORREÇÕES MANUAIS
function espelharCorrecoesManuais(ss, config) {
  const p1 = ss.getSheetByName(config.ABAS.P1);
  let corr = ss.getSheetByName(config.ABAS.CORR);

  if (!p1) {
    console.error("Erro Crítico: Aba Página 1 não encontrada.");
    return;
  }

  if (!corr) {
    corr = ss.insertSheet(config.ABAS.CORR);
    console.log("Aba de Correções criada automaticamente.");
  }

  const rangeP1 = p1.getDataRange();
  const dados = rangeP1.getValues();
  const cores = rangeP1.getBackgrounds();

  corr.clear();
  const destino = corr.getRange(1, 1, dados.length, dados[0].length);
  destino.setValues(dados);
  destino.setBackgrounds(cores);
  
  corr.setFrozenRows(1);
  corr.setFrozenColumns(4);
}

// 6. SALVA CORREÇÕES MANUAIS
function salvarCorrecoesManuais(ss, config) {
  const p1 = ss.getSheetByName(config.ABAS.P1);
  const corr = ss.getSheetByName(config.ABAS.CORR);
  if (!corr) return;

  const dadosP1 = p1.getDataRange().getValues();
  const dadosCorr = corr.getDataRange().getValues();
  const usuario = Session.getActiveUser().getEmail() || "Sistema";
  const agora = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM HH:mm");

  for (let i = 1; i < dadosCorr.length; i++) {
    const id = String(dadosCorr[i][0]);
    const rowIndexP1 = dadosP1.findIndex(r => String(r[0]) === id);
    if (rowIndexP1 === -1) continue;

    for (let j = 4; j < dadosCorr[0].length; j++) {
      const valorNovo = String(dadosCorr[i][j] || "").toUpperCase();
      const valorAntigo = String(dadosP1[rowIndexP1][j] || "").toUpperCase();

      if (valorNovo !== valorAntigo && valorNovo !== "") {
        const celula = p1.getRange(rowIndexP1 + 1, j + 1);
        celula.setValue(valorNovo).setBackground(config.CORES[valorNovo] || "#ffffff");
        celula.setComment(`[${agora} por ${usuario}]: ${valorAntigo} -> ${valorNovo}`);
      }
    }
  }
  aplicarRegraDesistenciaAutomatica(ss, config);
  validarStatusGeral(ss, config);
}

// 7. CARREGA ALUNOS
function carregarAlunosNaPresenca(ss, config) {
  const p1 = ss.getSheetByName(config.ABAS.P1);
  const lancar = ss.getSheetByName(config.ABAS.LANCAR);

  if (!p1 || !lancar) {
    console.error("Erro: Aba P1 ou Lancar não encontrada. Verifique o CONFIG.");
    return;
  }

  const rangeLimpeza = lancar.getRange("A13:C100");
  rangeLimpeza.clearContent();
  rangeLimpeza.removeCheckboxes();
  
  const ultimaLinhaP1 = p1.getLastRow();
  
  if (ultimaLinhaP1 > 1) {
    const numAlunos = ultimaLinhaP1 - 1;
    const dados = p1.getRange(2, 1, numAlunos, 2).getValues();
    const listaFiltrada = dados.filter(r => r[0] !== "" && r[1] !== "");

    if (listaFiltrada.length > 0) {
      lancar.getRange(13, 1, listaFiltrada.length, 2).setValues(listaFiltrada);
      lancar.getRange(13, 3, listaFiltrada.length, 1).insertCheckboxes();
    }
  }
}

// 8. STATUS GERAL
function validarStatusGeral(ss, config) {
  const abasParaValidar = [config.ABAS.P1, config.ABAS.CORR];
  
  abasParaValidar.forEach(nome => {
    const sheet = ss.getSheetByName(nome);
    
    if (sheet && sheet.getLastRow() > 1) {
      const dados = sheet.getDataRange().getValues();
      const novosStatus = [];
      
      for (let i = 1; i < dados.length; i++) {
        const linhaDeDatas = dados[i].slice(4);
        const temD = linhaDeDatas.some(c => String(c).toUpperCase() === "D");
        novosStatus.push([temD ? "Desistente" : "Cursando"]);
      }
      
      if (novosStatus.length > 0) {
        sheet.getRange(2, 4, novosStatus.length, 1).setValues(novosStatus);
      }
    }
  });
}

// 9. SALVAR IMAGEM
function salvarImagem(base64, tipo, dataStr, config) {
  try {
    const folderId = config.PASTAS_DRIVE[tipo]; // Aqui ele busca 'lanche', 'selfie' ou 'lista'
    if (!folderId) throw new Error("ID da pasta não encontrado para " + tipo);
    
    const bytes = Utilities.base64Decode(base64.split(",")[1]);
    const blob = Utilities.newBlob(bytes, "image/png", `${tipo}_${dataStr}.png`);
    
    return DriveApp.getFolderById(folderId).createFile(blob).getUrl();
  } catch (e) {
    throw new Error("Falha ao salvar " + tipo + ": " + e.message);
  }
}

function processarEnvioWeb(ss, config, lancheBase64, selfieBase64, listaBase64) {
  try {
    const lancar = ss.getSheetByName(config.ABAS.LANCAR);

    let dataAula = lancar.getRange("B4").getValue();

    if (!(dataAula instanceof Date)) {
      dataAula = new Date(dataAula);
    }

    if (isNaN(dataAula)) {
      throw new Error("Data inválida em B4.");
    }

    const dataStr = Utilities.formatDate(
      dataAula,
      ss.getSpreadsheetTimeZone(),
      "dd-MM-yyyy"
    );

    const links = {
      lanche: "",
      selfie: "",
      lista: ""
    };

    if (lancheBase64) links.lanche = salvarImagem(lancheBase64, "lanche", dataStr, config);
    if (selfieBase64) links.selfie = salvarImagem(selfieBase64, "selfie", dataStr, config);
    if (listaBase64)  links.lista  = salvarImagem(listaBase64,  "lista",  dataStr, config);

    return salvarPresenca(ss, config, links);

  } catch (e) {
    return "Erro no processamento: " + e.message;
  }
}
