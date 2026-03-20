/**
 * SISTEMA APOLLO ENTERPRISE (SAE) - CORE BACKEND
 * Stack: Google Apps Script (V8)
 * Database: Google Sheets (Abas como Entidades)
 * @author Gemini SAE Expert
 */

const DB_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const APP_TITLE = "Gestão EVE - Apollo Enterprise";

/**
 * Roteamento principal do WebApp
 */
function doGet(e) {
  const output = HtmlService.createTemplateFromFile('index');
  return output.evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setTitle(APP_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Inclui arquivos HTML no template principal
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =============================================================================
// UTILITÁRIOS DE BANCO DE DATA (PADRÃO SAE)
// =============================================================================

/**
 * Obtém todos os dados de uma entidade (aba) formatados como JSON
 * @param {string} entityName Nome da aba no Google Sheets
 */
function getTableData(entityName) {
  try {
    const ss = SpreadsheetApp.openById(DB_ID);
    const sheet = ss.getSheetByName(entityName);
    if (!sheet) throw new Error(`Entidade ${entityName} não encontrada.`);

    const values = sheet.getDataRange().getValues();
    const headers = values.shift();
    
    return values.map(row => {
      const obj = {};
      headers.forEach((header, i) => {
        let val = row[i];
        // Normalização de Datas para ISO
        if (Object.prototype.toString.call(val) === '[object Date]') {
          val = val.toISOString();
        }
        obj[header] = val;
      });
      return obj;
    });
  } catch (e) {
    console.error(`Erro em getTableData(${entityName}): ` + e.message);
    return [];
  }
}

/**
 * Salva ou Atualiza um registro com lógica de UUID e Logs
 */
function saveRecord(entityName, data, idFieldName) {
  const ss = SpreadsheetApp.openById(DB_ID);
  let sheet = ss.getSheetByName(entityName);
  
  if (!sheet) throw new Error("Entidade inexistente");

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const timestamp = new Date().toISOString();
  
  // Se não tem ID, é um CREATE
  if (!data[idFieldName]) {
    data[idFieldName] = generateSAEId(entityName);
    const newRow = headers.map(h => data[h] || "");
    sheet.appendRow(newRow);
    logAction("CREATE", entityName, data[idFieldName], data);
    return { success: true, id: data[idFieldName], action: 'create' };
  } 
  
  // Se tem ID, é um UPDATE
  const values = sheet.getDataRange().getValues();
  const idColIndex = headers.indexOf(idFieldName);
  
  for (let i = 1; i < values.length; i++) {
    if (values[i][idColIndex] == data[idFieldName]) {
      const rowRange = sheet.getRange(i + 1, 1, 1, headers.length);
      const updatedRow = headers.map(h => data[h] !== undefined ? data[h] : values[i][headers.indexOf(h)]);
      rowRange.setValues([updatedRow]);
      logAction("UPDATE", entityName, data[idFieldName], data);
      return { success: true, id: data[idFieldName], action: 'update' };
    }
  }
  
  throw new Error("Registro não encontrado para atualização");
}

/**
 * Gera ID sequencial baseado nos parâmetros do sistema
 */
function generateSAEId(entityName) {
  const ss = SpreadsheetApp.openById(DB_ID);
  const paramSheet = ss.getSheetByName("PARAMETROS");
  const data = paramSheet.getDataRange().getValues();
  
  let key = "";
  if (entityName === "EMPRESAS") key = "NEXT_ID_E";
  if (entityName === "PROJETOS") key = "NEXT_ID_P";
  if (entityName === "TAREFAS") key = "NEXT_ID_T";

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      let nextId = parseInt(data[i][1]);
      let prefix = key.split('_').pop(); // E, P ou T
      let formattedId = `${prefix}-${nextId.toString().padStart(4, '0')}`;
      
      // Atualiza o contador para o próximo
      paramSheet.getRange(i + 1, 2).setValue(nextId + 1);
      return formattedId;
    }
  }
  return "UUID-" + Utilities.getUuid().substring(0, 8);
}

/**
 * Grava Log de auditoria (Conforme seu arquivo LOGS.csv)
 */
function logAction(acao, entidade, entidadeId, detalhes) {
  try {
    const sheet = SpreadsheetApp.openById(DB_ID).getSheetByName("LOGS");
    sheet.appendRow([
      new Date(),
      acao,
      entidade,
      entidadeId,
      JSON.stringify(detalhes)
    ]);
  } catch(e) {
    console.warn("Falha ao gravar log");
  }
}

// =============================================================================
// API DE INTERFACE (CHAMADAS PELO FRONTEND)
// =============================================================================

/**
 * Retorna o estado inicial completo para o Vue (Single Request Pattern)
 */
function getInitialAppData() {
  return {
    empresas: getTableData("EMPRESAS"),
    projetos: getTableData("PROJETOS"),
    tarefas: getTableData("TAREFAS"),
    parametros: getTableData("PARAMETROS"),
    config: {
      version: "1.0.0-SAE",
      user: Session.getActiveUser().getEmail()
    }
  };
}

/**
 * Wrappers específicos para ações do Dashboard
 */
function apiSaveEmpresa(data) { return saveRecord("EMPRESAS", data, "empresaId"); }
function apiSaveProjeto(data) { return saveRecord("PROJETOS", data, "projetoId"); }
function apiSaveTarefa(data) { return saveRecord("TAREFAS", data, "tarefaId"); }
