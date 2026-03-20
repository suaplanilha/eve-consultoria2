const APP_TITLE = 'EVE Consultoria | SAE SaaS';
const DB_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const CACHE_KEY_INITIAL_DATA = 'EVE_INITIAL_DATA_V2';
const CACHE_TTL_SECONDS = 180;

const ENTITY_CONFIG = {
  EMPRESAS: {
    idField: 'empresaId',
    headers: ['empresaId', 'nomeRazao', 'nomeFantasia', 'cnpj', 'endereco', 'telefone', 'email', 'responsavel', 'observacao', 'status', 'dataInicio', 'dataFimPrevista', 'createdAt', 'updatedAt']
  },
  PROJETOS: {
    idField: 'projetoId',
    headers: ['projetoId', 'empresaId', 'nomeProjeto', 'escopo', 'faseAtual', 'consultorResponsavel', 'prioridade', 'statusManual', 'statusCalc', 'dataInicio', 'prazoFinal', 'orcamento', 'indicadorMeta', 'indicadorAtual', 'percentualConclusao', 'ultimaAtualizacao', 'createdAt', 'updatedAt']
  },
  TAREFAS: {
    idField: 'tarefaId',
    headers: ['tarefaId', 'projetoId', 'fase', 'descricao', 'responsavel', 'statusManual', 'statusCalc', 'prioridade', 'dataInicio', 'dataFim', 'ordem', 'semanaReferencia', 'dataConclusao', 'detalhes', 'resultado', 'observacao', 'metaTipo', 'metaValor', 'metaUnidade', 'resultadoFinal', 'createdAt', 'updatedAt']
  },
  TAREFAS_HISTORICO: {
    idField: 'historicoId',
    headers: ['historicoId', 'tarefaId', 'statusDe', 'statusPara', 'alteradoEm', 'origem', 'motivoConclusao', 'resultadoFinalNumber', 'resultadoFinalTexto']
  },
  LOGS: {
    idField: 'logId',
    headers: ['logId', 'timestamp', 'acao', 'entidade', 'entidadeId', 'detalhes']
  }
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getInitialAppData(forceRefresh) {
  const cache = CacheService.getScriptCache();
  if (!forceRefresh) {
    const cached = cache.get(CACHE_KEY_INITIAL_DATA);
    if (cached) {
      return JSON.parse(cached);
    }
  }

  const empresas = getEntityRecords_('EMPRESAS');
  const projetos = hydrateProjects_(getEntityRecords_('PROJETOS'), getEntityRecords_('TAREFAS'));
  const tarefas = hydrateTasks_(getEntityRecords_('TAREFAS'));
  const historico = getEntityRecords_('TAREFAS_HISTORICO');
  const summary = buildDashboardSummary_(empresas, projetos, tarefas);
  const projectReports = projetos.map(function(project) {
    return buildExecutiveReportForProject_(project, empresas, tarefas, historico);
  });

  const payload = {
    empresas: empresas,
    projetos: projetos,
    tarefas: tarefas,
    historico: historico,
    summary: summary,
    projectReports: projectReports,
    config: {
      appTitle: APP_TITLE,
      version: '2.0.0-SAE',
      user: safeGetUserEmail_(),
      sheetSchema: buildSheetSchema_(),
      lastSyncAt: new Date().toISOString()
    }
  };

  cache.put(CACHE_KEY_INITIAL_DATA, JSON.stringify(payload), CACHE_TTL_SECONDS);
  return payload;
}

function apiSaveEmpresa(payload) {
  const saved = upsertEntity_('EMPRESAS', payload);
  return buildMutationResponse_('Empresa salva com sucesso.', saved.id, 'EMPRESAS');
}

function apiSaveProjeto(payload) {
  if (!payload.empresaId) {
    throw new Error('Selecione a empresa vinculada ao projeto.');
  }
  const saved = upsertEntity_('PROJETOS', payload);
  return buildMutationResponse_('Projeto salvo com sucesso.', saved.id, 'PROJETOS');
}

function apiSaveTarefa(payload) {
  if (!payload.projetoId) {
    throw new Error('Selecione o projeto antes de salvar a tarefa.');
  }
  const saved = upsertEntity_('TAREFAS', payload, { trackHistory: true, origin: 'APP_FORM' });
  return buildMutationResponse_('Tarefa salva com sucesso.', saved.id, 'TAREFAS');
}

function apiConcluirTarefa(payload) {
  if (!payload || !payload.tarefaId) {
    throw new Error('Tarefa inválida para conclusão.');
  }

  const task = findRecordById_('TAREFAS', payload.tarefaId);
  if (!task) {
    throw new Error('Tarefa não encontrada.');
  }

  const merged = Object.assign({}, task.record, {
    statusManual: 'CONCLUIDA',
    statusCalc: 'CONCLUIDA',
    dataConclusao: normalizeIsoDate_(payload.dataConclusao || new Date()),
    resultado: sanitizeText_(payload.resultado || task.record.resultado || ''),
    resultadoFinal: normalizeNumberString_(payload.resultadoFinal || task.record.resultadoFinal || ''),
    observacao: sanitizeText_(payload.observacao || task.record.observacao || ''),
    updatedAt: new Date().toISOString()
  });

  writeRecordAtRow_('TAREFAS', task.rowIndex, merged);
  appendTaskHistory_(task.record, merged, 'APP_CONCLUSAO', payload.observacao || 'Conclusão registrada via painel');
  logAction_('CONCLUIR', 'TAREFAS', payload.tarefaId, merged);
  invalidateInitialDataCache_();

  return buildMutationResponse_('Tarefa concluída com sucesso.', payload.tarefaId, 'TAREFAS');
}

function apiArchiveEmpresa(empresaId) {
  const found = findRecordById_('EMPRESAS', empresaId);
  if (!found) {
    throw new Error('Empresa não encontrada.');
  }

  const updated = Object.assign({}, found.record, {
    status: 'INATIVA',
    updatedAt: new Date().toISOString()
  });

  writeRecordAtRow_('EMPRESAS', found.rowIndex, updated);
  logAction_('ARQUIVAR', 'EMPRESAS', empresaId, updated);
  invalidateInitialDataCache_();

  return buildMutationResponse_('Empresa arquivada com sucesso.', empresaId, 'EMPRESAS');
}

function apiDeleteTarefa(tarefaId) {
  const config = ENTITY_CONFIG.TAREFAS;
  const sheet = getOrCreateSheet_('TAREFAS');
  const values = sheet.getDataRange().getValues();
  const headers = values[0] || config.headers;
  const idIndex = headers.indexOf(config.idField);

  for (var rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
    if (String(values[rowIndex][idIndex]) === String(tarefaId)) {
      sheet.deleteRow(rowIndex + 1);
      logAction_('DELETE', 'TAREFAS', tarefaId, { tarefaId: tarefaId });
      invalidateInitialDataCache_();
      return buildMutationResponse_('Tarefa removida com sucesso.', tarefaId, 'TAREFAS');
    }
  }

  throw new Error('Tarefa não encontrada para exclusão.');
}

function apiGetExecutiveReport(projetoId) {
  const data = getInitialAppData(true);
  const project = data.projetos.filter(function(item) {
    return String(item.projetoId) === String(projetoId);
  })[0];

  if (!project) {
    throw new Error('Projeto não encontrado para relatório.');
  }

  return buildExecutiveReportForProject_(project, data.empresas, data.tarefas, data.historico);
}

function buildMutationResponse_(message, entityId, entityName) {
  return {
    success: true,
    message: message,
    entityId: entityId,
    entityName: entityName,
    data: getInitialAppData(true)
  };
}

function upsertEntity_(entityName, payload, options) {
  const config = ENTITY_CONFIG[entityName];
  const sheet = getOrCreateSheet_(entityName);
  const headers = ensureSheetHeaders_(sheet, config.headers);
  const nowIso = new Date().toISOString();
  const cleanPayload = normalizePayloadForEntity_(entityName, payload || {});
  const recordId = cleanPayload[config.idField] || Utilities.getUuid();
  cleanPayload[config.idField] = recordId;
  cleanPayload.updatedAt = nowIso;
  if (!cleanPayload.createdAt) {
    cleanPayload.createdAt = nowIso;
  }

  const values = sheet.getDataRange().getValues();
  const idIndex = headers.indexOf(config.idField);
  const previous = values.slice(1).map(function(row) {
    return rowToObject_(headers, row);
  }).filter(function(record) {
    return String(record[config.idField]) === String(recordId);
  })[0];

  const finalRecord = mergeWithDefaults_(entityName, previous, cleanPayload);
  finalRecord.statusCalc = calculateStatusForEntity_(entityName, finalRecord);
  if (entityName === 'PROJETOS') {
    finalRecord.percentualConclusao = normalizeNumberString_(finalRecord.percentualConclusao);
    finalRecord.indicadorAtual = normalizeNumberString_(finalRecord.indicadorAtual);
    finalRecord.orcamento = normalizeNumberString_(finalRecord.orcamento);
    finalRecord.ultimaAtualizacao = nowIso;
  }
  if (entityName === 'TAREFAS' && finalRecord.statusCalc === 'CONCLUIDA' && !finalRecord.dataConclusao) {
    finalRecord.dataConclusao = normalizeIsoDate_(new Date());
  }

  var updated = false;
  for (var rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
    if (String(values[rowIndex][idIndex]) === String(recordId)) {
      sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([headers.map(function(header) {
        return finalRecord[header] !== undefined ? finalRecord[header] : '';
      })]);
      updated = true;
      break;
    }
  }

  if (!updated) {
    sheet.appendRow(headers.map(function(header) {
      return finalRecord[header] !== undefined ? finalRecord[header] : '';
    }));
  }

  if (entityName === 'TAREFAS' && options && options.trackHistory) {
    appendTaskHistory_(previous || {}, finalRecord, options.origin || 'APP_FORM', 'Alteração de tarefa');
  }

  logAction_(updated ? 'UPDATE' : 'CREATE', entityName, recordId, finalRecord);
  invalidateInitialDataCache_();

  return {
    success: true,
    id: recordId,
    action: updated ? 'update' : 'create',
    record: finalRecord
  };
}

function getEntityRecords_(entityName) {
  const config = ENTITY_CONFIG[entityName];
  const sheet = getOrCreateSheet_(entityName);
  const headers = ensureSheetHeaders_(sheet, config.headers);
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return [];
  }

  return values.slice(1).filter(function(row) {
    return row.some(function(cell) {
      return String(cell || '').trim() !== '';
    });
  }).map(function(row) {
    return normalizeRecordByEntity_(entityName, rowToObject_(headers, row));
  });
}

function hydrateProjects_(projects, tasks) {
  return projects.map(function(project) {
    const projectTasks = tasks.filter(function(task) {
      return String(task.projetoId) === String(project.projetoId);
    });
    const completed = projectTasks.filter(function(task) {
      return task.statusCalc === 'CONCLUIDA';
    }).length;
    const progress = projectTasks.length ? Math.round((completed / projectTasks.length) * 100) : Number(project.percentualConclusao || 0);
    const overdue = projectTasks.filter(function(task) {
      return task.statusCalc === 'ATRASADA';
    }).length;
    const currentPhase = projectTasks.sort(compareDateFieldDesc_).map(function(task) {
      return task.fase;
    })[0] || project.faseAtual || 'DIAGNOSTICO';

    return Object.assign({}, project, {
      percentualConclusao: progress,
      faseAtual: currentPhase,
      tarefasTotal: projectTasks.length,
      tarefasConcluidas: completed,
      tarefasAtrasadas: overdue,
      statusCalc: calculateProjectStatus_(project, projectTasks)
    });
  });
}

function hydrateTasks_(tasks) {
  return tasks.map(function(task) {
    const status = calculateTaskStatus_(task);
    return Object.assign({}, task, {
      statusCalc: status
    });
  }).sort(compareDateFieldAsc_);
}

function buildDashboardSummary_(empresas, projetos, tarefas) {
  const activeCompanies = empresas.filter(function(item) {
    return item.status !== 'INATIVA';
  }).length;
  const activeProjects = projetos.filter(function(item) {
    return item.statusCalc !== 'CONCLUIDA' && item.statusCalc !== 'ARQUIVADO';
  }).length;
  const tasksCompleted = tarefas.filter(function(item) {
    return item.statusCalc === 'CONCLUIDA';
  }).length;
  const overdueTasks = tarefas.filter(function(item) {
    return item.statusCalc === 'ATRASADA';
  }).length;
  const dueSoon = tarefas.filter(function(item) {
    return item.statusCalc === 'VENCENDO_SEMANA';
  });
  const completionRate = tarefas.length ? Math.round((tasksCompleted / tarefas.length) * 100) : 0;

  return {
    empresasAtivas: activeCompanies,
    projetosAtivos: activeProjects,
    tarefasPendentes: tarefas.length - tasksCompleted,
    tarefasAtrasadas: overdueTasks,
    taxaConclusao: completionRate,
    proximosVencimentos: dueSoon.sort(compareDateFieldAsc_).slice(0, 5),
    pipeline: summarizePipeline_(projetos),
    alertas: buildExecutiveAlerts_(projetos, tarefas)
  };
}

function buildExecutiveReportForProject_(project, empresas, tarefas, historico) {
  const company = empresas.filter(function(item) {
    return String(item.empresaId) === String(project.empresaId);
  })[0] || {};
  const projectTasks = tarefas.filter(function(task) {
    return String(task.projetoId) === String(project.projetoId);
  });
  const completedTasks = projectTasks.filter(function(task) {
    return task.statusCalc === 'CONCLUIDA';
  });
  const overdueTasks = projectTasks.filter(function(task) {
    return task.statusCalc === 'ATRASADA';
  });
  const history = historico.filter(function(item) {
    return projectTasks.some(function(task) {
      return String(task.tarefaId) === String(item.tarefaId);
    });
  }).sort(compareDateFieldDesc_).slice(0, 5);

  const highlights = [];
  if (completedTasks.length) {
    highlights.push(completedTasks.length + ' entregas concluídas no período.');
  }
  if (project.percentualConclusao) {
    highlights.push('Conclusão estimada em ' + project.percentualConclusao + '%.');
  }
  if (overdueTasks.length) {
    highlights.push(overdueTasks.length + ' ações críticas demandam atenção imediata.');
  } else {
    highlights.push('Nenhuma atividade crítica em atraso.');
  }

  const nextActions = projectTasks.filter(function(task) {
    return task.statusCalc !== 'CONCLUIDA';
  }).sort(compareDateFieldAsc_).slice(0, 3).map(function(task) {
    return {
      titulo: task.descricao,
      responsavel: task.responsavel || 'Consultoria',
      prazo: task.dataFim,
      status: task.statusCalc
    };
  });

  return {
    projetoId: project.projetoId,
    nomeProjeto: project.nomeProjeto,
    empresa: company.nomeFantasia || company.nomeRazao || 'Cliente não identificado',
    status: project.statusCalc,
    percentualConclusao: Number(project.percentualConclusao || 0),
    faseAtual: project.faseAtual || 'DIAGNOSTICO',
    escopo: project.escopo || 'Escopo não informado.',
    resumoExecutivo: [
      'Projeto ' + project.nomeProjeto + ' para ' + (company.nomeFantasia || company.nomeRazao || 'cliente') + '.',
      'Status atual: ' + project.statusCalc + '.',
      'Percentual de conclusão: ' + Number(project.percentualConclusao || 0) + '%.',
      'Próximo marco previsto até ' + (project.prazoFinal || 'sem prazo definido') + '.'
    ].join(' '),
    indicadores: {
      tarefasTotal: projectTasks.length,
      tarefasConcluidas: completedTasks.length,
      tarefasAtrasadas: overdueTasks.length,
      tarefasEmCurso: projectTasks.filter(function(task) {
        return task.statusCalc === 'EM_DIA' || task.statusCalc === 'VENCENDO_SEMANA';
      }).length
    },
    destaques: highlights,
    proximasAcoes: nextActions,
    historicoRecente: history
  };
}

function buildSheetSchema_() {
  return {
    EMPRESAS: ENTITY_CONFIG.EMPRESAS.headers,
    PROJETOS: ENTITY_CONFIG.PROJETOS.headers,
    TAREFAS: ENTITY_CONFIG.TAREFAS.headers,
    TAREFAS_HISTORICO: ENTITY_CONFIG.TAREFAS_HISTORICO.headers,
    LOGS: ENTITY_CONFIG.LOGS.headers
  };
}

function buildExecutiveAlerts_(projects, tasks) {
  const alerts = [];
  const overdueTasks = tasks.filter(function(task) {
    return task.statusCalc === 'ATRASADA';
  });
  const stalledProjects = projects.filter(function(project) {
    return Number(project.percentualConclusao || 0) < 30 && project.statusCalc === 'ATRASADA';
  });

  if (overdueTasks.length) {
    alerts.push({
      type: 'danger',
      title: 'Ações em atraso',
      message: overdueTasks.length + ' tarefa(s) passaram do prazo definido.'
    });
  }
  if (stalledProjects.length) {
    alerts.push({
      type: 'warning',
      title: 'Projetos com risco',
      message: stalledProjects.length + ' projeto(s) exigem revisão de plano de ação.'
    });
  }
  if (!alerts.length) {
    alerts.push({
      type: 'success',
      title: 'Operação saudável',
      message: 'Não existem alertas críticos neste momento.'
    });
  }

  return alerts;
}

function summarizePipeline_(projects) {
  return ['DIAGNOSTICO', 'PLANEJAMENTO', 'IMPLANTACAO', 'MONITORAMENTO', 'ENCERRAMENTO'].map(function(phase) {
    return {
      fase: phase,
      quantidade: projects.filter(function(project) {
        return String(project.faseAtual || '').toUpperCase() === phase;
      }).length
    };
  });
}

function appendTaskHistory_(beforeRecord, afterRecord, origin, reason) {
  const previousStatus = beforeRecord.statusCalc || beforeRecord.statusManual || '';
  const nextStatus = afterRecord.statusCalc || afterRecord.statusManual || '';
  if (!afterRecord.tarefaId || previousStatus === nextStatus && beforeRecord.tarefaId) {
    return;
  }

  const historyRecord = {
    historicoId: Utilities.getUuid(),
    tarefaId: afterRecord.tarefaId,
    statusDe: previousStatus,
    statusPara: nextStatus,
    alteradoEm: new Date().toISOString(),
    origem: origin,
    motivoConclusao: reason || '',
    resultadoFinalNumber: normalizeNumberString_(afterRecord.resultadoFinal || ''),
    resultadoFinalTexto: sanitizeText_(afterRecord.resultado || '')
  };

  appendRow_('TAREFAS_HISTORICO', historyRecord);
}

function appendRow_(entityName, record) {
  const config = ENTITY_CONFIG[entityName];
  const sheet = getOrCreateSheet_(entityName);
  const headers = ensureSheetHeaders_(sheet, config.headers);
  sheet.appendRow(headers.map(function(header) {
    return record[header] !== undefined ? record[header] : '';
  }));
}

function logAction_(action, entityName, entityId, details) {
  appendRow_('LOGS', {
    logId: Utilities.getUuid(),
    timestamp: new Date().toISOString(),
    acao: action,
    entidade: entityName,
    entidadeId: entityId,
    detalhes: JSON.stringify(details || {})
  });
}

function findRecordById_(entityName, recordId) {
  const config = ENTITY_CONFIG[entityName];
  const records = getEntityRecords_(entityName);
  const record = records.filter(function(item) {
    return String(item[config.idField]) === String(recordId);
  })[0];

  if (!record) {
    return null;
  }

  const sheet = getOrCreateSheet_(entityName);
  const values = sheet.getDataRange().getValues();
  const headers = values[0] || config.headers;
  const idIndex = headers.indexOf(config.idField);
  for (var rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
    if (String(values[rowIndex][idIndex]) === String(recordId)) {
      return { record: record, rowIndex: rowIndex + 1 };
    }
  }

  return null;
}

function writeRecordAtRow_(entityName, rowIndex, record) {
  const config = ENTITY_CONFIG[entityName];
  const sheet = getOrCreateSheet_(entityName);
  const headers = ensureSheetHeaders_(sheet, config.headers);
  const normalized = mergeWithDefaults_(entityName, null, record);
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([headers.map(function(header) {
    return normalized[header] !== undefined ? normalized[header] : '';
  })]);
}

function mergeWithDefaults_(entityName, previous, incoming) {
  const config = ENTITY_CONFIG[entityName];
  const base = {};
  config.headers.forEach(function(header) {
    base[header] = previous && previous[header] !== undefined ? previous[header] : '';
  });

  Object.keys(incoming || {}).forEach(function(key) {
    base[key] = incoming[key];
  });

  return normalizeRecordByEntity_(entityName, base);
}

function normalizePayloadForEntity_(entityName, payload) {
  const record = Object.assign({}, payload);
  if (entityName === 'EMPRESAS') {
    record.nomeRazao = sanitizeText_(record.nomeRazao);
    record.nomeFantasia = sanitizeText_(record.nomeFantasia);
    record.cnpj = sanitizeText_(record.cnpj);
    record.endereco = sanitizeText_(record.endereco);
    record.telefone = sanitizeText_(record.telefone);
    record.email = sanitizeText_(record.email);
    record.responsavel = sanitizeText_(record.responsavel);
    record.observacao = sanitizeText_(record.observacao);
    record.status = record.status || 'ATIVA';
    record.dataInicio = normalizeIsoDate_(record.dataInicio);
    record.dataFimPrevista = normalizeIsoDate_(record.dataFimPrevista);
  }
  if (entityName === 'PROJETOS') {
    record.nomeProjeto = sanitizeText_(record.nomeProjeto);
    record.escopo = sanitizeText_(record.escopo);
    record.faseAtual = record.faseAtual || 'DIAGNOSTICO';
    record.consultorResponsavel = sanitizeText_(record.consultorResponsavel);
    record.prioridade = record.prioridade || 'MEDIA';
    record.statusManual = record.statusManual || 'EM_ANDAMENTO';
    record.dataInicio = normalizeIsoDate_(record.dataInicio);
    record.prazoFinal = normalizeIsoDate_(record.prazoFinal);
    record.orcamento = normalizeNumberString_(record.orcamento);
    record.indicadorMeta = sanitizeText_(record.indicadorMeta);
    record.indicadorAtual = normalizeNumberString_(record.indicadorAtual);
    record.percentualConclusao = normalizeNumberString_(record.percentualConclusao);
    record.ultimaAtualizacao = new Date().toISOString();
  }
  if (entityName === 'TAREFAS') {
    record.fase = record.fase || 'DIAGNOSTICO';
    record.descricao = sanitizeText_(record.descricao);
    record.responsavel = sanitizeText_(record.responsavel);
    record.statusManual = record.statusManual || 'BACKLOG';
    record.prioridade = record.prioridade || 'MEDIA';
    record.dataInicio = normalizeIsoDate_(record.dataInicio);
    record.dataFim = normalizeIsoDate_(record.dataFim);
    record.ordem = normalizeNumberString_(record.ordem || 0);
    record.semanaReferencia = sanitizeText_(record.semanaReferencia);
    record.dataConclusao = normalizeIsoDate_(record.dataConclusao);
    record.detalhes = sanitizeText_(record.detalhes);
    record.resultado = sanitizeText_(record.resultado);
    record.observacao = sanitizeText_(record.observacao);
    record.metaTipo = sanitizeText_(record.metaTipo || 'OPERACIONAL');
    record.metaValor = normalizeNumberString_(record.metaValor);
    record.metaUnidade = sanitizeText_(record.metaUnidade || '%');
    record.resultadoFinal = normalizeNumberString_(record.resultadoFinal);
  }

  return record;
}

function normalizeRecordByEntity_(entityName, record) {
  const config = ENTITY_CONFIG[entityName];
  const normalized = {};
  config.headers.forEach(function(header) {
    normalized[header] = normalizeCellValue_(record[header]);
  });

  if (entityName === 'TAREFAS') {
    normalized.statusCalc = calculateTaskStatus_(normalized);
  }
  if (entityName === 'PROJETOS') {
    normalized.statusCalc = normalized.statusCalc || calculateProjectStatus_(normalized, []);
  }

  return normalized;
}

function normalizeCellValue_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return value.toISOString();
  }
  if (typeof value === 'number') {
    return value;
  }
  return value === null || value === undefined ? '' : String(value).trim();
}

function calculateStatusForEntity_(entityName, record) {
  if (entityName === 'TAREFAS') {
    return calculateTaskStatus_(record);
  }
  if (entityName === 'PROJETOS') {
    return calculateProjectStatus_(record, []);
  }
  return sanitizeText_(record.status || 'ATIVA');
}

function calculateTaskStatus_(task) {
  const manual = String(task.statusManual || '').toUpperCase();
  if (manual === 'CONCLUIDA') {
    return 'CONCLUIDA';
  }

  const endDate = parseDateSafe_(task.dataFim);
  if (!endDate) {
    return manual || 'BACKLOG';
  }

  const today = truncateToDay_(new Date());
  const due = truncateToDay_(endDate);
  const diffDays = Math.round((due.getTime() - today.getTime()) / 86400000);

  if (diffDays < 0) {
    return 'ATRASADA';
  }
  if (diffDays <= 7) {
    return 'VENCENDO_SEMANA';
  }
  return manual === 'BACKLOG' ? 'BACKLOG' : 'EM_DIA';
}

function calculateProjectStatus_(project, tasks) {
  const manual = String(project.statusManual || '').toUpperCase();
  if (manual === 'CONCLUIDO' || manual === 'CONCLUIDA') {
    return 'CONCLUIDA';
  }
  if (manual === 'ARQUIVADO') {
    return 'ARQUIVADO';
  }

  if (tasks && tasks.length) {
    const allCompleted = tasks.every(function(task) {
      return task.statusCalc === 'CONCLUIDA';
    });
    if (allCompleted) {
      return 'CONCLUIDA';
    }
    const overdue = tasks.some(function(task) {
      return task.statusCalc === 'ATRASADA';
    });
    if (overdue) {
      return 'ATRASADA';
    }
    const nearDue = tasks.some(function(task) {
      return task.statusCalc === 'VENCENDO_SEMANA';
    });
    if (nearDue) {
      return 'VENCENDO_SEMANA';
    }
    return 'EM_DIA';
  }

  const deadline = parseDateSafe_(project.prazoFinal);
  if (deadline && truncateToDay_(deadline).getTime() < truncateToDay_(new Date()).getTime()) {
    return 'ATRASADA';
  }
  return 'EM_DIA';
}

function getOrCreateSheet_(sheetName) {
  const spreadsheet = SpreadsheetApp.openById(DB_ID);
  const existing = spreadsheet.getSheetByName(sheetName);
  if (existing) {
    return existing;
  }

  const created = spreadsheet.insertSheet(sheetName);
  const config = ENTITY_CONFIG[sheetName];
  if (config) {
    created.getRange(1, 1, 1, config.headers.length).setValues([config.headers]);
  }
  return created;
}

function ensureSheetHeaders_(sheet, expectedHeaders) {
  const lastColumn = Math.max(sheet.getLastColumn(), expectedHeaders.length);
  const range = sheet.getRange(1, 1, 1, lastColumn);
  const currentHeaders = range.getValues()[0];
  const hasHeader = currentHeaders.some(function(cell) {
    return String(cell || '').trim() !== '';
  });

  if (!hasHeader) {
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    return expectedHeaders.slice();
  }

  var changed = false;
  expectedHeaders.forEach(function(header, index) {
    if (!currentHeaders[index]) {
      currentHeaders[index] = header;
      changed = true;
    }
  });

  if (changed) {
    sheet.getRange(1, 1, 1, currentHeaders.length).setValues([currentHeaders]);
  }

  return currentHeaders.slice(0, Math.max(currentHeaders.length, expectedHeaders.length));
}

function rowToObject_(headers, row) {
  const record = {};
  headers.forEach(function(header, index) {
    if (!header) {
      return;
    }
    record[String(header).trim()] = row[index];
  });
  return record;
}

function safeGetUserEmail_() {
  try {
    return Session.getActiveUser().getEmail() || 'consultor@eve.local';
  } catch (error) {
    return 'consultor@eve.local';
  }
}

function normalizeIsoDate_(value) {
  if (!value) {
    return '';
  }
  const parsed = parseDateSafe_(value);
  return parsed ? Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(value);
}

function normalizeNumberString_(value) {
  if (value === '' || value === null || value === undefined) {
    return '';
  }
  if (typeof value === 'number') {
    return value;
  }

  const raw = String(value).trim();
  if (!raw) {
    return '';
  }

  const hasComma = raw.indexOf(',') > -1;
  const hasDot = raw.indexOf('.') > -1;
  var normalizedText = raw;

  if (hasComma && hasDot) {
    normalizedText = raw.replace(/\./g, '').replace(',', '.');
  } else if (hasComma) {
    normalizedText = raw.replace(',', '.');
  }

  const normalized = Number(normalizedText);
  return isNaN(normalized) ? raw : normalized;
}

function sanitizeText_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function parseDateSafe_(value) {
  if (!value) {
    return null;
  }
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return value;
  }
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function truncateToDay_(dateValue) {
  return new Date(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate());
}

function compareDateFieldAsc_(left, right) {
  const leftDate = parseDateSafe_(left.dataFim || left.prazoFinal || left.alteradoEm || left.updatedAt) || new Date('2999-12-31');
  const rightDate = parseDateSafe_(right.dataFim || right.prazoFinal || right.alteradoEm || right.updatedAt) || new Date('2999-12-31');
  return leftDate.getTime() - rightDate.getTime();
}

function compareDateFieldDesc_(left, right) {
  const leftDate = parseDateSafe_(left.alteradoEm || left.updatedAt || left.dataFim || left.prazoFinal) || new Date('1970-01-01');
  const rightDate = parseDateSafe_(right.alteradoEm || right.updatedAt || right.dataFim || right.prazoFinal) || new Date('1970-01-01');
  return rightDate.getTime() - leftDate.getTime();
}

function invalidateInitialDataCache_() {
  CacheService.getScriptCache().remove(CACHE_KEY_INITIAL_DATA);
}
