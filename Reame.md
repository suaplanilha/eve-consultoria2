# EVE Consultoria | SAE SaaS

WebApp de consultoria organizacional construído em **Google Apps Script + Google Sheets + Vue 3 via CDN**, com foco em operação real usando a planilha existente como banco de dados. O sistema foi remodelado para conectar o frontend ao backend do Apps Script com **`google.script.run` + Promises**, mantendo a proposta de um SaaS enxuto, moderno e pronto para evolução.

## Estrutura das abas no Google Sheets

> Padrão adotado: **1 aba = 1 entidade**, datas em **ISO 8601 / `yyyy-MM-dd`**, chave primária em **UUID** para novos registros, preservando compatibilidade com dados legados já existentes.

### 1) Aba `EMPRESAS`

Colunas:

- `empresaId`
- `nomeRazao`
- `nomeFantasia`
- `cnpj`
- `endereco`
- `telefone`
- `email`
- `responsavel`
- `observacao`
- `status`
- `dataInicio`
- `dataFimPrevista`
- `createdAt`
- `updatedAt`

### 2) Aba `PROJETOS`

Colunas:

- `projetoId`
- `empresaId`
- `nomeProjeto`
- `escopo`
- `faseAtual`
- `consultorResponsavel`
- `prioridade`
- `statusManual`
- `statusCalc`
- `dataInicio`
- `prazoFinal`
- `orcamento`
- `indicadorMeta`
- `indicadorAtual`
- `percentualConclusao`
- `ultimaAtualizacao`
- `createdAt`
- `updatedAt`

### 3) Aba `TAREFAS`

Colunas:

- `tarefaId`
- `projetoId`
- `fase`
- `descricao`
- `responsavel`
- `statusManual`
- `statusCalc`
- `prioridade`
- `dataInicio`
- `dataFim`
- `ordem`
- `semanaReferencia`
- `dataConclusao`
- `detalhes`
- `resultado`
- `observacao`
- `metaTipo`
- `metaValor`
- `metaUnidade`
- `resultadoFinal`
- `createdAt`
- `updatedAt`

### 4) Aba `TAREFAS_HISTORICO`

Colunas:

- `historicoId`
- `tarefaId`
- `statusDe`
- `statusPara`
- `alteradoEm`
- `origem`
- `motivoConclusao`
- `resultadoFinalNumber`
- `resultadoFinalTexto`

### 5) Aba `LOGS`

Colunas:

- `logId`
- `timestamp`
- `acao`
- `entidade`
- `entidadeId`
- `detalhes`

## Backend (`code.gs`)

O backend foi reorganizado para separar claramente:

- **Entrega do HTML** (`doGet`)
- **Leitura e normalização de entidades** (`getEntityRecords_`, `normalizeRecordByEntity_`)
- **CRUD genérico e persistência** (`upsertEntity_`, `writeRecordAtRow_`, `appendRow_`)
- **Regras de negócio** (`calculateTaskStatus_`, `calculateProjectStatus_`, `buildDashboardSummary_`)
- **Relatórios executivos** (`buildExecutiveReportForProject_`, `apiGetExecutiveReport`)
- **Auditoria e histórico** (`logAction_`, `appendTaskHistory_`)
- **Performance** com `CacheService` para a carga inicial

### APIs expostas ao frontend

- `getInitialAppData(forceRefresh)`
- `apiSaveEmpresa(payload)`
- `apiSaveProjeto(payload)`
- `apiSaveTarefa(payload)`
- `apiConcluirTarefa(payload)`
- `apiArchiveEmpresa(empresaId)`
- `apiDeleteTarefa(tarefaId)`
- `apiGetExecutiveReport(projetoId)`

## Frontend (`index.html`)

Arquivo único com:

- **Vue 3 via CDN**
- **Tailwind via CDN** apenas para utilitários leves
- **CSS customizado no mesmo arquivo**
- **Layout Glassmorphism Dark mobile-first**
- **Sidebar retrátil em telas menores**
- **Skeleton loaders** durante a carga inicial
- **Kanban de tarefas** por status
- **Formulários modais completos** para empresas, projetos e tarefas
- **Conclusão de tarefas com histórico**
- **Relatório executivo por projeto**

## Fluxo operacional suportado

1. Carregar empresas, projetos, tarefas, histórico e indicadores em uma única chamada inicial.
2. Cadastrar e editar empresas conectadas à planilha real.
3. Cadastrar e editar projetos vinculados a empresas.
4. Cadastrar, editar, concluir e excluir tarefas.
5. Registrar histórico de status e log de auditoria.
6. Emitir relatórios executivos por projeto a partir dos dados persistidos.

## Observações de performance

Como o projeto usa Google Apps Script + Sheets:

- A carga inicial usa **cache de 3 minutos** com `CacheService`.
- As alterações invalidam o cache para evitar inconsistência.
- O modelo está preparado para crescimento, mas em bases maiores o próximo passo recomendado é:
  - leitura em lote por intervalo específico,
  - agregações pré-computadas,
  - ou estratégias de paginação e cache por entidade.
