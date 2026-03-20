Gestão EVE - Sistema Apollo Enterprise (SAE)

Este documento detalha a arquitetura, estrutura de dados e funcionalidades do ERP Gestão EVE, desenvolvido sob o padrão SAE (Sistema Apollo Enterprise) utilizando a stack Google Apps Script (GAS) e Vue 3.

🚀 Arquitetura Técnica

Backend: Google Apps Script (V8 Runtime).

Frontend: Vue 3 (via CDN) em modo Single File Component (SFC).

Banco de Dados: Google Sheets (Abas atuando como tabelas relacionais).

Design System: SAE Glassmorphism Dark (Mobile-first).

📊 Estrutura do Banco de Dados (Google Sheets)

O banco de dados é composto por abas específicas, onde cada linha representa um registro e a primeira coluna é sempre o identificador único (ID).

1. Aba: EMPRESAS

Armazena o cadastro de clientes e parceiros.

Colunas: empresaId, nomeRazao, nomeFantasia, cnpj, endereco, telefone, observacao, status, dataInicio, dataFimPrevista.

Padrão de ID: E-0000.

2. Aba: PROJETOS

Agrupa as atividades por escopo de projeto.

Colunas: projetoId, empresaId, nomeProjeto, faseAtual, dataInicio, prazoFinal, statusCalc.

Padrão de ID: P-0000.

3. Aba: TAREFAS

O núcleo operacional do sistema, onde as ações são controladas.

Colunas: tarefaId, projetoId, fase, descricao, statusManual, dataInicio, dataFim, ordem, semanaReferencia, statusCalc, dataConclusao, detalhes, resultado, observacao, metaTipo, metaValor, metaUnidade.

Padrão de ID: T-0000.

4. Aba: TAREFAS_HISTORICO

Log de mudanças de status para auditoria e linha do tempo.

Colunas: historicoId, tarefaId, statusDe, statusPara, alteradoEm, origem, motivoConclusao, resultadoFinalNumber, resultadoFinalTexto.

5. Aba: PARAMETROS

Controle de variáveis de sistema e contadores de ID.

Estrutura: Chave/Valor.

Chaves Críticas: NEXT_ID_E, NEXT_ID_P, NEXT_ID_T, SYSTEM_VERSION.

6. Aba: LOGS

Registro de todas as operações de escrita (Create/Update).

Colunas: timestamp, acao, entidade, entidadeId, detalhes (JSON).

🛠️ Elementos Implementados (codigo.gs)

O arquivo codigo.gs gerencia a inteligência do negócio e a ponte com o Sheets:

Roteador (doGet): Serve o WebApp configurando viewport e segurança.

Motor de Persistência (saveRecord): - Lógica inteligente de Upsert (identifica se deve criar ou atualizar).

Atribuição automática de IDs sequenciais baseada em PARAMETROS.

Registro automático de Logs de auditoria.

Normalizador de Dados (getTableData): - Converte abas em objetos JSON para o Vue.

Tratamento automático de objetos de Data para strings ISO.

Gerador de IDs (generateSAEId): Garante a unicidade seguindo os prefixos E, P e T.

API de Inicialização (getInitialAppData): Carrega todo o estado do sistema em uma única requisição para performance (Single Request Pattern).

🎨 Interface e UI (Padrão SAE)

Glassmorphism: Uso de fundos translúcidos e desfoque (backdrop-filter).

Responsive Design: Foco em uso mobile sem sacrificar a visão desktop.

Feedback: Logs em tempo real e skeletons para carregamento.

📝 Notas de Versão

Versão Atual: 1.0.0-SAE

Status: Backend Core finalizado e integrado ao Banco de Dados.
