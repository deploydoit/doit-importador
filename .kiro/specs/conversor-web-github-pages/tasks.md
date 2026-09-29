# Implementation Plan: Conversor Web (GitHub Pages)

## Overview

Este plano porta o Conversor de Dados DOit (hoje Streamlit/Python) para um site 100% estático executado no navegador, publicável no GitHub Pages. A stack é **Vite + TypeScript** (build estático com `base` para subcaminho), **SheetJS** para leitura/escrita de `.xlsx`/`.xls`/`.csv`, e **Vitest + fast-check** para testes unitários e de propriedade sobre o núcleo puro.

A estratégia é construir primeiro a base (projeto, tipos, testes, assets), depois portar cada componente do núcleo puro (funções sem I/O/DOM) com seus testes de propriedade próximos à implementação, em seguida montar a orquestração da UI e, por fim, o workflow de deploy. Cada uma das 42 propriedades de correção do design vira um teste de propriedade dedicado (`fast-check`, mínimo 100 iterações), anotado com o número da propriedade e a cláusula de requisito que valida.

Convenção de tag de propriedade (em cada teste de propriedade):
`// Feature: conversor-web-github-pages, Property {número}: {texto da propriedade}`

## Tasks

- [x] 1. Configurar projeto, build estático e tipos base
  - [x] 1.1 Inicializar projeto Vite + TypeScript com build estático e base path configurável
    - Criar estrutura `src/`, `src/core/`, `src/ui/`, `public/assets/`
    - Configurar `vite.config.ts` com `base` derivado do subcaminho do repositório (`/<repositorio>/`)
    - Garantir saída de build apenas HTML/CSS/JS estático e adicionar SheetJS (`xlsx`) como dependência
    - _Requirements: 1.2, 1.5, 2.1, 2.5_
  - [x] 1.2 Configurar Vitest + fast-check e geradores customizados de teste
    - Adicionar Vitest e fast-check; configurar `numRuns: 100` como padrão dos testes de propriedade
    - Criar `src/core/__tests__/generators.ts` com `arbTable`, `arbColumnName`, `arbAcentText`, `arbMojibake`, `arbPhoneDigits`, `arbBpoSheet`, `arbNavisSheet`, `arbReference`
    - _Requirements: 1.5_
  - [x] 1.3 Definir tipos base do núcleo (`src/core/types.ts`)
    - Definir `Cell`, `Table`, `DataType`, `SourceSystem`, `CaseStyle`, `Alert`, `StepResult<T>` e a classe `ConversionError` com `code` estável
    - _Requirements: 4.1, 5.1, 14.1_

- [x] 2. Empacotar modelos e listas de padronização como assets estáticos
  - [x] 2.1 Copiar `models/*.xlsx` e `mappings/*.txt` para `public/assets/` e implementar carregadores
    - Copiar modelos para `assets/models/` e listas para `assets/mappings/`
    - Implementar `loadModelColumns(type)` e `loadStandardLists()` usando `fetch(import.meta.env.BASE_URL + 'assets/...')` + SheetJS
    - Lançar `ConversionError` com `MODEL_LOAD_FAILED` / `MAPPINGS_LOAD_FAILED` quando um recurso não puder ser carregado
    - _Requirements: 11.6, 11.7, 11.8, 4.6, 2.5_
  - [x]* 2.2 Escrever testes de integração de carregamento de assets
    - Verificar que cada modelo por tipo carrega as colunas e que as listas de padronização carregam as entradas esperadas
    - _Requirements: 11.6, 11.7, 4.3_

- [x] 3. Implementar Leitor_Planilha (`src/core/reader.ts`)
  - [x] 3.1 Implementar leitura de workbook, validação de extensão e CSV com fallback de encoding
    - Implementar `isSupportedExtension` (case-insensitive), `readWorkbook` (abas + `readSheet`), `readCsv` (UTF-8 → Latin-1)
    - Validar tamanho ≤ 50 MB e rejeitar arquivo vazio; lançar `FILE_TOO_LARGE`, `UNSUPPORTED_FORMAT`, `EMPTY_FILE`, `CSV_ENCODING`
    - _Requirements: 1.6, 3.1, 3.2, 3.3, 3.4, 3.5, 3.6, 3.7, 3.8, 16.4_
  - [x]* 3.2 Escrever teste de propriedade para aceitação de extensões
    - **Property 3: Extensões de arquivo são aceitas de forma insensível à caixa**
    - **Validates: Requirements 3.1, 3.5**
  - [x]* 3.3 Escrever teste de propriedade para listagem/seleção de abas
    - **Property 4: Seleção de aba lista todas as abas e usa a primeira como padrão**
    - **Validates: Requirements 3.4**
  - [x]* 3.4 Escrever teste de propriedade para fallback Latin-1 no CSV
    - **Property 5: Leitura de CSV recupera texto Latin-1 quando UTF-8 falha**
    - **Validates: Requirements 3.7**
  - [x]* 3.5 Escrever testes unitários das condições de erro do leitor
    - Testar `FILE_TOO_LARGE` (>50 MB), `UNSUPPORTED_FORMAT`, `EMPTY_FILE`, `CSV_ENCODING`
    - _Requirements: 1.6, 3.3, 3.5, 3.6, 3.8_

- [x] 4. Implementar Corretor_Encoding (`src/core/encoding.ts`)
  - [x] 4.1 Implementar correção de mojibake em textos e tabela
    - Implementar `fixMojibake` (reinterpretação Latin-1 e Windows-1252) e `fixTableEncoding` (colunas + valores)
    - Preservar byte a byte textos sem marcador; preservar e sinalizar valores não resolvíveis
    - _Requirements: 6.1, 6.2, 6.3, 6.4, 6.5, 6.6_
  - [x]* 4.2 Escrever teste de propriedade para recuperação de acentuação
    - **Property 1: Correção de mojibake recupera acentuação e elimina marcadores**
    - **Validates: Requirements 6.1, 6.2, 6.3, 6.5**
  - [x]* 4.3 Escrever teste de propriedade para preservação de textos sem mojibake
    - **Property 2: Textos sem marcador de mojibake são preservados**
    - **Validates: Requirements 6.4**

- [x] 5. Checkpoint - Ensure all tests pass
  - Ensure all tests pass, ask the user if questions arise.

- [x] 6. Implementar Detector_Cabecalho (`src/core/header.ts`)
  - [x] 6.1 Implementar seleção de aba por tipo, detecção de cabeçalho e descarte de colunas inválidas
    - Implementar `selectSheetByType`, `detectHeaderRow` (20 linhas, ≥2 correspondências, case/trim-insensível) e `dropInvalidColumns`
    - _Requirements: 10.1, 10.2, 10.3, 10.4, 10.5_
  - [x]* 6.2 Escrever teste de propriedade para detecção da linha de cabeçalho
    - **Property 10: Detecção de cabeçalho seleciona a linha de maior correspondência**
    - **Validates: Requirements 10.1, 7.5**
  - [x]* 6.3 Escrever teste de propriedade para seleção de aba por palavra-chave
    - **Property 11: Seleção de aba por palavra-chave do tipo**
    - **Validates: Requirements 10.3**
  - [x]* 6.4 Escrever teste de propriedade para descarte de colunas inválidas
    - **Property 12: Descarte de colunas inválidas na detecção de cabeçalho**
    - **Validates: Requirements 10.5**

- [x] 7. Implementar Parser_Navis (`src/core/parsers/navis.ts`)
  - [x] 7.1 Implementar detecção de tipo, detecção de cabeçalho e extração de registros
    - Implementar `detectReportType` (15 linhas), `detectHeaderRow` (20 linhas, maior score) e `parseNavis` por tipo
    - Consolidar registros multilinha (ficha/previsão); retorno vazio tratado; lançar `NAVIS_UNKNOWN_TYPE` e `NAVIS_PARSE_FAILED`
    - _Requirements: 7.1, 7.2, 7.3, 7.4, 7.5, 7.6, 7.7, 7.9_
  - [x]* 7.2 Escrever teste de propriedade para uso do tipo de relatório escolhido
    - **Property 13: O parser Navis usa o tipo de relatório escolhido**
    - **Validates: Requirements 7.3**
  - [x]* 7.3 Escrever testes unitários de classificação por amostras Navis
    - Testar detecção por palavras-chave, os 8 tipos de extração e consolidação multilinha em amostras
    - _Requirements: 7.1, 7.4, 7.6_

- [x] 8. Implementar Parser_BPO (`src/core/parsers/bpo.ts`)
  - [x] 8.1 Implementar validação de entrada, detecção de meses e extração de lançamentos
    - Implementar `detectMonthColumns` e `parseFinanceiroHorizontal` (campos data/vencimento/descrição/valor/tipo/conciliado/categorias)
    - Validar aba/ano (1900..2100); omitir valores vazios/nulos/zero; compor data por ano+mês; lançar `BPO_INVALID_INPUT`, `BPO_NO_MONTHS`, `BPO_PARSE_FAILED`
    - _Requirements: 8.1, 8.2, 8.3, 8.4, 8.5, 8.6, 8.7, 8.8_
  - [x]* 8.2 Escrever teste de propriedade para validação do ano de referência
    - **Property 14: Validação do ano de referência do BPO**
    - **Validates: Requirements 8.2**
  - [x]* 8.3 Escrever teste de propriedade para detecção de colunas de meses
    - **Property 15: Detecção de colunas de meses do BPO**
    - **Validates: Requirements 8.3**
  - [x]* 8.4 Escrever teste de propriedade para campos requeridos dos lançamentos
    - **Property 16: Lançamentos do BPO têm todos os campos requeridos**
    - **Validates: Requirements 8.5**
  - [x]* 8.5 Escrever teste de propriedade para omissão de lançamentos zero/vazio
    - **Property 17: Lançamentos com valor vazio, nulo ou zero são omitidos**
    - **Validates: Requirements 8.6**
  - [x]* 8.6 Escrever teste de propriedade para composição da data por ano+mês
    - **Property 18: Composição da data a partir do ano informado e do mês da coluna**
    - **Validates: Requirements 8.7**

- [x] 9. Implementar Processador_Outlook (`src/core/processors/outlook.ts`)
  - [x] 9.1 Implementar concatenação de nome, distribuição de telefones e consolidação de anotações
    - Implementar `concatName`, `distributePhones`, `consolidateExtrasToNotes`, `processOutlook` (critérios 1..6, idêntico para CSV/XLSX)
    - Remover linhas com nome vazio; excedentes de telefone → Fax Comercial separados por `, `
    - _Requirements: 9.1, 9.2, 9.3, 9.4, 9.5, 9.6, 9.7_
  - [x]* 9.2 Escrever teste de propriedade para concatenação de nome
    - **Property 19: Concatenação de nome do Outlook preserva ordem sem espaços supérfluos**
    - **Validates: Requirements 9.1, 9.2**
  - [x]* 9.3 Escrever teste de propriedade para distribuição de telefones
    - **Property 20: Distribuição de telefones do Outlook preenche principais e transborda para Fax**
    - **Validates: Requirements 9.3, 9.4**
  - [x]* 9.4 Escrever teste de propriedade para remoção de linhas com nome vazio
    - **Property 23: Linhas com nome vazio são removidas (Outlook)**
    - **Validates: Requirements 9.6**
  - [x]* 9.5 Escrever teste de propriedade para equivalência CSV/XLSX
    - **Property 24: Equivalência entre entradas CSV e XLSX do Outlook**
    - **Validates: Requirements 9.7**

- [x] 10. Checkpoint - Ensure all tests pass
  - Ensure all tests pass, ask the user if questions arise.

- [x] 11. Implementar Mapeador_Colunas (`src/core/mapper.ts`)
  - [x] 11.1 Implementar mapeamento por regra pré-definida e por similaridade
    - Implementar `getPredefinedRule`, `similarity` (0..1 normalizado), `mapColumns` (regra → similaridade ≥ 80%)
    - Resolver empates por maior grau e ordem de ocorrência; mapeamento injetivo; colunas sem correspondência → string vazia; sinalizar não mapeadas
    - Lançar `RULE_UNAVAILABLE` quando as regras da origem não existirem
    - _Requirements: 5.4, 5.5, 5.6, 11.1, 11.2, 11.3, 11.4, 11.5_
  - [x]* 11.2 Escrever teste de propriedade para mapeamento por similaridade
    - **Property 6: Mapeamento por similaridade respeita o limiar e o desempate**
    - **Validates: Requirements 5.4, 5.5, 11.2, 11.3**
  - [x]* 11.3 Escrever teste de propriedade para regra pré-definida
    - **Property 7: Regra pré-definida mapeia para a coluna candidata existente**
    - **Validates: Requirements 11.1**
  - [x]* 11.4 Escrever teste de propriedade para injetividade do mapeamento
    - **Property 8: O mapeamento é injetivo nas colunas de origem**
    - **Validates: Requirements 11.4**
  - [x]* 11.5 Escrever teste de propriedade para colunas de destino vazias
    - **Property 9: Colunas de destino sem correspondência recebem valor vazio**
    - **Validates: Requirements 11.5**

- [x] 12. Implementar Consolidador_NaoMapeados (`src/core/unmapped.ts`)
  - [x] 12.1 Implementar consolidação de campos não mapeados por tipo
    - Implementar `detectCellType` e `consolidateUnmapped`: Cadastros → `ANOTAÇÕES`; Financeiro/Horas → `DESCRIÇÃO` (quando existe); Projetos → colunas customizáveis por tipo
    - Omitir colunas sem valores não vazios; delimitador único e consistente no formato "nome: valor"
    - _Requirements: 12.1, 12.2, 12.3, 12.4, 12.5, 12.6_
  - [x]* 12.2 Escrever teste de propriedade para consolidação de valores preenchidos
    - **Property 21: Consolidação de campos não mapeados inclui apenas valores preenchidos**
    - **Validates: Requirements 9.5, 12.1, 12.2, 12.5**
  - [x]* 12.3 Escrever teste de propriedade para distribuição em colunas customizáveis (Projetos)
    - **Property 22: Distribuição de não mapeados em colunas customizáveis por tipo (Projetos)**
    - **Validates: Requirements 12.3**

- [x] 13. Implementar Vinculador_IDs (`src/core/linker.ts`)
  - [x] 13.1 Implementar vínculo de IDs por nome normalizado
    - Implementar `normalizeName`, `buildIndex`, `linkIds` para Projetos/Financeiro/Contatos Relacionados
    - Nome ausente → ID vazio + alerta; ambíguo → primeiro registro + alerta; sem referência → aviso; lançar `REFERENCE_MISSING_COLUMNS`
    - _Requirements: 13.1, 13.2, 13.3, 13.4, 13.5, 13.6, 13.7, 13.8, 13.9, 13.10_
  - [x]* 13.2 Escrever teste de propriedade para vínculo por nome normalizado
    - **Property 25: Vínculo de IDs por nome normalizado**
    - **Validates: Requirements 13.1, 13.2, 13.3, 13.4, 13.5, 13.8**
  - [x]* 13.3 Escrever teste de propriedade para nomes ausentes
    - **Property 26: Nomes ausentes na referência resultam em ID vazio e alerta**
    - **Validates: Requirements 13.6**
  - [x]* 13.4 Escrever teste de propriedade para nomes ambíguos
    - **Property 27: Nome ambíguo usa o primeiro registro e gera alerta**
    - **Validates: Requirements 13.9**

- [x] 14. Implementar Formatador_Texto (`src/core/formatter.ts`)
  - [x] 14.1 Implementar estilos de caixa e formatações de campo
    - Implementar `applyCase` (preposições/artigos), `formatPhone`, `formatCep`, `formatDate`, `formatEmail`, `formatState`, `formatTable`
    - Preservar e sinalizar valores não formatáveis (telefone/CEP/data); preservar campos vazios/nulos
    - _Requirements: 14.2, 14.3, 14.4, 14.5, 14.6, 14.7, 14.8, 14.9, 14.10, 14.11, 14.12, 14.13, 14.14_
  - [x]* 14.2 Escrever teste de propriedade para estilo Primeira Maiúscula
    - **Property 28: Estilo Primeira Maiúscula capitaliza palavras preservando preposições e artigos**
    - **Validates: Requirements 14.2**
  - [x]* 14.3 Escrever teste de propriedade para estilo TUDO MAIÚSCULO
    - **Property 29: Estilo TUDO MAIÚSCULO equivale à conversão integral para maiúsculas**
    - **Validates: Requirements 14.3**
  - [x]* 14.4 Escrever teste de propriedade para estilo Original
    - **Property 30: Estilo Original preserva o texto**
    - **Validates: Requirements 14.4**
  - [x]* 14.5 Escrever teste de propriedade para telefone de 11 dígitos
    - **Property 31: Formatação de telefone de 11 dígitos**
    - **Validates: Requirements 14.6**
  - [x]* 14.6 Escrever teste de propriedade para telefone de 10 dígitos
    - **Property 32: Formatação de telefone de 10 dígitos**
    - **Validates: Requirements 14.7**
  - [x]* 14.7 Escrever teste de propriedade para CEP de 8 dígitos
    - **Property 33: Formatação de CEP de 8 dígitos**
    - **Validates: Requirements 14.9**
  - [x]* 14.8 Escrever teste de propriedade para data reconhecível
    - **Property 34: Formatação de data reconhecível**
    - **Validates: Requirements 14.11**
  - [x]* 14.9 Escrever teste de propriedade para e-mail
    - **Property 35: Formatação de e-mail**
    - **Validates: Requirements 14.13**
  - [x]* 14.10 Escrever teste de propriedade para sigla de estado
    - **Property 36: Formatação de sigla de estado**
    - **Validates: Requirements 14.14**

- [x] 15. Checkpoint - Ensure all tests pass
  - Ensure all tests pass, ask the user if questions arise.

- [x] 16. Implementar Removedor_Linhas e Atribuidor_IDs (`src/core/rows.ts`, `src/core/ids.ts`)
  - [x] 16.1 Implementar remoção de linhas de exemplo e sem identificador (`rows.ts`)
    - Implementar `removeExampleRows` (DOit Coleta) e `removeRowsWithoutIdentifier`; interromper com `EMPTY_FILE` quando não restarem registros
    - _Requirements: 16.1, 16.2, 16.3, 16.4_
  - [x]* 16.2 Escrever teste de propriedade para remoção de linhas sem identificador
    - **Property 39: Remoção de linhas sem identificador preserva as demais**
    - **Validates: Requirements 16.2, 16.3**
  - [x] 16.3 Implementar atribuição e validação de ID inicial (`ids.ts`)
    - Implementar `assignIds` (sequencial a partir do inicial), `validateStartId` (inteiro 1..999.999.999) e padrões por tipo; lançar/sinalizar `INVALID_START_ID` mantendo último válido
    - _Requirements: 15.1, 15.2, 15.3, 15.4, 15.5_
  - [x]* 16.4 Escrever teste de propriedade para validação de ID inicial
    - **Property 37: Validação de ID inicial por intervalo**
    - **Validates: Requirements 15.4, 15.5**
  - [x]* 16.5 Escrever teste de propriedade para atribuição sequencial de IDs
    - **Property 38: Atribuição de IDs sequenciais a partir do ID inicial**
    - **Validates: Requirements 15.2**

- [x] 17. Implementar Validador_Padronizacao (`src/core/validator.ts`)
  - [x] 17.1 Implementar validação de padronização e geração de alertas
    - Implementar `validateStandardization` comparando campos com lista aplicável (case/trim-insensível), gerando alertas por campo/valor/linha; ignorar campos sem lista
    - _Requirements: 17.1, 17.2, 17.3, 17.5_
  - [x]* 17.2 Escrever teste de propriedade para detecção de divergências
    - **Property 40: Detecção de divergências de padronização**
    - **Validates: Requirements 17.1, 17.2, 17.5**
  - [x]* 17.3 Escrever teste unitário para mensagem de ausência de divergências
    - Verificar mensagem informando que nenhum valor divergente foi encontrado
    - _Requirements: 17.4_

- [x] 18. Implementar Gerador_Arquivo e abas auxiliares (`src/core/auxSheets.ts`, `src/core/generator.ts`)
  - [x] 18.1 Implementar geração das abas auxiliares do Financeiro (`auxSheets.ts`)
    - Implementar `detectBankData`, `detectChartOfAccounts`, `buildPending` (Dados Bancários, Plano de Contas, Pendências)
    - _Requirements: 18.4, 18.5_
  - [x] 18.2 Implementar geração do `.xlsx` e download local (`generator.ts`)
    - Implementar `generateXlsx` (todas as colunas do modelo na ordem/nome corretos + abas auxiliares condicionais) e `triggerDownload`
    - Lançar `GENERATION_FAILED` sem arquivo parcial, preservando dados em memória
    - _Requirements: 18.1, 18.2, 18.3, 18.6, 18.7_
  - [x]* 18.3 Escrever teste de propriedade para round-trip da estrutura do arquivo
    - **Property 41: Round-trip de estrutura do arquivo gerado**
    - **Validates: Requirements 18.1, 18.3**
  - [x]* 18.4 Escrever teste de propriedade para presença condicional de abas auxiliares
    - **Property 42: Presença condicional de abas auxiliares (Financeiro)**
    - **Validates: Requirements 18.4, 18.5**

- [x] 19. Checkpoint - Ensure all tests pass
  - Ensure all tests pass, ask the user if questions arise.

- [x] 20. Implementar a camada de UI de orquestração
  - [x] 20.1 Implementar estado da aplicação e seletores (tipo, origem, estilo, IDs iniciais)
    - Implementar `AppState`; seletores com as listas fixas de Tipo_De_Dado, Sistema_De_Origem e Estilo_De_Caixa (padrão Original)
    - Bloquear conversão enquanto tipo/origem não selecionados; carregar/apresentar colunas do modelo ao selecionar tipo
    - _Requirements: 4.1, 4.2, 4.3, 4.4, 4.5, 4.6, 5.1, 5.3, 14.1, 15.1, 15.3_
  - [x] 20.2 Implementar controles de upload, seleção de aba e opções de parser
    - Controle de upload (limite 50 MB), seleção de aba, opções de tipo de relatório Navis, aba/ano do BPO e seleção manual de cabeçalho
    - _Requirements: 3.4, 7.3, 8.1, 10.2, 10.4_
  - [x] 20.3 Implementar a orquestração do pipeline de conversão
    - Encadear leitor → pré-processamento por origem → encoding → remoção de linhas → mapeador → consolidação → vínculo → formatação → remoção sem identificador → atribuição de IDs → validação → geração
    - Propagar `ConversionError` sem gerar arquivo parcial; acumular `Alert`s
    - _Requirements: 5.2, 7.8, 16.1_
  - [x] 20.4 Implementar apresentação de alertas, mensagens de erro e controle de download
    - Exibir todos os alertas acumulados, mensagens de erro tipadas e o controle de download visível
    - _Requirements: 13.6, 13.7, 17.3, 17.4, 18.2_
  - [x] 20.5 Implementar guarda de inicialização (bibliotecas e recursos)
    - Verificar inicialização de SheetJS e carregamento de assets; bloquear conversão com `LIB_INIT_FAILED` / `*_LOAD_FAILED`
    - _Requirements: 1.1, 1.3, 1.4, 1.7, 11.8, 18.6_
  - [x]* 20.6 Escrever testes unitários das listas e guardas da UI
    - Testar listas fixas (tipos/origens/estilos), guardas de estado (sem tipo/origem) e bloqueio de inicialização
    - _Requirements: 4.1, 4.2, 5.1, 5.3, 1.7_

- [x] 21. Configurar deploy no GitHub Pages
  - [x] 21.1 Criar workflow GitHub Actions de build e publicação no Pages
    - Criar `.github/workflows/deploy.yml` usando `actions/configure-pages`, `actions/upload-pages-artifact`, `actions/deploy-pages`
    - Build a partir do branch de publicação; build com `base=/<repositorio>/`; falha interrompe publicação preservando a versão anterior com log consultável
    - _Requirements: 2.2, 2.3, 2.4, 2.6_
  - [x]* 21.2 Verificar build estático e resolução de assets por base path
    - Verificar que a saída de build é apenas HTML/CSS/JS e que URLs de assets usam `BASE_URL` (sem referências absolutas quebradas)
    - _Requirements: 1.2, 2.1, 2.5_

- [x] 22. Checkpoint final - Ensure all tests pass
  - Ensure all tests pass, ask the user if questions arise.

## Notes

- Tarefas marcadas com `*` são opcionais (testes) e podem ser puladas para um MVP mais rápido; as tarefas de implementação (sem `*`) são obrigatórias.
- Cada tarefa referencia requisitos específicos para rastreabilidade.
- Cada uma das 42 propriedades de correção do design é implementada por um único teste de propriedade dedicado (fast-check, ≥ 100 iterações), anotado com número da propriedade e cláusula de requisito.
- Critérios puramente arquiteturais (client-side, sem backend), de deploy (GitHub Actions/Pages) e de apresentação de UI são cobertos por testes de exemplo/unitários e verificação de build, não por testes de propriedade.
- Os checkpoints garantem validação incremental do núcleo puro antes de avançar para a orquestração e o deploy.

## Task Dependency Graph

```json
{
  "waves": [
    { "id": 0, "tasks": ["1.1"] },
    { "id": 1, "tasks": ["1.2", "1.3", "21.1"] },
    { "id": 2, "tasks": ["2.1", "3.1", "4.1", "6.1", "7.1", "8.1", "9.1", "12.1", "13.1", "14.1", "16.1", "16.3", "18.1"] },
    { "id": 3, "tasks": ["11.1", "17.1", "18.2", "2.2"] },
    { "id": 4, "tasks": ["3.2", "3.3", "3.4", "3.5", "4.2", "4.3", "6.2", "6.3", "6.4", "7.2", "7.3", "8.2", "8.3", "8.4", "8.5", "8.6", "9.2", "9.3", "9.4", "9.5", "11.2", "11.3", "11.4", "11.5", "12.2", "12.3", "13.2", "13.3", "13.4", "14.2", "14.3", "14.4", "14.5", "14.6", "14.7", "14.8", "14.9", "14.10", "16.2", "16.4", "16.5", "17.2", "17.3", "18.3", "18.4"] },
    { "id": 5, "tasks": ["20.1", "20.2", "20.3", "20.4", "20.5"] },
    { "id": 6, "tasks": ["20.6", "21.2"] }
  ]
}
```
