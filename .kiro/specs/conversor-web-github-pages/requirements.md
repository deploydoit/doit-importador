# Requirements Document

## Introduction

Esta funcionalidade reescreve o Conversor de Dados DOit — atualmente uma aplicação Streamlit/Python (`dashboards/app_conversor.py` + `scripts/`) — como um site estático que roda inteiramente no navegador do usuário (client-side, em JavaScript), sem qualquer backend. O objetivo é permitir a publicação via GitHub Pages, que serve apenas arquivos estáticos (HTML/CSS/JS) e não executa Python.

Toda a lógica de conversão hoje implementada em Python (leitura de planilhas, correção de encoding, parsers especiais, mapeamento de colunas, vínculo de IDs, formatação e geração do arquivo de saída) deve ser portada para o navegador. A leitura e escrita de arquivos Excel deve ser feita com biblioteca JavaScript (por exemplo SheetJS). Os artefatos de configuração hoje lidos do disco (modelos em `models/*.xlsx` e listas de padronização em `mappings/*.txt`) devem ser empacotados como recursos estáticos carregáveis no navegador.

O resultado da conversão deve preservar as funcionalidades atuais e permanecer no padrão DOit, gerando um arquivo `.xlsx` para download local. Nenhum dado do usuário deve ser transmitido para servidores externos — todo o processamento ocorre no navegador.

## Glossary

- **Conversor_Web**: A aplicação web estática completa executada no navegador, sem backend, responsável por converter planilhas para o padrão DOit.
- **Padrão_DOit**: O conjunto de colunas e formatos esperados para importação no sistema DOit, definido pelos arquivos de modelo (`models/*.xlsx`).
- **Planilha_Origem**: O arquivo enviado pelo usuário (`.xlsx`, `.xls` ou `.csv`) contendo os dados a serem convertidos.
- **Arquivo_Convertido**: O arquivo `.xlsx` gerado pelo Conversor_Web no Padrão_DOit, disponibilizado para download.
- **Tipo_De_Dado**: A categoria de dados a converter. Valores permitidos: Cadastros, Contatos Relacionados, Projetos, Financeiro, Horas Trabalhadas, Usuários, Produtos, Vendas.
- **Sistema_De_Origem**: O sistema que gerou a Planilha_Origem. Valores permitidos: DOit Coleta, Conta Azul, Navis, Omie, ClickUp, Sienge, Trello, Excel Desestruturado, Financeiro Horizontal (BPO), Outlook, Excel Manual.
- **Leitor_Planilha**: O componente do Conversor_Web que lê arquivos `.xlsx`, `.xls` e `.csv` no navegador e produz uma tabela de dados em memória.
- **Corretor_Encoding**: O componente que detecta e corrige mojibake (ex.: "RazÃ£o Social" → "Razão Social") em nomes de colunas e valores de texto.
- **Parser_Navis**: O componente que interpreta relatórios do sistema Navis, detectando o tipo de relatório e extraindo os registros.
- **Parser_BPO**: O componente que interpreta planilhas Financeiro Horizontal (BPO), extraindo lançamentos organizados por mês.
- **Processador_Outlook**: O componente que pós-processa contatos exportados do Outlook (junção de nome, telefones e anotações).
- **Detector_Cabecalho**: O componente que identifica automaticamente a linha de cabeçalho real em planilhas (ex.: Conta Azul).
- **Mapeador_Colunas**: O componente que associa colunas da Planilha_Origem às colunas do Padrão_DOit, por regra pré-definida e por similaridade automática.
- **Vinculador_IDs**: O componente que preenche colunas de ID em uma planilha convertida a partir de planilhas de referência já convertidas.
- **Formatador_Texto**: O componente que aplica estilo de caixa (Primeira Maiúscula, TUDO MAIÚSCULO ou Original) e formatações de campo (telefone, CPF/CNPJ, CEP, data, e-mail, estado).
- **Gerador_Arquivo**: O componente que monta e serializa o Arquivo_Convertido em `.xlsx`.
- **Validador_Padronizacao**: O componente que compara os valores convertidos com as listas de padronização do DOit e emite alertas.
- **Recurso_Estatico**: Arquivo empacotado com o site (modelo `.xlsx` ou lista `.txt`) carregado pelo navegador em tempo de execução.
- **Estilo_De_Caixa**: A opção de formatação de texto escolhida pelo usuário: Primeira Maiúscula, TUDO MAIÚSCULO ou Original.
- **Site_Estatico**: O conjunto de arquivos HTML/CSS/JS publicável no GitHub Pages.
- **Pipeline_Deploy**: O fluxo de trabalho automatizado (GitHub Actions) que publica o Site_Estatico no GitHub Pages.
- **Mojibake**: Texto corrompido por interpretação incorreta de encoding (ex.: UTF-8 lido como Latin-1).

## Requirements

### Requirement 1: Execução 100% no navegador sem backend

**User Story:** Como usuário, quero que o conversor rode inteiramente no meu navegador, para que eu possa usá-lo publicado no GitHub Pages sem depender de servidor Python.

#### Acceptance Criteria

1. THE Conversor_Web SHALL executar toda a lógica de leitura, conversão, formatação e geração de arquivos no navegador do usuário, sem realizar chamadas de rede que transmitam o conteúdo da Planilha_Origem.
2. THE Conversor_Web SHALL ser composto exclusivamente por arquivos estáticos HTML, CSS e JavaScript servidos sem processamento no lado do servidor.
3. WHILE uma conversão está em andamento, THE Conversor_Web SHALL processar os dados na memória do navegador do usuário, sem enviar o conteúdo da Planilha_Origem para servidores externos.
4. THE Conversor_Web SHALL funcionar sem executar código Python em tempo de execução no lado do servidor ou no navegador.
5. WHERE o Conversor_Web precisa de bibliotecas de terceiros para ler ou escrever arquivos Excel, THE Conversor_Web SHALL utilizar bibliotecas JavaScript executáveis no navegador, carregadas como parte dos arquivos estáticos.
6. WHEN o usuário seleciona um arquivo cujo tamanho excede 50 MB, THE Conversor_Web SHALL rejeitar o arquivo e exibir uma mensagem de erro indicando que o limite de tamanho foi excedido, preservando o estado anterior da interface sem iniciar a conversão.
7. IF a biblioteca JavaScript de leitura ou escrita de arquivos Excel não puder ser carregada, THEN THE Conversor_Web SHALL exibir uma mensagem de erro indicando que o conversor não pôde ser inicializado e SHALL impedir o início de qualquer conversão.

### Requirement 2: Publicação como site estático no GitHub Pages

**User Story:** Como mantenedor, quero publicar o conversor no GitHub Pages através de um fluxo automatizado, para que atualizações no código sejam publicadas sem passos manuais.

#### Acceptance Criteria

1. THE Site_Estatico SHALL ser servível pelo GitHub Pages usando apenas arquivos estáticos HTML, CSS e JavaScript, sem execução de código server-side.
2. WHEN alterações são integradas ao branch de publicação configurado, THE Pipeline_Deploy SHALL construir o Site_Estatico a partir do estado versionado desse branch.
3. WHEN a etapa de build do Pipeline_Deploy é concluída com sucesso, THE Pipeline_Deploy SHALL publicar o Site_Estatico no GitHub Pages sem intervenção manual.
4. IF a etapa de build do Pipeline_Deploy falha, THEN THE Pipeline_Deploy SHALL interromper a publicação, preservar a versão publicada anteriormente sem alteração e registrar a falha de forma consultável no histórico de execuções do fluxo de trabalho.
5. WHEN o Site_Estatico é servido a partir de um subcaminho de projeto do GitHub Pages (por exemplo `https://<usuario>.github.io/<repositorio>/`), THE Site_Estatico SHALL resolver todas as referências a arquivos estáticos (HTML, CSS, JavaScript e Recurso_Estatico) relativas a esse subcaminho, sem referências quebradas.
6. THE Pipeline_Deploy SHALL ser definido como um fluxo de trabalho GitHub Actions versionado no repositório.

### Requirement 3: Upload de planilha do cliente

**User Story:** Como usuário, quero enviar a planilha do cliente em diferentes formatos, para que eu possa converter dados independentemente do formato de exportação.

#### Acceptance Criteria

1. THE Leitor_Planilha SHALL aceitar arquivos com extensão `.xlsx`, `.xls` ou `.csv`, avaliando a extensão de forma insensível a maiúsculas e minúsculas (por exemplo, `.CSV` e `.csv` são aceitos de forma equivalente).
2. WHEN o usuário seleciona um arquivo suportado com tamanho menor ou igual a 50 MB, THE Leitor_Planilha SHALL carregar o conteúdo em uma tabela de dados em memória em até 30 segundos.
3. IF o arquivo selecionado possui tamanho superior a 50 MB, THEN THE Conversor_Web SHALL rejeitar o arquivo e exibir uma mensagem indicando que o arquivo excede o tamanho máximo permitido de 50 MB, preservando o estado anterior sem carregar dados.
4. WHEN a Planilha_Origem contém múltiplas abas, THE Leitor_Planilha SHALL apresentar a lista de abas disponíveis para seleção do usuário e selecionar a primeira aba como padrão até que o usuário escolha outra.
5. IF o arquivo selecionado não possui extensão `.xlsx`, `.xls` ou `.csv`, THEN THE Conversor_Web SHALL rejeitar o arquivo e exibir uma mensagem de formato não suportado, sem carregar dados em memória.
6. IF a Planilha_Origem não contém nenhuma linha de dados além do cabeçalho após a leitura, THEN THE Conversor_Web SHALL interromper a conversão e exibir uma mensagem indicando que o arquivo está vazio.
7. WHEN um arquivo `.csv` não é lido com sucesso em codificação UTF-8, THE Leitor_Planilha SHALL tentar a leitura em codificação Latin-1.
8. IF a leitura de um arquivo `.csv` falha tanto em codificação UTF-8 quanto em Latin-1, THEN THE Conversor_Web SHALL rejeitar o arquivo e exibir uma mensagem indicando que a codificação do arquivo não é suportada, sem carregar dados em memória.

### Requirement 4: Seleção de tipo de dado

**User Story:** Como usuário, quero escolher o tipo de dado que estou convertendo, para que o conversor use o modelo correto do Padrão_DOit.

#### Acceptance Criteria

1. THE Conversor_Web SHALL oferecer a seleção dos seguintes Tipo_De_Dado, e somente destes: Cadastros, Contatos Relacionados, Projetos, Financeiro, Horas Trabalhadas, Usuários, Produtos e Vendas.
2. WHILE nenhum Tipo_De_Dado tiver sido selecionado pelo usuário, THE Conversor_Web SHALL exibir um estado inicial sem Tipo_De_Dado selecionado e SHALL impedir o início da conversão.
3. WHEN o usuário seleciona um Tipo_De_Dado, THE Conversor_Web SHALL carregar as colunas do modelo correspondente do Padrão_DOit em até 5 segundos.
4. WHEN o carregamento das colunas do modelo correspondente é concluído com sucesso, THE Conversor_Web SHALL exibir a lista de colunas do modelo carregado e indicar qual Tipo_De_Dado está selecionado.
5. WHEN o usuário seleciona um Tipo_De_Dado diferente do atualmente selecionado, THE Conversor_Web SHALL substituir as colunas exibidas pelas colunas do modelo recém-selecionado.
6. IF o modelo correspondente ao Tipo_De_Dado selecionado não pode ser carregado, THEN THE Conversor_Web SHALL exibir uma mensagem de erro identificando o Tipo_De_Dado e o modelo ausente, SHALL manter o estado como nenhum Tipo_De_Dado selecionado e SHALL impedir o início da conversão.

### Requirement 5: Seleção de sistema de origem

**User Story:** Como usuário, quero escolher o sistema que gerou a planilha, para que o conversor aplique o mapeamento e o tratamento adequados.

#### Acceptance Criteria

1. THE Conversor_Web SHALL oferecer a seleção dos Sistema_De_Origem a partir de uma lista contendo exatamente as 11 opções: DOit Coleta, Conta Azul, Navis, Omie, ClickUp, Sienge, Trello, Excel Desestruturado, Financeiro Horizontal (BPO), Outlook e Excel Manual.
2. WHEN o usuário seleciona um Sistema_De_Origem, THE Conversor_Web SHALL aplicar o conjunto de regras de mapeamento e pré-processamento associado a esse sistema.
3. IF o usuário tenta iniciar a conversão sem ter selecionado um Sistema_De_Origem, THEN THE Conversor_Web SHALL impedir o início da conversão e exibir uma mensagem indicando que a seleção do sistema de origem é obrigatória.
4. WHERE o Sistema_De_Origem selecionado é Excel Manual ou Excel Desestruturado, THE Mapeador_Colunas SHALL mapear cada coluna de origem para a coluna de destino cujo nome apresente similaridade igual ou superior a 80%, quando não houver regra pré-definida aplicável.
5. IF nenhuma coluna de destino atinge similaridade de nome igual ou superior a 80% para uma coluna de origem, THEN THE Mapeador_Colunas SHALL deixar essa coluna sem mapeamento automático e sinalizar ao usuário que o mapeamento manual é necessário para essa coluna.
6. IF o conjunto de regras de mapeamento associado ao Sistema_De_Origem selecionado não estiver disponível, THEN THE Conversor_Web SHALL interromper o processamento, preservar o arquivo de entrada sem alterações e exibir uma mensagem indicando a indisponibilidade das regras para o sistema selecionado.

### Requirement 6: Correção de encoding (mojibake)

**User Story:** Como usuário, quero que textos corrompidos por encoding sejam corrigidos automaticamente, para que os dados convertidos fiquem legíveis.

#### Acceptance Criteria

1. WHEN a Planilha_Origem contém nomes de colunas com pelo menos um marcador de Mojibake (sequência de caracteres resultante da interpretação incorreta de bytes UTF-8 como Latin-1 ou Windows-1252, por exemplo "Ã§", "Ã£", "Ã©", "Ãª"), THE Corretor_Encoding SHALL substituir cada nome de coluna afetado pelo texto acentuado UTF-8 correspondente, mantendo inalterados os nomes de colunas sem marcadores.
2. WHEN a Planilha_Origem contém valores de texto com pelo menos um marcador de Mojibake, THE Corretor_Encoding SHALL substituir cada valor afetado pelo texto acentuado UTF-8 correspondente.
3. THE Corretor_Encoding SHALL detectar e corrigir Mojibake originado tanto da interpretação incorreta como Latin-1 quanto como Windows-1252.
4. THE Corretor_Encoding SHALL preservar, byte a byte, os valores de texto e nomes de colunas que não contêm nenhum marcador de Mojibake.
5. WHEN o Arquivo_Convertido é gerado, THE Corretor_Encoding SHALL aplicar a correção de encoding sobre todo o conteúdo textual do Arquivo_Convertido antes da geração final do arquivo.
6. IF um valor de texto contém marcadores de Mojibake que não podem ser mapeados de forma inequívoca para um texto UTF-8 válido, THEN THE Corretor_Encoding SHALL preservar o valor original sem alteração e registrar uma indicação de que o valor não pôde ser corrigido, sem interromper o processamento dos demais valores.

### Requirement 7: Parser especial para Navis

**User Story:** Como usuário, quero converter relatórios do Navis, para que eu possa importar clientes, fornecedores, contatos, projetos, horas e financeiro do Navis no DOit.

#### Acceptance Criteria

1. WHEN o Sistema_De_Origem é Navis, THE Parser_Navis SHALL detectar automaticamente o tipo de relatório analisando as primeiras 15 linhas da planilha em busca de palavras-chave identificadoras, retornando exatamente um dos valores: Movimentos de Conta Corrente, Contas a Pagar/Receber (Baixados), Contas a Pagar/Receber (Previsão), Consulta Projetos, Clientes, Fornecedores, Contatos ou Aplicação de Horas.
2. IF nenhum tipo de relatório é identificado após analisar as primeiras 15 linhas, THEN THE Parser_Navis SHALL classificar o relatório como tipo desconhecido e o Conversor_Web SHALL exibir uma mensagem indicando que o tipo de relatório não pôde ser identificado.
3. THE Parser_Navis SHALL permitir que o usuário sobrescreva manualmente o tipo de relatório detectado, passando a usar o tipo selecionado pelo usuário na extração dos registros.
4. THE Parser_Navis SHALL extrair registros dos tipos de relatório: Movimentos de Conta Corrente, Contas a Pagar/Receber (Baixados), Contas a Pagar/Receber (Previsão), Consulta Projetos, Clientes, Fornecedores, Contatos e Aplicação de Horas.
5. WHEN um relatório Navis possui linhas de metadados anteriores ao cabeçalho de dados, THE Parser_Navis SHALL localizar a linha de cabeçalho examinando as primeiras 20 linhas e selecionando a linha com maior número de correspondências às palavras-chave esperadas para o tipo de relatório, antes de extrair os registros.
6. WHEN um registro de relatório Navis ocupa múltiplas linhas (formato ficha ou classificação financeira continuada em linha subsequente), THE Parser_Navis SHALL consolidar as linhas relacionadas em um único registro antes de finalizar a extração.
7. IF um relatório Navis não contém nenhuma linha de dados após a linha de cabeçalho, THEN THE Parser_Navis SHALL retornar um conjunto de registros vazio e o Conversor_Web SHALL exibir uma mensagem indicando que nenhum registro foi encontrado, sem interromper a aplicação.
8. IF o Tipo_De_Dado selecionado é incompatível com o tipo de relatório Navis detectado, THEN THE Conversor_Web SHALL exibir um alerta indicando a incompatibilidade entre o Tipo_De_Dado selecionado e o tipo de relatório detectado.
9. IF o Parser_Navis não consegue interpretar o relatório, THEN THE Conversor_Web SHALL exibir uma mensagem de erro descritiva indicando a causa da falha, preservar os dados de origem sem alteração e interromper a conversão.

### Requirement 8: Parser especial para Financeiro Horizontal (BPO)

**User Story:** Como usuário, quero converter a planilha Financeiro Horizontal (BPO), para que os lançamentos distribuídos por mês sejam extraídos para o formato financeiro do DOit.

#### Acceptance Criteria

1. WHEN o Sistema_De_Origem selecionado é Financeiro Horizontal (BPO), THE Parser_BPO SHALL solicitar ao usuário a seleção da aba a ser processada e a informação do ano de referência, sendo o ano um valor inteiro entre 1900 e 2100.
2. IF o usuário não informar a aba ou o ano de referência, ou informar um ano fora do intervalo de 1900 a 2100, THEN THE Conversor_Web SHALL exibir uma mensagem de erro indicando o campo inválido ou ausente e interromper a conversão sem gerar arquivo de saída.
3. THE Parser_BPO SHALL detectar as colunas correspondentes a cada um dos 12 meses (janeiro a dezembro) na aba selecionada.
4. IF nenhuma coluna de mês for detectada na aba selecionada, THEN THE Conversor_Web SHALL exibir uma mensagem de erro indicando que a estrutura de colunas mensais não foi encontrada e interromper a conversão sem gerar arquivo de saída.
5. THE Parser_BPO SHALL extrair os lançamentos de cada mês detectado para linhas individuais, cada uma contendo os campos data, vencimento, descrição, valor, tipo, conciliado e categorias.
6. WHILE um lançamento de determinado mês possui valor vazio, nulo ou igual a zero, THE Parser_BPO SHALL omitir esse lançamento da saída, extraindo apenas os lançamentos com valor diferente de zero.
7. WHEN o ano de referência é informado, THE Parser_BPO SHALL compor o campo data de cada lançamento extraído combinando o ano de referência informado com o mês correspondente à coluna de origem do lançamento.
8. IF o Parser_BPO não consegue interpretar a planilha, THEN THE Conversor_Web SHALL exibir uma mensagem de erro descritiva indicando a causa da falha e interromper a conversão sem gerar arquivo de saída.

### Requirement 9: Pós-processamento de contatos do Outlook

**User Story:** Como usuário, quero converter contatos exportados do Outlook, para que nomes, telefones e informações adicionais sejam consolidados no formato de cadastro do DOit.

#### Acceptance Criteria

1. WHEN o Sistema_De_Origem é Outlook, THE Processador_Outlook SHALL concatenar os campos de nome (Primeiro nome, Segundo nome, Sobrenome, Sufixo) em um único campo de nome, na ordem Primeiro nome, Segundo nome, Sobrenome, Sufixo, separando os componentes preenchidos por um único espaço.
2. WHEN um ou mais dos componentes de nome (Primeiro nome, Segundo nome, Sobrenome, Sufixo) estão vazios, THE Processador_Outlook SHALL omitir os componentes vazios da concatenação, sem inserir espaços consecutivos, e SHALL remover espaços no início e no fim do campo de nome resultante.
3. WHEN o Sistema_De_Origem é Outlook, THE Processador_Outlook SHALL distribuir os telefones adicionais nos campos de telefone principais disponíveis do Padrão_DOit, preenchendo-os na ordem em que aparecem na origem e ocupando os campos principais vazios do primeiro ao último.
4. IF a quantidade de telefones a distribuir excede a quantidade de campos de telefone principais disponíveis, THEN THE Processador_Outlook SHALL consolidar todos os telefones excedentes no campo Fax Comercial, separando cada telefone excedente por vírgula seguida de espaço.
5. WHEN o Sistema_De_Origem é Outlook, THE Processador_Outlook SHALL consolidar todos os campos de origem não mapeados para campos do Padrão_DOit em um único campo de Anotações, incluindo apenas os campos com valor preenchido e separando cada par por quebra de linha no formato "nome do campo: valor".
6. IF o campo de nome resultante está vazio ou contém apenas espaços após a concatenação e remoção de espaços, THEN THE Processador_Outlook SHALL remover a linha correspondente do resultado.
7. WHEN a exportação do Outlook está no formato `.csv`, THE Processador_Outlook SHALL aplicar o mesmo pós-processamento definido nos critérios 1 a 6 aplicado à exportação `.xlsx`, produzindo resultado idêntico para dados de conteúdo equivalente.

### Requirement 10: Detecção automática de cabeçalho (Conta Azul)

**User Story:** Como usuário, quero que o conversor identifique a linha de cabeçalho correta em planilhas do Conta Azul, para que colunas de índice e linhas de título não atrapalhem o mapeamento.

#### Acceptance Criteria

1. WHEN o Sistema_De_Origem é Conta Azul, THE Detector_Cabecalho SHALL examinar as primeiras 20 linhas da aba e identificar como linha de cabeçalho a primeira linha cujas células correspondam a pelo menos 2 dos nomes de coluna esperados para o Tipo_De_Dado, comparando de forma insensível a maiúsculas/minúsculas e a espaços em branco nas extremidades.
2. IF nenhuma das primeiras 20 linhas da aba contém pelo menos 2 nomes de coluna esperados para o Tipo_De_Dado, THEN THE Conversor_Web SHALL interromper a detecção automática, apresentar mensagem indicando que o cabeçalho não pôde ser identificado e permitir que o usuário selecione manualmente a linha de cabeçalho.
3. WHEN o Sistema_De_Origem é Conta Azul, THE Detector_Cabecalho SHALL selecionar automaticamente a aba cujo nome contenha uma das palavras-chave associadas ao Tipo_De_Dado, comparando de forma insensível a maiúsculas/minúsculas e a espaços em branco nas extremidades.
4. IF nenhuma aba pode ser identificada automaticamente pelo nome, THEN THE Conversor_Web SHALL apresentar a lista de todas as abas disponíveis na planilha e aguardar a seleção de uma aba pelo usuário antes de prosseguir.
5. WHEN a linha de cabeçalho é identificada, THE Detector_Cabecalho SHALL descartar todas as colunas cujo nome esteja vazio, seja composto apenas por espaços em branco, ou seja composto exclusivamente por caracteres numéricos, mantendo as demais colunas para o mapeamento.

### Requirement 11: Mapeamento automático de colunas para o Padrão DOit

**User Story:** Como usuário, quero que as colunas da planilha de origem sejam associadas às colunas do Padrão_DOit, para que eu não precise mapear manualmente cada coluna.

#### Acceptance Criteria

1. WHEN um Sistema_De_Origem e um Tipo_De_Dado possuem regra de mapeamento pré-definida, THE Mapeador_Colunas SHALL mapear cada coluna da Planilha_Origem para a coluna correspondente do Padrão_DOit conforme a associação declarada nessa regra.
2. WHERE uma coluna do Padrão_DOit não é preenchida por regra pré-definida, THE Mapeador_Colunas SHALL associá-la à coluna da Planilha_Origem cujo nome apresente grau de similaridade igual ou superior a 80%, ignorando diferenças de maiúsculas/minúsculas e de espaços em branco nas extremidades.
3. IF, ao mapear por similaridade, mais de uma coluna da Planilha_Origem atinge o limiar de 80% para a mesma coluna do Padrão_DOit, THEN THE Mapeador_Colunas SHALL selecionar a coluna de origem de maior grau de similaridade e, em caso de empate, a primeira em ordem de ocorrência na Planilha_Origem.
4. IF uma coluna da Planilha_Origem já foi associada a uma coluna do Padrão_DOit, THEN THE Mapeador_Colunas SHALL impedir que essa mesma coluna de origem seja associada a qualquer outra coluna do Padrão_DOit.
5. WHEN uma coluna do Padrão_DOit não possui nenhuma coluna de origem correspondente após a aplicação das regras pré-definidas e da similaridade, THE Mapeador_Colunas SHALL preenchê-la com valor vazio (string de comprimento zero) em todas as linhas.
6. WHEN o Conversor_Web é iniciado, THE Conversor_Web SHALL carregar as colunas dos modelos do Padrão_DOit a partir de Recurso_Estatico empacotado com o site.
7. WHEN o Conversor_Web é iniciado, THE Conversor_Web SHALL carregar as listas de padronização a partir de Recurso_Estatico empacotado com o site.
8. IF um Recurso_Estatico necessário (modelo do Padrão_DOit ou lista de padronização) não pode ser localizado ou lido durante o carregamento, THEN THE Conversor_Web SHALL interromper a operação de mapeamento e exibir mensagem de erro indicando qual recurso não pôde ser carregado, sem produzir uma Planilha_Origem mapeada parcial.

### Requirement 12: Consolidação de campos não mapeados

**User Story:** Como usuário, quero que dados de colunas não mapeadas não sejam perdidos, para que informações adicionais permaneçam no arquivo convertido.

#### Acceptance Criteria

1. WHERE o Tipo_De_Dado é Cadastros, WHEN existem colunas de origem não mapeadas com pelo menos um valor não vazio, THE Conversor_Web SHALL consolidar esses valores no campo Anotações, registrando cada valor precedido do nome da coluna de origem e separando os pares entre si por um delimitador único e consistente.
2. WHERE o Tipo_De_Dado é Cadastros ou (Financeiro ou Horas Trabalhadas com campo Descrição), IF uma coluna de origem não mapeada não possui nenhum valor não vazio, THEN THE Conversor_Web SHALL omitir essa coluna da consolidação sem gerar entrada correspondente.
3. WHERE o Tipo_De_Dado é Projetos, THE Conversor_Web SHALL distribuir os valores de cada coluna de origem não mapeada que possua ao menos um valor não vazio na coluna customizável cujo tipo (texto, número, data, moeda ou booleano) corresponda ao conteúdo da coluna, respeitando a ordem em que as colunas customizáveis estão disponíveis.
4. WHERE o Tipo_De_Dado é Projetos, IF não há coluna customizável disponível do tipo correspondente a uma coluna de origem não mapeada, THEN THE Conversor_Web SHALL deixar essa coluna de origem sem distribuição em coluna customizável.
5. WHERE o Tipo_De_Dado é Financeiro ou Horas Trabalhadas, WHEN o campo Descrição existe no modelo correspondente e há colunas de origem não mapeadas com valores não vazios, THE Conversor_Web SHALL consolidar esses valores no campo Descrição, registrando cada valor precedido do nome da coluna de origem e separando os pares entre si por um delimitador único e consistente.
6. WHERE o Tipo_De_Dado é Financeiro ou Horas Trabalhadas, IF o campo Descrição não existe no modelo correspondente, THEN THE Conversor_Web SHALL concluir a conversão sem consolidar as colunas de origem não mapeadas.

### Requirement 13: Vínculo de IDs entre planilhas convertidas

**User Story:** Como usuário, quero vincular IDs de planilhas já convertidas, para que projetos, financeiro e contatos relacionados referenciem os registros corretos.

#### Acceptance Criteria

1. WHERE o Tipo_De_Dado é Projetos e uma planilha de Cadastros convertida é fornecida, THE Vinculador_IDs SHALL preencher a coluna ID DO CADASTRO relacionando o nome do cliente ao ID do cadastro.
2. WHERE o Tipo_De_Dado é Projetos e uma planilha de Usuários convertida é fornecida, THE Vinculador_IDs SHALL preencher a coluna ID LÍDER relacionando o nome do responsável ao ID do usuário.
3. WHERE o Tipo_De_Dado é Financeiro e uma planilha de Cadastros convertida é fornecida, THE Vinculador_IDs SHALL preencher a coluna ID DE / PARA relacionando o nome do favorecido ao ID do cadastro.
4. WHERE o Tipo_De_Dado é Financeiro e uma planilha de Projetos convertida é fornecida, THE Vinculador_IDs SHALL preencher a coluna ID PROJETO relacionando o nome do projeto ao ID do projeto.
5. WHERE o Tipo_De_Dado é Contatos Relacionados e uma planilha de Cadastros convertida é fornecida, THE Vinculador_IDs SHALL preencher as colunas ID PAI e ID FILHO relacionando os nomes aos IDs do cadastro.
6. IF um nome não é encontrado na planilha de referência, THEN THE Vinculador_IDs SHALL deixar o ID correspondente vazio e THE Conversor_Web SHALL exibir um alerta listando os nomes não encontrados.
7. IF há nomes preenchidos que exigem vínculo mas nenhuma planilha de referência é fornecida, THEN THE Conversor_Web SHALL exibir um aviso solicitando o upload da planilha de referência.
8. THE Vinculador_IDs SHALL relacionar um nome ao seu ID por correspondência exata após normalização, considerando os textos equivalentes quando forem iguais após a remoção de espaços à esquerda e à direita e a comparação sem diferenciação entre maiúsculas e minúsculas.
9. IF um nome corresponde a mais de um registro na planilha de referência, THEN THE Vinculador_IDs SHALL preencher o ID correspondente com o ID do primeiro registro correspondente e THE Conversor_Web SHALL exibir um alerta listando os nomes com correspondência ambígua.
10. IF a planilha de referência fornecida não contém as colunas de nome e de ID necessárias para o vínculo, THEN THE Vinculador_IDs SHALL deixar os IDs correspondentes vazios e THE Conversor_Web SHALL exibir uma mensagem de erro identificando a coluna de referência ausente.

### Requirement 14: Formatação de texto e de campos

**User Story:** Como usuário, quero controlar como os textos são formatados, para que o resultado siga o padrão de apresentação desejado.

#### Acceptance Criteria

1. THE Conversor_Web SHALL oferecer os Estilo_De_Caixa: Primeira Maiúscula, TUDO MAIÚSCULO e Original, adotando Original como estado padrão do seletor até que o usuário escolha outro.
2. WHEN o Estilo_De_Caixa é Primeira Maiúscula, THE Formatador_Texto SHALL converter cada campo de texto palavra a palavra, colocando a primeira letra de cada palavra em maiúscula e as demais em minúscula, mantendo em minúsculo as preposições e artigos (de, da, do, das, dos, e, a, o, as, os) quando não forem a primeira palavra do campo.
3. WHEN o Estilo_De_Caixa é TUDO MAIÚSCULO, THE Formatador_Texto SHALL converter os campos de texto integralmente para maiúsculas.
4. WHEN o Estilo_De_Caixa é Original, THE Formatador_Texto SHALL preservar o texto sem alteração de caixa.
5. WHEN um campo de texto está vazio ou nulo, THE Formatador_Texto SHALL preservar o campo vazio sem aplicar formatação de caixa.
6. WHEN um campo de telefone contém 11 dígitos de assinante, THE Formatador_Texto SHALL formatá-lo no padrão `+55 (XX) XXXXX-XXXX`, ignorando caracteres não numéricos presentes no valor de origem.
7. WHEN um campo de telefone contém 10 dígitos de assinante, THE Formatador_Texto SHALL formatá-lo no padrão `+55 (XX) XXXX-XXXX`, ignorando caracteres não numéricos presentes no valor de origem.
8. IF um campo de telefone não contém 10 nem 11 dígitos de assinante, THEN THE Formatador_Texto SHALL preservar o valor original e sinalizar que o telefone não pôde ser formatado.
9. WHEN um campo de CEP contém exatamente 8 dígitos numéricos, THE Formatador_Texto SHALL formatá-lo no padrão `XXXXX-XXX`.
10. IF um campo de CEP não contém exatamente 8 dígitos numéricos, THEN THE Formatador_Texto SHALL preservar o valor original e sinalizar que o CEP não pôde ser formatado.
11. WHEN um campo de data contém um valor de data reconhecível, THE Formatador_Texto SHALL formatá-lo como texto no padrão `dd/mm/aaaa`.
12. IF um campo de data não contém um valor de data reconhecível, THEN THE Formatador_Texto SHALL preservar o valor original e sinalizar que a data não pôde ser formatada.
13. THE Formatador_Texto SHALL formatar os campos de e-mail em letras minúsculas e sem espaços.
14. WHEN um campo de estado contém uma sigla de duas letras, THE Formatador_Texto SHALL formatá-lo com os dois caracteres em maiúsculas; caso contrário, THE Formatador_Texto SHALL preservar o valor original.

### Requirement 15: ID inicial configurável por tipo

**User Story:** Como usuário, quero definir o ID inicial de cada tipo, para que os IDs gerados continuem a numeração já existente no DOit.

#### Acceptance Criteria

1. THE Conversor_Web SHALL permitir que o usuário defina o ID inicial para os tipos Cadastros, Projetos, Financeiro, Horas Trabalhadas e Usuários.
2. WHEN o Arquivo_Convertido possui coluna ID, THE Gerador_Arquivo SHALL atribuir IDs sequenciais inteiros iniciando no ID inicial definido para o Tipo_De_Dado e incrementando de 1 em 1 a cada registro subsequente.
3. WHERE o usuário não define um ID inicial para um Tipo_De_Dado configurável, THE Conversor_Web SHALL usar o valor padrão definido para esse Tipo_De_Dado.
4. THE Conversor_Web SHALL aceitar como ID inicial apenas valores inteiros no intervalo de 1 a 999.999.999.
5. IF o usuário informa um ID inicial não inteiro ou fora do intervalo de 1 a 999.999.999, THEN THE Conversor_Web SHALL rejeitar o valor, exibir uma mensagem indicando o intervalo válido e manter o último ID inicial válido definido para o Tipo_De_Dado.

### Requirement 16: Remoção de linhas inválidas

**User Story:** Como usuário, quero que linhas vazias ou de exemplo sejam removidas, para que o arquivo convertido contenha apenas registros reais.

#### Acceptance Criteria

1. WHEN o Sistema_De_Origem é DOit Coleta, THE Conversor_Web SHALL remover, antes da conversão, as linhas que precedem o primeiro registro real e que contêm textos descritivos/instrucionais dos campos ou valores de exemplo demonstrativo do modelo.
2. THE Conversor_Web SHALL remover, antes de atribuir IDs, as linhas em que todos os campos de identificação do Tipo_De_Dado (nome ou identificador principal) estão vazios ou contêm apenas espaços em branco.
3. THE Conversor_Web SHALL preservar todas as linhas que possuem ao menos um campo de identificação do Tipo_De_Dado preenchido com valor não vazio.
4. IF a Planilha_Origem não contém nenhum registro restante após a remoção de linhas inválidas, THEN THE Conversor_Web SHALL interromper a conversão e exibir uma mensagem indicando a ausência de dados para conversão.

### Requirement 17: Validação de padronizações e alertas

**User Story:** Como usuário, quero ser avisado sobre valores fora dos padrões do DOit, para que eu possa corrigi-los antes de importar.

#### Acceptance Criteria

1. WHEN a conversão é concluída, THE Validador_Padronizacao SHALL comparar, para cada campo do Arquivo_Convertido que possui lista de padronização aplicável, o valor convertido com a lista de padronização correspondente do DOit.
2. IF um valor não vazio de um campo com lista de padronização aplicável não corresponde a nenhum item da lista, desconsiderando diferenças de caixa e espaços no início e no fim do texto, THEN THE Validador_Padronizacao SHALL registrar um alerta identificando o nome do campo, o valor divergente e a linha correspondente.
3. WHEN a validação gera um ou mais alertas, THE Conversor_Web SHALL apresentar ao usuário todos os alertas gerados, cada um identificando o campo e o valor divergente.
4. WHEN a validação é concluída sem nenhum alerta, THE Conversor_Web SHALL informar ao usuário que nenhum valor divergente das listas de padronização foi encontrado.
5. WHERE um campo do Arquivo_Convertido não possui lista de padronização aplicável, THE Validador_Padronizacao SHALL não gerar alertas de padronização para esse campo.

### Requirement 18: Geração e download do arquivo convertido

**User Story:** Como usuário, quero baixar o arquivo convertido no padrão DOit, para que eu possa importá-lo diretamente no sistema.

#### Acceptance Criteria

1. WHEN a conversão é concluída com sucesso, THE Gerador_Arquivo SHALL produzir o Arquivo_Convertido em formato `.xlsx` no Padrão_DOit.
2. WHEN o Arquivo_Convertido é produzido com sucesso, THE Gerador_Arquivo SHALL disponibilizar o Arquivo_Convertido para download local no navegador do usuário por meio de um controle de download visível na interface.
3. THE Gerador_Arquivo SHALL montar o Arquivo_Convertido contendo todas as colunas definidas no modelo do Padrão_DOit correspondente ao Tipo_De_Dado, na mesma ordem e com os mesmos nomes de cabeçalho do modelo.
4. WHERE o Tipo_De_Dado é Financeiro, WHEN existirem dados correspondentes a Dados Bancários, Plano de Contas ou Pendências, THE Gerador_Arquivo SHALL incluir no Arquivo_Convertido uma aba auxiliar para cada um desses conjuntos que contiver ao menos um registro.
5. WHERE o Tipo_De_Dado é Financeiro, IF não existirem dados correspondentes a uma das abas auxiliares (Dados Bancários, Plano de Contas ou Pendências), THEN THE Gerador_Arquivo SHALL omitir a aba auxiliar correspondente sem interromper a geração das demais abas.
6. THE Gerador_Arquivo SHALL gerar o Arquivo_Convertido integralmente no navegador do usuário, sem envio de dados de origem ou convertidos para servidores externos.
7. IF a geração do Arquivo_Convertido falhar, THEN THE Gerador_Arquivo SHALL exibir uma mensagem de erro indicando que o arquivo não pôde ser gerado, SHALL não disponibilizar arquivo parcial para download e SHALL preservar os dados convertidos em memória para nova tentativa.
