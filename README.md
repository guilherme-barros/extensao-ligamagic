# LigaMagic - Captura Local

Extensão para Google Chrome que captura cotações de cartas de **Magic: The Gathering** na LigaMagic, consulta listas de forma assistida, estima coleções importadas de CSV e gera relatórios Excel com imagens e metadados do Scryfall.

**Projeto independente e não oficial.** Não possui vínculo com a LigaMagic, o Scryfall ou a Wizards of the Coast. Os valores são referências de consulta, não ofertas de compra ou garantia de disponibilidade.

## Índice

- [Visão geral](#visão-geral)
- [Tecnologias](#tecnologias)
- [Instalação](#instalação)
- [Funcionalidades](#funcionalidades)
- [Como usar](#como-usar)
- [Formatos de entrada](#formatos-de-entrada)
- [Exportação Excel](#exportação-excel)
- [Memória e cache](#memória-e-cache)
- [Permissões e privacidade](#permissões-e-privacidade)
- [Estrutura do projeto](#estrutura-do-projeto)
- [Desenvolvimento e testes](#desenvolvimento-e-testes)
- [Limitações](#limitações)
- [Licenças](#licenças)

## Visão geral

| Modo | Entrada | Finalidade |
| --- | --- | --- |
| Carta aberta | Página de uma carta na LigaMagic | Capturar uma cotação, ajustar a quantidade e montar uma lista manual. |
| Lista de cartas | Texto, arquivo TXT ou CSV compatível com ManaBox | Consultar cartas por nome, acompanhar pendências e calcular o total do deck. |
| Consulta de coleção | CSV com edição, número, acabamento e idioma | Estimar uma coleção mantendo a identidade de suas variantes. |

O popup oferece captura, consulta, acompanhamento e exportação. Um service worker controla as consultas de listas e os alertas em segundo plano. As capturas, os rascunhos e o progresso são armazenados localmente no navegador.

## Tecnologias

| Tecnologia | Uso no projeto |
| --- | --- |
| JavaScript com ES Modules | Implementação modular em arquivos `.mjs`, com operações assíncronas e tratamento de erros. |
| Chrome Extensions Manifest V3 | Configuração da extensão, popup, service worker e permissões. |
| HTML5 e CSS3 | Interface do popup, controles de consulta, relatórios de progresso e alertas acessíveis. |
| Chrome Extension APIs | Acesso à aba ativa, execução de scripts, controle de abas auxiliares, mensagens, alarmes, armazenamento e informações de RAM. |
| IndexedDB | Cache persistente de preços, metadados, categorias e imagens. |
| Fetch API | Consultas ao Scryfall e download de imagens, com limites de tempo e repetição de requisições transitórias. |
| Scryfall REST API | Identificação de cartas, edições, números de coleção, tipos e faces. |
| Scryfall Tagger / GraphQL | Consulta e filtragem das categorias funcionais das cartas. |
| ExcelJS | Criação de arquivos XLSX com tabelas, fórmulas, formatos monetários e imagens incorporadas. |
| SheetJS / CFB | Manipulação do pacote XLSX para registrar referências que sustentam os links internos da galeria. |
| csv-parse | Leitura estruturada de CSV, incluindo delimitadores, campos entre aspas e BOM. |
| Lucide | Ícones da interface distribuídos como arquivos SVG locais. |
| Node.js, `node:test` e `node:assert/strict` | Testes automatizados de regras de negócio, consultas simuladas, cache, exportação e interface. |

As bibliotecas usadas pela extensão estão incluídas em [assets](assets). Não é necessário instalar pacotes com npm, compilar o projeto ou executar um servidor para usá-la.

## Instalação

### Requisitos

- Google Chrome **120 ou superior**, conforme [manifest.json](manifest.json).
- A pasta completa do projeto, incluindo os arquivos de [assets](assets).
- Conexão com a internet para novas consultas à LigaMagic e ao Scryfall.

### Carregar no Chrome

1. Baixe ou clone o projeto. Se o download estiver compactado, extraia-o.
2. Abra `chrome://extensions`.
3. Ative **Modo do desenvolvedor**.
4. Clique em **Carregar sem compactação**.
5. Selecione a pasta que contém [manifest.json](manifest.json).
6. Fixe a extensão na barra do navegador para facilitar o acesso ao popup.

As permissões para os domínios externos são solicitadas quando uma operação precisa delas. Ao atualizar os arquivos, use **Recarregar** na página de extensões; alterações de permissões podem exigir nova autorização.

Abrir [popup.html](popup.html) como uma página comum não substitui essa instalação: as APIs `chrome.*` dependem do contexto de extensão.

## Funcionalidades

### Captura manual

- Lê o nome e seleciona o menor preço verde positivo entre as edições e variantes retornadas pela página da LigaMagic, registrando a referência e a data da consulta.
- Permite restringir a cotação com **Somente Normal** ou considerar as variantes disponíveis na página.
- Consulta tipos e categorias funcionais no Scryfall quando o acesso é autorizado.
- Usa o nome normalizado como identidade: capturar a mesma carta novamente atualiza a cotação e mantém a quantidade, em vez de criar outra linha por edição.
- Permite alterar quantidades, remover cartas e limpar a lista com confirmação.
- Exibe preço unitário, subtotal, número de cartas distintas e quantidade total de unidades.
- Mantém as capturas ao fechar e reabrir o popup.

### Consulta assistida de listas

- Aceita cartas digitadas, arquivos TXT e CSV de coleção compatível com ManaBox.
- Agrega nomes repetidos e preserva suas quantidades.
- Consulta as cartas em uma aba da LigaMagic controlada pela extensão.
- Permite configurar o intervalo entre acessos de **3 a 300 segundos**, em passos de **0,1 segundo**, com padrão de **4 segundos**.
- Oferece **pausar**, **retomar**, **pular a carta atual**, **abrir a aba da lista**, **cancelar** e **retentar pendências**.
- Mostra progresso, carta atual, total parcial, falhas e indicadores de reutilização do cache.
- Preserva cartas já cotadas quando são retomadas somente as pendências.
- Permite exportar o resultado parcial, sem exigir que todas as cartas tenham uma cotação.
- Continua o processamento em segundo plano enquanto o navegador permanece aberto, sem depender de manter o popup visível.

O progresso é persistido localmente. Ao reiniciar o navegador ou a extensão, uma consulta em andamento é pausada para que o usuário a retome quando desejar. Fechar a aba da lista também pausa o processamento; **Retomar** pode abrir uma nova aba.

| Controle | Comportamento |
| --- | --- |
| Pausar | Interrompe o avanço da lista, preservando o resultado obtido. |
| Retomar | Reprocessa a carta atual de uma lista pausada. |
| Pular | Disponível com a lista pausada; marca a carta atual como ignorada e avança. |
| Retentar pendências | Em listas pausadas, concluídas ou canceladas, reconsulta apenas itens que ainda não estão válidos. |
| Abrir aba da lista | Ativa a aba controlada pela extensão para acompanhar ou concluir uma verificação. |
| Cancelar | Encerra as consultas restantes, sem apagar os resultados anteriores. |

Verificações de segurança, páginas indisponíveis ou referências ambíguas podem exigir intervenção na aba da lista antes da retomada.

**Antes de iniciar uma lista ou estimativa de coleção, exporte as capturas que deseja manter.** Iniciar um novo processamento substitui as capturas atuais. Limpar a lista também remove seus resultados e rascunhos; durante uma consulta, pause-a antes de limpar.

### Comandante e parceiros

O campo **Comandante / parceiros** aceita um comandante ou dois parceiros separados por `+`. Eles são incluídos na consulta sem duplicar nomes que já estejam na lista.

No total de deck identificado como **Commander 500**, os comandantes informados e os terrenos básicos ficam fora da soma. Na planilha, são mantidos em uma seção separada abaixo do total. Essa regra não se aplica ao total do modo coleção e não representa uma validação completa de legalidade do deck.

### Estimativa de coleção por variante

Com **Consulta de coleção** ativada, a extensão usa as informações do CSV para distinguir edição, número de coleção, acabamento e idioma.

- Mantém variantes distintas da mesma carta separadas.
- Multiplica o preço unitário pela quantidade de cada variante.
- Usa o **preço médio amarelo da LigaMagic** para a referência de edição e acabamento selecionada.
- Confirma impressões e tratamentos pelo Scryfall para evitar associações incorretas entre edições e artes.
- Procura uma versão recente para terrenos básicos comuns; terrenos Foil e Full Art procuram preservar a referência específica antes de recorrer a uma alternativa compatível.
- Registra referências substituídas e seus motivos no resultado da consulta.
- Informa quantas unidades foram avaliadas e mantém o total parcial quando há pendências.

O idioma é preservado como parte da identidade importada, mas **não funciona como filtro de preço**. A estimativa também não filtra conservação, não inclui frete e não representa oferta de recompra.

### Referências de preço

| Fluxo | Valor utilizado | Observação |
| --- | --- | --- |
| Captura manual | Menor preço verde positivo retornado pela página. | Não se limita necessariamente à edição que aparece selecionada na interface do site. |
| Lista comum | Menor preço verde por nome, respeitando **Somente Normal** quando ativado. | Não estima cada variante de um CSV separadamente; terrenos básicos usam uma referência recente. |
| Consulta de coleção | Preço médio amarelo da edição e do acabamento selecionados pelo resolvedor de referências. | Procura manter a impressão importada; alternativas compatíveis de terrenos são identificadas. |

Preços mínimos e médios têm significados diferentes e não devem ser comparados como se fossem a mesma base de avaliação. Para cartas não básicas, uma impressão sem referência compatível ou sem preço médio válido permanece pendente, em vez de receber arbitrariamente a cotação de outra edição.

### Categorias funcionais do Scryfall Tagger

Na captura manual e nas listas comuns, as categorias complementam a cotação com funções da carta. A consulta de coleção concentra-se na identificação da impressão e no preço médio, sem executar esse enriquecimento funcional.

| Família | Exemplos de classificações |
| --- | --- |
| Remoção e controle | Removal, Destroy Creature, Exile Permanent, Board Wipe, Counter Spell, Bounce to Hand e Force Sacrifice. |
| Compra e seleção | Draw Cards, Card Advantage, Scry, Surveil, Tutor / Search Library e Impulse Draw. |
| Mana e desenvolvimento | Ramp, Mana Dork, Mana Rock, Ritual, Cost Reduction, Treasure Tokens e Extra Land Drops. |
| Cemitério | Reanimate, Self-Reanimate, Return to Hand, Recursion, Self-Mill e Graveyard Hate. |
| Fichas | Create Creature Tokens, Create Artifact Tokens, tipos específicos de criaturas e Food, Clue, Blood ou Map Tokens. |
| Marcadores e gatilhos | Proliferate, Poison / Toxic, Landfall, ETB Effects, Death Triggers e Trigger Doubler. |
| Sinergias e interação | Sacrifice Outlet, Blink / Flicker, Copy Effects, Steal / Gain Control e Untap. |
| Vida e dano | Life Gain, Life Drain e Group Slug / Direct Damage. |
| Combate e proteção | Evasion, Protection, Pump / Buff e habilidades concedidas, como Gives Haste e Gives Vigilance. |
| Turnos e vitória | Phase Manipulation, Extra Turns, Extra Combat e Alternate Win Condition. |

A extensão valida a identidade Oracle e considera tags funcionais do namespace `card` com status `GOOD_STANDING`. Usa uma seleção de categorias permitidas, considera relações entre tags e remove redundâncias quando uma categoria específica já descreve a função. Cartas de duas faces podem combinar tipos e categorias de ambas as faces; supertipos como Basic e Legendary não são usados como tipos na coluna correspondente.

Falhas de rede, falta de permissão ou indisponibilidade do Tagger são apresentadas como avisos, sem descartar uma cotação de preço válida. As regras e a lista de categorias estão em [tagger.mjs](tagger.mjs).

### Controle e segurança dos dados

As quantidades e os totais são validados. Os cálculos de preço são feitos em centavos, evitando acumular erros de ponto flutuante na soma de valores monetários.

A interface bloqueia alterações incompatíveis com uma consulta em andamento. Limpar uma lista exige confirmação e não apaga arquivos Excel que já foram exportados nem o cache persistente de consultas.

## Como usar

### Capturar uma carta

1. Abra uma página de carta em `https://www.ligamagic.com.br` ou `https://ligamagic.com.br`.
2. Aguarde a página carregar e conclua qualquer verificação de segurança exigida pelo site.
3. Abra o popup e selecione **Carta aberta**.
4. Configure **Somente Normal**, se necessário, e clique em **Capturar carta**.
5. Autorize as consultas externas quando solicitado e ajuste a quantidade capturada.
6. Use **Exportar Excel** para gerar a planilha.

### Consultar um deck ou uma lista

1. Selecione **Lista de cartas** e mantenha **Consulta de coleção** desativada.
2. Digite uma carta por linha ou use **Importar CSV / TXT**.
3. Informe o comandante ou os parceiros, quando aplicável.
4. Ajuste o intervalo. Marque **Atualizar preços** para ignorar cotações de preço já armazenadas no cache.
5. Clique em **Iniciar lista** e acompanhe a aba aberta pela extensão.
6. Resolva eventuais verificações, retome ou retente as pendências e exporte a lista.

### Estimar uma coleção

1. Selecione **Lista de cartas** e ative **Consulta de coleção**.
2. Importe o CSV completo com os dados das variantes.
3. Confira o resumo de variantes, unidades e proxies excluídos.
4. Defina o intervalo e clique em **Estimar coleção**.
5. Acompanhe a estimativa parcial e os avisos de referência.
6. Use **Exportar estimativa Excel** para gerar o relatório.

## Formatos de entrada

### Lista em texto ou TXT

Use uma carta por linha. A quantidade é opcional e assume o valor `1` quando omitida.

```text
# Deck de exemplo
4 Lightning Bolt
1 Sol Ring
Eternal Witness
10 Forest
```

Também são aceitas quantidades como `4x Lightning Bolt` e `4 x Lightning Bolt`. Linhas vazias e comentários iniciados por `#` ou `//` são ignorados. A comparação de nomes considera português ou inglês conforme a página, além do nome completo ou de uma face de cartas duplas.

Para parceiros, use o campo específico do popup:

```text
Thrasios, Triton Hero + Tymna the Weaver
```

### CSV compatível com ManaBox

Para importar uma lista por nome, as colunas obrigatórias são `Name` e `Quantity`. O modo coleção exige dados completos da variante.

| Coluna | Lista por nome | Consulta de coleção | Descrição |
| --- | --- | --- | --- |
| `Name` | Obrigatória | Obrigatória | Nome da carta. |
| `Quantity` | Obrigatória | Obrigatória | Quantidade inteira de 1 a 9999. |
| `Set code` | Opcional | Obrigatória | Código da edição, como `CMM`. |
| `Set name` | Opcional | Opcional | Nome descritivo da edição. |
| `Collector number` | Opcional | Obrigatória | Número da impressão, preservado como texto. |
| `Foil` | Opcional | Obrigatória | `normal`, `foil` ou `etched`. |
| `Language` | Opcional | Obrigatória | Código de idioma, como `en` ou `pt`. |
| `Proxy` | Opcional | Opcional | Registros com `true` são excluídos da importação. |
| `Scryfall ID` | Opcional | Opcional | Identificador adicional para resolver referências. |
| `Full Art` ou `Full_Art` | Opcional | Opcional | Indicador booleano `true` ou `false`. |

Exemplo:

```csv
Name,Quantity,Set code,Set name,Collector number,Foil,Language,Proxy
Sol Ring,2,CMM,Commander Masters,410,normal,en,false
"Thrasios, Triton Hero",1,CMR,Commander Legends,513,foil,en,false
```

O parser reconhece vírgula, ponto e vírgula ou tabulação, campos entre aspas e arquivos com BOM. Linhas sem nome válido, quantidades inválidas, colunas duplicadas e CSV malformado são rejeitados com uma mensagem de erro. Preços de compra eventualmente presentes no CSV não são utilizados como cotação.

### Limites de importação

| Entrada | Limite |
| --- | --- |
| Lista digitada ou TXT | 500 cartas distintas. |
| CSV convertido em lista por nome | 5000 cartas distintas. |
| CSV no modo coleção | 5000 variantes. |
| Quantidade por carta ou variante agregada | 9999 unidades. |
| Arquivo CSV | 2 MB, equivalentes a 2.000.000 bytes. |
| Arquivo TXT | 200 KB, equivalentes a 200.000 bytes. |

## Exportação Excel

Os botões de exportação geram arquivos **XLSX reais**, e não arquivos CSV renomeados. Os relatórios usam as abas `Capturas`, `Lista` ou `Colecao`, conforme o fluxo.

- Preços e quantidades são células numéricas, com formatação monetária em reais.
- O total usa fórmulas e considera as quantidades, respeitando as exclusões do modo deck.
- A primeira linha é congelada e as tabelas oferecem filtros por coluna.
- Relatórios com tipos de carta recebem a coluna **Filtro**, calculada por fórmula, para agrupar criaturas, instants, sorceries, artefatos, encantamentos, planeswalkers, terrenos e outros.
- As datas de preço são identificadas em UTC.
- Números de coleção são preservados como texto.
- A galeria de cartas fica na mesma aba, após a área de dados, e incorpora as imagens no arquivo.
- Links permitem navegar da carta na tabela para sua imagem e voltar à referência correspondente.
- Cartas com duas faces podem receber imagens de ambas as faces.
- Falhas de imagem são registradas na galeria sem impedir a exportação das demais cartas.

A planilha de capturas ou listas comuns apresenta nome, preço mínimo, tipos, categorias principais, filtro por tipo, data e quantidade. A de coleção apresenta nome, preço médio unitário, subtotal, quantidade, acabamento, edição, código, número, data e URL de referência. Os detalhes de pendências e substituições permanecem no resultado local da consulta; nem todos os campos do CSV interno aparecem na planilha final.

Na coluna **Categorias (principal)**, o exportador usa `Land` como valor padrão quando não há categorias filtradas. Esse padrão, isoladamente, não confirma que a carta seja um terreno: confira os tipos e os avisos de consulta.

A resolução de imagens prioriza a impressão identificada. Imagens obtidas apenas por nome são sinalizadas; no relatório de coleção, uma imagem alternativa pode ser usada como ilustração após falha da referência exata, com identificação explícita. **Uma imagem ilustrativa não altera a cotação.**

A exportação da galeria exige permissão para a API e as imagens do Scryfall. Os downloads são limitados a quatro cartas em paralelo e imagens individuais acima de 10 MB são rejeitadas.

O código também possui geradores CSV usados como base interna para montar os relatórios. A interface disponibiliza a exportação em Excel, não um botão separado de download CSV.

## Memória e cache

### Avisos de memória

O popup exibe um aviso separado das mensagens de captura quando a RAM está baixa ou a memória ou o armazenamento se aproximam do limite.

| Medição | Atenção | Nível crítico |
| --- | --- | --- |
| RAM disponível no computador | Até 10% da capacidade ou até 512 MiB livres. | Até 5% da capacidade ou até 256 MiB livres. |
| Heap JavaScript do popup, quando disponível | A partir de 85% do limite. | A partir de 95% do limite. |
| Dados em `chrome.storage.local` | A partir de 85% da quota. | A partir de 95% da quota. |

A verificação ocorre a cada **5 segundos** com o popup aberto e a cada **30 segundos** em segundo plano. O ícone exibe `!` para alertas de RAM ou armazenamento, com cor de atenção ou crítica e detalhes ao passar o mouse. Quando esses recursos voltam ao normal, o ícone recupera a contagem de capturas.

Os avisos desaparecem após a recuperação. Eles são informativos: **não apagam dados nem pausam operações automaticamente**. Fechar abas ou programas e reduzir o tamanho de exportações com imagens pode aliviarar o uso de memória. Antes de limpar capturas, exporte os dados que deseja manter.

O heap depende da disponibilidade de `performance.memory`; medições indisponíveis não geram avisos falsos. O monitor não mede toda a memória de todos os processos da extensão.

### Cache persistente

- Preços, metadados, categorias e imagens têm validade de **30 dias**.
- O cache usa IndexedDB e possui limite interno de **1 GiB**.
- Entradas expiradas ou menos recentemente usadas são removidas nas gravações quando necessário para respeitar esse limite.
- Uma implementação em memória é usada como alternativa quando o banco não pode ser aberto; nesse caso, não há persistência entre sessões.
- **Atualizar preços** ignora cotações de preço em cache, mas não equivale a apagar todos os metadados e imagens armazenados.
- O limite do cache é independente da quota de `chrome.storage.local`. O alerta de armazenamento das capturas não representa a ocupação do IndexedDB.

Filas limitam a concorrência, reutilizam consultas e reduzem acessos repetidos. A extensão registra métricas de cache, navegação e exportação para diagnóstico; as métricas de exportação são registradas no console e no armazenamento local.

## Permissões e privacidade

### Permissões da extensão

| Permissão | Motivo |
| --- | --- |
| `activeTab` | Acesso temporário à aba ativa após interação do usuário. |
| `scripting` | Leitura dos dados de cartas e consulta ao Tagger por scripts executados nas páginas autorizadas. |
| `storage` | Persistência das capturas, preferências, rascunhos, progresso e métricas. |
| `alarms` | Agendamento das consultas de listas e da verificação de memória em segundo plano. |
| `system.memory` | Consulta da capacidade e da RAM disponível no computador. |

### Acesso opcional a sites

| Domínio | Finalidade |
| --- | --- |
| `ligamagic.com.br` e `www.ligamagic.com.br` | Leitura e navegação assistida nas páginas de cartas. |
| `api.scryfall.com` | Metadados, impressão, tipos e faces das cartas. |
| `tagger.scryfall.com` | Categorias funcionais. |
| `cards.scryfall.io` | Imagens usadas na galeria Excel. |

Não há backend próprio, login próprio, sincronização em nuvem ou envio de relatórios para um servidor do projeto. Os arquivos Excel são gerados no navegador e baixados no computador.

Isso não torna as consultas anônimas: nomes e identificadores das cartas são usados nas requisições e os serviços externos recebem os dados normais de conexão e de sessão do navegador. Suas políticas de acesso e privacidade continuam aplicáveis.

## Estrutura do projeto

| Arquivo ou diretório | Responsabilidade |
| --- | --- |
| [manifest.json](manifest.json) | Metadados, versão mínima do Chrome, permissões e pontos de entrada. |
| [background.mjs](background.mjs) | Eventos do navegador, mensagens, alarmes e indicador no ícone. |
| [popup.html](popup.html) | Estrutura dos controles e das áreas de resultado. |
| [popup.css](popup.css) | Aparência, layout e adaptação da interface. |
| [popup.mjs](popup.mjs) | Interação com o usuário, persistência, comandos de lista e downloads. |
| [core.mjs](core.mjs) | Captura, importação, identidade de variantes, cálculo e geração de CSV. |
| [lote.mjs](lote.mjs) | Estado e execução das consultas assistidas. |
| [tagger.mjs](tagger.mjs) | Consulta ao Scryfall Tagger e regras de classificação. |
| [imagens.mjs](imagens.mjs) | Resolução e download de imagens. |
| [excel.mjs](excel.mjs) | Montagem dos XLSX, tabelas, fórmulas e galeria. |
| [cache.mjs](cache.mjs) | Cache, filas, repetição de chamadas e métricas. |
| [memoria.mjs](memoria.mjs) | Medições e regras dos alertas de RAM, heap e armazenamento. |
| [tagger.test.mjs](tagger.test.mjs) | Suíte de testes automatizados. |
| [assets](assets) | Bibliotecas incluídas localmente, ícones e licenças de terceiros. |

## Desenvolvimento e testes

Node.js é necessário apenas para executar os testes, não para utilizar a extensão instalada. Use uma versão atual compatível com `node:test` e as APIs JavaScript usadas pelo projeto.

Na pasta raiz, execute:

```bash
node --test tagger.test.mjs
```

Para executar somente os testes do monitor de memória:

```bash
node --test --test-name-pattern="monitor de memoria" tagger.test.mjs
```

A suíte cobre importação e agregação de listas, variantes de coleção, regras de preço e referência, tipos e categorias, consultas simuladas, retomada de lotes, cache, imagens, estrutura dos XLSX e estados do popup e dos avisos de memória.

Os testes usam simulações para serviços externos e APIs do navegador. Eles não substituem a validação manual da extensão instalada contra páginas reais da LigaMagic e do Scryfall.

Após uma alteração, recarregue a extensão e confira o popup, as permissões e o service worker pela página `chrome://extensions`. Não há etapa obrigatória de build.

## Limitações

### Problemas comuns

| Situação | Ação recomendada |
| --- | --- |
| A extensão não carrega | Confira a versão do Chrome e selecione a pasta completa que contém o manifesto, não apenas um arquivo ou uma subpasta de recursos. |
| A LigaMagic exibe uma verificação ou a lista fica pausada | Abra a aba da lista, conclua a verificação, aguarde o carregamento e use **Retomar**. |
| As categorias do Tagger não aparecem | Confira as permissões, abra o Scryfall Tagger para resolver verificações e capture novamente ou refaça a consulta necessária. A cotação válida permanece preservada quando o enriquecimento falha. |
| Uma variante da coleção fica sem preço médio | Confira edição, número, acabamento e nome no CSV. A referência precisa ter preço médio válido e ser compatível com a impressão solicitada. |
| O CSV é rejeitado | Verifique as colunas obrigatórias, aspas, delimitador, valores de quantidade, indicadores booleanos e limites de arquivo. |
| Há aviso de memória ou falha em uma exportação grande | Libere RAM fechando abas ou programas e processe listas menores. Exporte dados importantes antes de limpar capturas. |

### Restrições conhecidas

- Os preços refletem a referência consultada e podem mudar; capturas em cache podem ter até 30 dias.
- O total não inclui frete e não garante estoque, disponibilidade ou valor de venda da coleção.
- O modo coleção preserva idioma, mas não filtra idioma ou conservação na cotação.
- O total Commander 500 possui regras de exclusão; a extensão não verifica todas as regras de construção de decks.
- Alterações no HTML da LigaMagic ou nas APIs do Scryfall podem exigir manutenção dos extratores.
- Verificações de segurança devem ser concluídas pelo usuário. A extensão não as contorna.
- A consulta pode ser afetada por limites de acesso, respostas HTTP `429`/`503` e falhas de rede.
- Exportações com muitas imagens podem consumir bastante memória; um aviso não garante que uma operação grande não esgote os recursos.
- O limite interno de 1 GiB do cache não garante que todo esse espaço esteja disponível no disco ou na quota do navegador.
- Outros navegadores baseados em Chromium dependem do suporte às APIs utilizadas; o Chrome é a referência de instalação declarada.
- A extensão não realiza compras, vendas ou alterações de inventário nos serviços externos.

## Licenças

As bibliotecas e os ícones de terceiros mantêm seus avisos de licença no projeto:

| Componente | Licença | Aviso incluído |
| --- | --- | --- |
| ExcelJS | MIT | [assets/EXCELJS-LICENSE](assets/EXCELJS-LICENSE) |
| SheetJS | Apache 2.0 | [assets/SHEETJS-LICENSE](assets/SHEETJS-LICENSE) |
| csv-parse | MIT | [assets/CSV-PARSE-LICENSE](assets/CSV-PARSE-LICENSE) |
| Lucide | ISC, com avisos MIT para ícones indicados no arquivo | [assets/LUCIDE-LICENSE](assets/LUCIDE-LICENSE) |

Esses avisos se aplicam aos componentes correspondentes. O repositório não contém um arquivo de licença próprio para o código da extensão; não presuma uma licença geral a partir das dependências.