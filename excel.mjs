import "./assets/exceljs.mjs";
import { parse } from "./assets/csv-parse.mjs";
import { gerarCsv, gerarCsvLote, gerarCsvColecao, terrenoBasico, foraDoTotalDeck } from "./core.mjs";
import { criarResolvedorImagens } from "./imagens.mjs";
import { medir, obterMetricas } from "./cache.mjs";

async function adicionarGaleria(arquivo, planilha, itens, fimTabela, opcoes, origens) {
  if (!itens.length || opcoes.incluirImagens === false) return;
  const consultar = opcoes.resolverImagem || criarResolvedorImagens(undefined, undefined, { permitirIlustracao: planilha.name === "Colecao" });
  const imagens = new Map();
  const consultas = new Map();
  const prefixo = "LM_" + globalThis.crypto.randomUUID().replaceAll("-", "") + "_";
  const identidades = itens.map((item, indice) => prefixo + indice);
  const tabela = planilha.getTable("Tabela" + planilha.name);
  tabela.addColumn({ name: "Referencia da carta", filterButton: false }, identidades.filter((identidade, indice) => origens[indice] <= fimTabela));
  tabela.commit();
  const colunaIdentidade = planilha.columnCount;
  planilha.getColumn(colunaIdentidade).hidden = true;
  for (const [indice, origem] of origens.entries()) planilha.getCell(origem, colunaIdentidade).value = identidades[indice];
  const fimDados = Math.max(fimTabela, ...origens);
  const letraIdentidade = planilha.getColumn(colunaIdentidade).letter;
  const aba = "'" + planilha.name.replaceAll("'", "''") + "'!";
  const referencias = identidades.map((nome) => ({ nome, formula: "INDEX(" + aba + "$A$2:$A$" + fimDados
    + ',MATCH("' + nome + '",' + aba + "$" + letraIdentidade + "$2:$" + letraIdentidade + "$" + fimDados + ",0))" }));
  const iniciar = (indice) => {
    if (indice < itens.length) consultas.set(indice, Promise.resolve().then(() => consultar(itens[indice]))
      .then((resultado) => ({ resultado }), (erro) => ({ erro })));
  };
  for (let indice = 0; indice < Math.min(4, itens.length); indice++) iniciar(indice);
  const destino = (linha) => "#A" + linha;
  let linha = Math.max(500, fimDados + 20);
  planilha.getCell("A" + linha).value = "Galeria de cartas";
  planilha.getCell("A" + linha).font = { bold: true, size: 16, color: { argb: "FF245A81" } };
  linha++;
  for (const [indice, item] of itens.entries()) {
    opcoes.onProgresso?.(indice + 1, itens.length);
    const origem = origens[indice];
    const retorno = "#" + identidades[indice];
    planilha.getCell("A" + origem).value = { text: String(item.nome), hyperlink: destino(linha), tooltip: "Ver imagem da carta" };
    planilha.getCell("A" + origem).font = { color: { argb: "FF0563C1" }, underline: true };
    planilha.getCell("A" + linha).value = String(item.nome);
    planilha.getCell("A" + linha).font = { bold: true, size: 13 };
    planilha.getCell("A" + (linha + 1)).value = { text: "Voltar a carta na tabela: " + String(item.nome), hyperlink: retorno };
    planilha.getCell("A" + (linha + 1)).font = { color: { argb: "FF0563C1" }, underline: true };
    try {
      const resposta = await consultas.get(indice);
      consultas.delete(indice);
      iniciar(indice + 4);
      if (resposta.erro) throw resposta.erro;
      const resultado = resposta.resultado;
      const legenda = planilha.getCell("A" + (linha + 2));
      legenda.value = resultado.ilustrativa ? "Imagem apenas para ilustracao. Impressao original: " + resultado.referenciaOriginal
        + ". Imagem exibida: " + resultado.referencia + "."
        : resultado.referencia + (resultado.porNome ? " - imagem por nome; edicao da cotacao nao confirmada" : "");
      if (resultado.ilustrativa) {
        planilha.mergeCells(linha + 2, 1, linha + 2, Math.min(4, planilha.columnCount));
        legenda.alignment = { wrapText: true, vertical: "top" };
        legenda.font = { bold: true, color: { argb: "FF9C5700" } };
        planilha.getRow(linha + 2).height = 60;
      }
      linha += 3;
      for (const face of resultado.faces) {
        planilha.getCell("A" + linha).value = face.nome;
        if (face.erro) {
          planilha.getCell("A" + (linha + 1)).value = "Imagem indisponivel: " + face.erro;
          linha += 4;
          continue;
        }
        const chave = face.url || face.buffer;
        if (!imagens.has(chave)) imagens.set(chave, arquivo.addImage({ buffer: face.buffer, extension: face.extension }));
        for (let deslocamento = 0; deslocamento < 24; deslocamento++) planilha.getRow(linha + 1 + deslocamento).height = 15;
        planilha.addImage(imagens.get(chave), { tl: { col: 0, row: linha }, ext: { width: 320, height: 446 },
          editAs: "oneCell", hyperlinks: { hyperlink: retorno, tooltip: "Voltar a carta na tabela" } });
        linha += 27;
      }
      linha += 2;
    } catch (erro) {
      planilha.getCell("A" + (linha + 3)).value = "Imagem indisponivel: " + erro.message;
      linha += 7;
    }
  }
  return referencias;
}

async function gravarReferencias(arquivo, referencias) {
  const buffer = await arquivo.xlsx.writeBuffer();
  if (!referencias?.length) return buffer;
  const { CFB } = await import("./assets/xlsx.mjs");
  const zip = CFB.read(new Uint8Array(buffer), { type: "array" });
  const indice = zip.FullPaths.findIndex((caminho) => caminho.endsWith("xl/workbook.xml"));
  if (indice === -1) throw new Error("Estrutura do Excel indisponivel para registrar os links das cartas.");
  const entrada = zip.FileIndex[indice];
  const escapar = (texto) => texto.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;");
  const nomes = "<definedNames>" + referencias.map(({ nome, formula }) => '<definedName name="' + nome
    + '" localSheetId="0">' + escapar(formula) + "</definedName>").join("") + "</definedNames>";
  const xml = new TextDecoder().decode(entrada.content);
  entrada.content = new TextEncoder().encode(xml.replace("</sheets>", "</sheets>" + nomes));
  entrada.size = entrada.content.length;
  return CFB.write(zip, { type: "array", fileType: "zip", compression: true });
}

async function gerarPlanilha(csv, nome, itens, opcoes) {
  const inicioMetricas = obterMetricas();
  const arquivo = new globalThis.ExcelJS.Workbook();
  const planilha = arquivo.addWorksheet(nome);
  arquivo.calcProperties.fullCalcOnLoad = true;
  let linhas = parse(csv, { bom: true, delimiter: ";", skip_empty_lines: true });
  if (nome === "Colecao") {
    const campos = [
      ["Nome da carta"], ["Preco medio unitario (BRL)", "Preco minimo unitario (BRL)"], ["Subtotal estimado (BRL)"],
      ["Quantidade"], ["Acabamento", "Acabamento CSV", "Acabamento de referencia"], ["Edicao", "Edicao CSV", "Edicao de referencia"],
      ["Codigo edicao"], ["Numero"], ["Data do preco (UTC)", "Data da consulta"], ["URL da referencia"],
    ];
    const indices = campos.map((opcoes) => opcoes.map((campo) => linhas[0].indexOf(campo)).find((indice) => indice !== -1));
    linhas = [campos.map(([titulo]) => titulo), ...linhas.slice(1).map((linha) => indices.map((indice) => linha[indice] ?? ""))];
  }
  const colunas = linhas[0].map((titulo) => titulo.normalize("NFD").replace(/\p{Diacritic}/gu, "").toLowerCase());
  const moeda = (coluna) => /^(preco|subtotal)\s/.test(coluna);
  const dados = linhas.map((linha, indice) => linha.map((valor, coluna) => {
    if (indice === 0 || valor === "" || !(moeda(colunas[coluna]) || colunas[coluna] === "quantidade")) return valor;
    const numero = Number(valor.replace(",", "."));
    if (!Number.isFinite(numero)) throw new Error("Valor numerico invalido na planilha.");
    return numero;
  }));
  const colunaTipos = colunas.indexOf("tipos");
  if (colunaTipos !== -1) {
    const classificacoes = [["Creature", "01 - Criaturas"], ["Instant", "02 - Instants"], ["Sorcery", "03 - Sorceries"],
      ["Artifact", "04 - Artefatos"], ["Enchantment", "05 - Encantamentos"], ["Planeswalker", "06 - Planeswalkers"], ["Land", "07 - Terrenos"]];
    colunas.push("filtro");
    dados[0].push("Filtro");
    for (let linha = 1; linha < dados.length; linha++) {
      const referencia = "$" + planilha.getColumn(colunaTipos + 1).letter + (linha + 1);
      const formula = classificacoes.reduceRight((restante, [tipo, categoria]) =>
        'IF(ISNUMBER(SEARCH("' + tipo + '",' + referencia + ')),"' + categoria + '",' + restante + ')', '"08 - Outros"');
      const tipos = String(dados[linha][colunaTipos]).toLowerCase();
      const categoria = classificacoes.find(([tipo]) => tipos.includes(tipo.toLowerCase()))?.[1] || "08 - Outros";
      dados[linha].push({ t: "s", f: formula, v: categoria });
    }
  }
  const colunaPreco = colunas.findIndex((coluna) => coluna.startsWith("preco"));
  if (!colunas.includes("quantidade")) {
    colunas.push("data do preco (utc)", "quantidade");
    dados[0].push("Data do preco (UTC)", "Quantidade");
    for (let indice = 1; indice < dados.length; indice++) {
      dados[indice].push("", typeof dados[indice][colunaPreco] === "number" ? itens[indice - 1].quantidade ?? 1 : "");
    }
  }
  const colunaData = colunas.findIndex((coluna) => ["data da consulta", "data do preco (utc)"].includes(coluna));
  if (colunaData !== -1) {
    dados[0][colunaData] = "Data do preco (UTC)";
    for (const [indice, item] of itens.entries()) {
      const instante = item.capturadoEm || item.cachePrecoEm;
      const data = instante !== undefined && instante !== null ? new Date(instante) : null;
      dados[indice + 1][colunaData] = typeof dados[indice + 1][colunaPreco] === "number" && data && Number.isFinite(data.getTime()) ? data : "";
    }
  }
  const comandantes = opcoes.comandantes || [];
  const excluidas = itens.map((item) => nome !== "Colecao" && foraDoTotalDeck(item, comandantes));
  const principais = itens.flatMap((item, indice) => excluidas[indice] ? [] : [indice]);
  const separadas = itens.flatMap((item, indice) => excluidas[indice] ? [indice] : []);
  separadas.sort((primeira, segunda) => Number(terrenoBasico(itens[primeira].nome)) - Number(terrenoBasico(itens[segunda].nome)));
  const fimTabela = principais.length + 1;
  const linhaTotal = Math.max(100, fimTabela + 1);
  const origens = [];
  for (const [posicao, indice] of principais.entries()) origens[indice] = posicao + 2;
  for (const [posicao, indice] of separadas.entries()) origens[indice] = linhaTotal + 2 + posicao;
  const valoresLinha = (indice) => dados[indice + 1].map((valor) => valor && typeof valor === "object" && "f" in valor
    ? { formula: valor.f.replaceAll("$" + planilha.getColumn(colunaTipos + 1).letter + (indice + 2),
      "$" + planilha.getColumn(colunaTipos + 1).letter + origens[indice]), result: valor.v } : valor);
  planilha.addTable({
    name: "Tabela" + nome,
    ref: "A1",
    headerRow: true,
    totalsRow: false,
    style: { theme: "TableStyleMedium2", showRowStripes: true, showColumnStripes: false },
    columns: dados[0].map((titulo) => ({ name: titulo, filterButton: true })),
    rows: principais.map(valoresLinha),
  });
  for (const indice of separadas) planilha.getRow(origens[indice]).values = valoresLinha(indice);
  for (const [indice, coluna] of colunas.entries()) {
    const formato = planilha.getColumn(indice + 1);
    formato.width = Math.min(60, Math.max(12, ...dados.map((linha) => String(linha[indice]?.v ?? linha[indice] ?? "").length + 2)));
    if (moeda(coluna)) formato.numFmt = '"R$" #,##0.00';
    if (indice === colunaData) { formato.numFmt = "dd/mm/yyyy hh:mm:ss"; formato.width = 22; }
  }
  const colunaSubtotal = colunas.findIndex((coluna) => coluna.startsWith("subtotal"));
  const colunaValor = colunaSubtotal === -1 ? colunaPreco : colunaSubtotal;
  const colunaQuantidade = colunas.indexOf("quantidade");
  const intervalo = (coluna) => planilha.getColumn(coluna + 1).letter + "2:" + planilha.getColumn(coluna + 1).letter + fimTabela;
  const formulaTotal = !principais.length ? "0" : colunaSubtotal === -1
    ? "SUMPRODUCT(" + intervalo(colunaValor) + "," + intervalo(colunaQuantidade) + ")" : "SUM(" + intervalo(colunaValor) + ")";
  const total = principais.map((indice) => dados[indice + 1]).reduce((soma, linha) => soma + Math.round(Number(linha[colunaValor] || 0) * 100)
    * (colunaSubtotal === -1 ? Number(linha[colunaQuantidade] || 0) : 1), 0) / 100;
  const rotuloTotal = planilha.getCell("A" + linhaTotal);
  rotuloTotal.value = "Total Valor:";
  const valorTotal = planilha.getCell("B" + linhaTotal);
  valorTotal.value = { formula: formulaTotal, result: total };
  valorTotal.numFmt = '"R$" * #,##0.00';
  valorTotal.alignment = { horizontal: "right" };
  for (const celula of [rotuloTotal, valorTotal]) {
    celula.font = { bold: true };
    celula.border = { top: { style: "thin", color: { argb: "FF8EB4E3" } }, bottom: { style: "thin", color: { argb: "FF8EB4E3" } } };
  }
  planilha.views = [{ state: "frozen", ySplit: 1 }];
  const referencias = await medir("galeria", () => adicionarGaleria(arquivo, planilha, itens, fimTabela, opcoes, origens));
  const buffer = await medir("montagemExcel", () => gravarReferencias(arquivo, referencias));
  await opcoes.onMetricas?.(obterMetricas(inicioMetricas));
  return new Uint8Array(buffer);
}

export function gerarExcel(capturas, opcoes = {}) {
  return gerarPlanilha(gerarCsv(capturas), "Capturas", capturas, opcoes);
}

export function gerarExcelLote(lote, opcoes = {}) {
  return gerarPlanilha(gerarCsvLote(lote), "Lista", lote.itens.map((item) => ({
    ...(item.status === "ok" ? item.cotacao : {}), nome: item.nome, quantidade: item.quantidade ?? item.cotacao?.quantidade ?? 1,
  })), { ...opcoes, comandantes: lote.comandantes || opcoes.comandantes || [] });
}

export function gerarExcelColecao(lote, opcoes = {}) {
  return gerarPlanilha(gerarCsvColecao(lote), "Colecao", lote.itens.map((item) => ({
    ...item, ...(item.status === "ok" ? item.cotacao : {}), nome: item.nome,
    quantidade: item.quantidade,
    codigoEdicao: item.codigoEdicao || item.cotacao?.codigoEdicao,
    numero: item.numero || item.cotacao?.numero,
    edicao: item.edicao || item.cotacao?.edicao,
  })), opcoes);
}