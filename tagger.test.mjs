import test from "node:test";
import assert from "node:assert/strict";
import { CATEGORIAS_PERMITIDAS, filtrarCategorias, consultarTagger, enriquecerCotacao, consultarTagsNavegador, lerTagsNaPagina, permiteSacrificioBarato } from "./tagger.mjs";
import { gerarCsv, gerarCsvLote, nomesEquivalentes, conferirNome, capturarPagina } from "./core.mjs";
import { criarControlador } from "./lote.mjs";
import { gerarExcel, gerarExcelLote, gerarExcelColecao } from "./excel.mjs";
import { read as lerExcel } from "./assets/xlsx.mjs";
import { criarResolvedorImagens } from "./imagens.mjs";
import { criarCachePersistente, criarFila, consultarJsonScryfall, VALIDADE } from "./cache.mjs";
import { verificarMemoria } from "./memoria.mjs";

test("monitor de memoria respeita os limites de aviso e de nivel critico", async () => {
  const MB = 1024 * 1024;
  const casos = [
    { nome: "RAM normal", livre: 2048 * MB, quantidade: 0 },
    { nome: "RAM acima de 10%", livre: 820 * MB, quantidade: 0 },
    { nome: "RAM em 10%", livre: 819.2 * MB, quantidade: 1 },
    { nome: "RAM em 5%", livre: 409.6 * MB, quantidade: 1, critico: true },
    { nome: "RAM em 512 MB", capacidade: 4096 * MB, livre: 512 * MB, quantidade: 1 },
    { nome: "RAM em 256 MB", capacidade: 2048 * MB, livre: 256 * MB, quantidade: 1, critico: true },
    { nome: "RAM esgotada", livre: 0, quantidade: 1, critico: true },
    { nome: "heap abaixo de 85%", heap: 849, quantidade: 0 },
    { nome: "heap em 85%", heap: 850, quantidade: 1 },
    { nome: "heap em 95%", heap: 950, quantidade: 1, critico: true },
    { nome: "armazenamento abaixo de 85%", salvos: 849, quantidade: 0 },
    { nome: "armazenamento em 85%", salvos: 850, quantidade: 1 },
    { nome: "armazenamento em 95%", salvos: 950, quantidade: 1, critico: true },
    { nome: "todos os avisos", livre: 0, heap: 950, salvos: 950, quantidade: 3, critico: true },
  ];
  for (const caso of casos) {
    const resultado = await verificarMemoria({
      sistema: { getInfo: async () => ({ capacity: caso.capacidade ?? 8192 * MB, availableCapacity: caso.livre ?? 2048 * MB }) },
      armazenamento: { QUOTA_BYTES: 1000, getBytesInUse: async () => caso.salvos ?? 0 },
      desempenho: { memory: { jsHeapSizeLimit: 1000, usedJSHeapSize: caso.heap ?? 0 } },
    });
    assert.equal(resultado.avisos.length, caso.quantidade, caso.nome);
    assert.equal(resultado.critico, Boolean(caso.critico), caso.nome);
  }
});

test("monitor de memoria ignora APIs indisponiveis e medicoes invalidas", async () => {
  assert.deepEqual(await verificarMemoria({ sistema: null, armazenamento: null, desempenho: null }), { avisos: [], critico: false });
  for (const invalido of [undefined, NaN, -1, Infinity]) {
    assert.deepEqual(await verificarMemoria({
      sistema: { getInfo: async () => ({ capacity: 1000, availableCapacity: invalido }) },
      armazenamento: { QUOTA_BYTES: 1000, getBytesInUse: async () => invalido },
      desempenho: { memory: { jsHeapSizeLimit: 1000, usedJSHeapSize: invalido } },
    }), { avisos: [], critico: false });
  }
  assert.deepEqual(await verificarMemoria({
    sistema: { getInfo: async () => ({ capacity: 0, availableCapacity: 0 }) },
    armazenamento: { QUOTA_BYTES: 0, getBytesInUse: async () => 0 },
    desempenho: { memory: { jsHeapSizeLimit: 0, usedJSHeapSize: 0 } },
  }), { avisos: [], critico: false });
});

test("monitor de memoria mantem outras medicoes quando uma API falha", async () => {
  const falhar = async () => { throw new Error("API indisponivel"); };
  const resultado = await verificarMemoria({
    sistema: { getInfo: falhar },
    armazenamento: { QUOTA_BYTES: 1000, getBytesInUse: falhar },
    desempenho: { memory: { jsHeapSizeLimit: 1000, usedJSHeapSize: 950 } },
  });
  assert.equal(resultado.avisos.length, 1);
  assert.match(resultado.avisos[0], /Memoria da extensao/);
  assert.equal(resultado.critico, true);
});

test("comandante e parceiros entram na consulta sem duplicar cartas ja presentes", async () => {
  const { prepararListaDeck } = await import("./core.mjs");
  const deck = prepararListaDeck("Sol Ring\nGrist, Voracious Larva // Grist, the Plague Swarm\n10 Forest", "Grist, Voracious Larva");
  assert.equal(deck.itens.length, 3);
  assert.deepEqual(deck.comandantes, ["Grist, Voracious Larva"]);
  assert.deepEqual(prepararListaDeck("Sol Ring", "Thrasios, Triton Hero + Tymna the Weaver").itens.map((item) => item.nome),
    ["Sol Ring", "Thrasios, Triton Hero", "Tymna the Weaver"]);
  assert.equal(prepararListaDeck("", "Grist, Voracious Larva").itens.length, 1);
  assert.throws(() => prepararListaDeck("Sol Ring", "2 Grist, Voracious Larva"), /uma copia/);
  assert.throws(() => prepararListaDeck("2 Grist, Voracious Larva", "Grist, Voracious Larva"), /unica copia/);
  assert.throws(() => prepararListaDeck("Sol Ring", "Thrasios + Tymna + Grist"), /dois parceiros/);
});

test("Excel de deck separa comandante e terrenos basicos depois do total sem soma-los", async () => {
  const cartas = [
    { nome: "Grist, Voracious Larva // Grist, the Plague Swarm", centavos: 799, quantidade: 1, tiposCarta: ["Creature", "Planeswalker"] },
    { nome: "Sol Ring", centavos: 200, quantidade: 2, tiposCarta: ["Artifact"] },
    { nome: "Forest", centavos: 25, quantidade: 10, tiposCarta: ["Land"] },
    { nome: "Swamp", centavos: 23, quantidade: 5, tiposCarta: ["Land"] },
    { nome: "Tranquil Cove", centavos: 56, quantidade: 1, tiposCarta: ["Land"] },
  ];
  const arquivo = new globalThis.ExcelJS.Workbook();
  await arquivo.xlsx.load(await gerarExcel(cartas, { incluirImagens: false, comandantes: ["Grist, Voracious Larva"] }));
  const folha = arquivo.getWorksheet("Capturas");
  assert.equal(folha.getCell("A2").value, "Sol Ring");
  assert.equal(folha.getCell("A3").value, "Tranquil Cove");
  assert.equal(folha.getCell("A4").value, null);
  assert.equal(folha.getCell("A100").value, "Total Valor:");
  assert.equal(folha.getCell("B100").value.result, 4.56);
  assert.equal(folha.getCell("A102").value, cartas[0].nome);
  assert.equal(folha.getCell("A103").value, "Forest");
  assert.equal(folha.getCell("A104").value, "Swamp");
  assert.ok(folha.getCell("E102").value.formula.includes("$C102"));
  const bytesGaleria = await gerarExcel(cartas, { comandantes: ["Grist, Voracious Larva"], resolverImagem: async () => ({ referencia: "Teste", faces: [] }) });
  const galeria = new globalThis.ExcelJS.Workbook();
  await galeria.xlsx.load(bytesGaleria);
  const folhaGaleria = galeria.getWorksheet("Capturas");
  assert.equal(folhaGaleria.getTable("TabelaCapturas").table.tableRef, "A1:H3");
  assert.equal(folhaGaleria.getCell("A102").value.hyperlink, "#A501");
  assert.equal(folhaGaleria.getCell("A2").value.hyperlink, "#A506");
  assert.equal(folhaGaleria.getCell("A103").value.hyperlink, "#A511");
  assert.equal(folhaGaleria.getCell("A104").value.hyperlink, "#A516");
  const nomes = lerExcel(bytesGaleria, { type: "array" }).Workbook.Names;
  assert.equal(nomes.length, cartas.length);
  assert.ok(nomes.every((nome) => nome.Ref.includes("$A$2:$A$104")));
  const apenasExcluidas = new globalThis.ExcelJS.Workbook();
  const bytesExcluidas = await gerarExcelLote({ comandantes: ["Grist, Voracious Larva"], itens: [cartas[0], cartas[2]]
    .map((cotacao) => ({ nome: cotacao.nome, quantidade: cotacao.quantidade, status: "ok", cotacao })) },
  { resolverImagem: async () => ({ referencia: "Teste", faces: [] }) });
  await apenasExcluidas.xlsx.load(bytesExcluidas);
  const folhaExcluidas = apenasExcluidas.getWorksheet("Lista");
  assert.equal(folhaExcluidas.getCell("B100").value.formula, "0");
  assert.equal(lerExcel(bytesExcluidas, { type: "array" }).Sheets.Lista.B100.v, 0);
  assert.equal(folhaExcluidas.getCell("A102").value.text, cartas[0].nome);
  assert.equal(folhaExcluidas.getCell("A102").value.hyperlink, "#A501");
  assert.equal(folhaExcluidas.getCell("A103").value.text, "Forest");
  assert.match(folhaExcluidas.getTable("TabelaLista").table.tableRef, /^A1:[A-Z]+1$/);
});

test("todos os tipos de cache expiram exatamente em 30 dias", async () => {
  let agora = 0;
  const cache = criarCachePersistente({ banco: null, agora: () => agora });
  for (const [tipo, prazo] of Object.entries(VALIDADE)) {
    assert.equal(prazo, 30 * 24 * 60 * 60 * 1000, tipo);
    await cache.set(tipo, { salvo: true }, prazo);
  }
  agora = 30 * 24 * 60 * 60 * 1000 - 1;
  for (const tipo of Object.keys(VALIDADE)) assert.deepEqual(await cache.get(tipo), { salvo: true });
  agora++;
  for (const tipo of Object.keys(VALIDADE)) assert.equal(await cache.get(tipo), null);
});

test("cache tem validade limite LRU e copia isolada, fila limita concorrencia e Scryfall respeita retry", async () => {
  let agora = 0;
  const cache = criarCachePersistente({ banco: null, agora: () => agora, limiteBytes: 6 });
  await cache.set("a", new Uint8Array([1, 2, 3]), 100);
  agora = 1;
  await cache.set("b", new Uint8Array([4, 5, 6]), 100);
  agora = 2;
  const copia = await cache.get("a"); copia[0] = 9;
  assert.equal((await cache.get("a"))[0], 1);
  agora = 3;
  await cache.set("c", new Uint8Array([7, 8, 9]), 100);
  assert.equal(await cache.get("b"), null);
  agora = 200;
  assert.equal(await cache.get("a"), null);
  const fila = criarFila(4);
  let ativas = 0;
  let maximo = 0;
  await Promise.all(Array.from({ length: 10 }, () => fila(async () => {
    ativas++; maximo = Math.max(maximo, ativas); await new Promise((resolver) => setImmediate(resolver)); ativas--;
  })));
  assert.equal(maximo, 4);
  const esperas = [];
  let chamadas = 0;
  assert.deepEqual(await consultarJsonScryfall("https://api.scryfall.com/cards/teste", async () => ++chamadas === 1
    ? { ok: false, status: 429, headers: { get: () => "2" } } : { ok: true, json: async () => ({ nome: "ok" }) }, async (ms) => esperas.push(ms)), { nome: "ok" });
  assert.deepEqual(esperas, [110, 2000, 110]);
});

test("categorias sao reutilizadas pela identidade Oracle com validade independente da impressao", async () => {
  const cache = criarCachePersistente({ banco: null });
  const mock = simular();
  let consultas = 0;
  const consultar = async (...argumentos) => { consultas++; return mock.consultarTags(...argumentos); };
  await consultarTagger("Sol Ring", mock.requisitar, consultar, { cache });
  const dados = await consultarTagger("Sol Ring", mock.requisitar, consultar, { cache });
  assert.equal(consultas, 1);
  assert.equal(dados.statusTagger, "ok");
  assert.ok(dados.categoriasTagger.includes("Mana Rock"));
});

test("imagens persistem entre resolvedores e galeria baixa quatro cartas em paralelo sem reordenar", async () => {
  const cache = criarCachePersistente({ banco: null });
  let downloads = 0;
  const requisitar = async (url) => url.startsWith("https://cards.scryfall.io/")
    ? (downloads++, { ok: true, arrayBuffer: async () => new Uint8Array([255, 216, 255, 0]).buffer })
    : { ok: true, json: async () => ({ name: "Swamp", set: "eld", collector_number: "261", image_uris: { normal: "https://cards.scryfall.io/teste.jpg" } }) };
  for (let tentativa = 0; tentativa < 2; tentativa++) await criarResolvedorImagens(requisitar, async () => {}, { cache })({ nome: "Swamp" });
  assert.equal(downloads, 1);
  let ativas = 0;
  let maximo = 0;
  const cartas = Array.from({ length: 8 }, (_, indice) => ({ nome: "Carta " + indice, centavos: 100 }));
  const bytes = await gerarExcel(cartas, { resolverImagem: async () => {
    ativas++; maximo = Math.max(maximo, ativas); await new Promise((resolver) => setImmediate(resolver)); ativas--;
    throw new Error("Sem imagem");
  } });
  assert.equal(maximo, 4);
  const arquivo = new globalThis.ExcelJS.Workbook(); await arquivo.xlsx.load(bytes);
  for (const [indice, carta] of cartas.entries()) assert.equal(arquivo.getWorksheet("Capturas").getCell("A" + (indice + 2)).value.text, carta.nome);
});

test("lote respeita intervalo escolhido e padrao de 4 segundos sem espera extra em consultas demoradas", async () => {
  let tempo = 0;
  let url;
  const estado = {};
  const navegacoes = [];
  const cache = criarCachePersistente({ banco: null, agora: () => tempo });
  const api = {
    storage: { local: { get: async (chave) => structuredClone({ [chave]: estado[chave] }), set: async (dados) => Object.assign(estado, structuredClone(dados)) } },
    permissions: { contains: async () => true },
    tabs: { create: async () => ({ id: 7 }), get: async () => ({ status: "complete", url }), update: async (id, dados) => {
      url = dados.url;
      const nome = new URL(url).searchParams.get("card");
      if (nome === "Carta C") {
        assert.equal(estado.lote.itens[0].status, "ok");
        assert.equal(estado.lote.itens[1].status, "ok");
        assert.equal(estado.capturas.find((item) => item.nome === "Carta A").centavos, 100);
      }
      navegacoes.push(nome);
    } },
    scripting: { executeScript: async () => [{ result: { nome: new URL(url).searchParams.get("card"), url, edicoes: [{ code: "set", num: "1", name: "Set", price: { 0: { p: "1.00" } } }] } }] },
    alarms: { create: async () => {}, clear: async () => {} },
  };
  const controle = criarControlador(api, () => tempo, async (cotacao) => ({ ...cotacao, statusTagger: "ok" }), undefined, { cache });
  for (const intervalo of [NaN, Infinity, -1, 2.9, 301]) {
    await assert.rejects(controle.comando("iniciar", { texto: "Carta B", intervalo }), /Intervalo/);
  }
  assert.deepEqual(navegacoes, []);
  await controle.comando("iniciar", { texto: "Carta B" }); await controle.processar(7);
  assert.equal(estado.lote.intervalo, 4);
  await controle.comando("iniciar", { texto: "Carta A\n3 Carta B\nCarta C\nCarta D", intervalo: 8 }); await controle.processar(7);
  assert.equal(estado.lote.intervalo, 8);
  assert.equal(estado.lote.indice, 2);
  assert.equal(estado.lote.itens[1].status, "ok");
  assert.equal(estado.lote.metricas.cartasCache, 1);
  assert.equal(estado.capturas.find((item) => item.nome === "Carta B").quantidade, 3);
  assert.deepEqual(navegacoes, ["Carta B", "Carta A"]);
  assert.equal(tempo, 0);
  assert.equal(estado.lote.proxima, 8000);
  tempo = 7999;
  const retomado = criarControlador(api, () => tempo, async (cotacao) => ({ ...cotacao, statusTagger: "ok" }), undefined, { cache });
  await retomado.processar();
  assert.deepEqual(navegacoes, ["Carta B", "Carta A"]);
  tempo = 8000;
  await retomado.processar();
  assert.deepEqual(navegacoes, ["Carta B", "Carta A", "Carta C"]);
  tempo = 16100;
  await retomado.processar(7);
  assert.deepEqual(navegacoes, ["Carta B", "Carta A", "Carta C", "Carta D"]);
  await retomado.processar(7);
  assert.equal(estado.lote.fase, "concluido");
  await controle.comando("iniciar", { texto: "Carta B", atualizarPrecos: true });
  assert.deepEqual(estado.capturas, []);
  assert.equal(navegacoes.length, 5);
  await controle.processar(7);
  const atualizadoEm = tempo;
  tempo = atualizadoEm + VALIDADE.precos - 1;
  await controle.comando("iniciar", { texto: "Carta B" });
  assert.equal(estado.lote.fase, "concluido");
  assert.equal(navegacoes.length, 5);
  assert.deepEqual(estado.capturas.map((item) => item.nome), ["Carta B"]);
  const arquivo = new globalThis.ExcelJS.Workbook();
  await arquivo.xlsx.load(await gerarExcel(estado.capturas, { incluirImagens: false }));
  assert.equal(arquivo.getWorksheet("Capturas").getCell("A2").value, "Carta B");
  assert.equal(arquivo.getWorksheet("Capturas").getCell("A3").value, null);
  tempo = atualizadoEm + VALIDADE.precos + 1;
  await controle.comando("iniciar", { texto: "Carta B" });
  assert.equal(navegacoes.length, 6);
});

test("lote reutiliza uma pagina entre normal e foil e antecipa as imagens durante a consulta", async () => {
  const { prepararColecao } = await import("./core.mjs");
  const estado = {};
  const cache = criarCachePersistente({ banco: null, agora: () => 0 });
  let url;
  let navegacoes = 0;
  let leituras = 0;
  const verificadas = [];
  const antecipadas = [];
  const api = {
    storage: { local: { get: async (chave) => structuredClone({ [chave]: estado[chave] }), set: async (dados) => Object.assign(estado, structuredClone(dados)) } },
    permissions: { contains: async () => true },
    tabs: { create: async () => ({ id: 7 }), get: async () => ({ status: "complete", url }), update: async (id, dados) => { url = dados.url; navegacoes++; } },
    scripting: { executeScript: async () => { leituras++; return [{ result: { nome: "Meteor Golem", url,
      edicoes: [{ code: "fdn", num: "256", name: "Foundations", price: { 0: { m: "0.71" }, 2: { m: "1.33" } } }], extras: [{ id: 2, label: "Foil" }] } }]; } },
    alarms: { create: async () => {}, clear: async () => {} },
  };
  const controle = criarControlador(api, () => 0, undefined, async (item) => {
    verificadas.push(item.acabamento);
    return { fullArt: false, tratamento: "padrao", referenciasTerreno: [] };
  }, {
    cache,
    anteciparImagem: async (item) => { antecipadas.push(item); },
  });
  const variante = { nome: "Meteor Golem", quantidade: 2, codigoEdicao: "fdn", edicao: "Foundations", numero: "256", idioma: "en",
    fullArt: true, tratamento: "fullart", referenciasTerreno: [] };
  const [antiga] = prepararColecao([{ ...variante, acabamento: "normal" }]);
  await cache.set("cotacao:" + JSON.stringify(["meteor golem", antiga.idColecao]),
    { id: antiga.idColecao, nome: variante.nome, centavos: 9999, cachePrecoEm: 0 }, 900000);
  await controle.comando("iniciar", { colecao: { criterios: "variantes", variantes: [{ ...variante, acabamento: "normal" }, { ...variante, acabamento: "foil" }] } });
  await controle.processar(7);
  await new Promise((resolver) => setImmediate(resolver));
  assert.equal(estado.lote.fase, "concluido");
  assert.equal(navegacoes, 1);
  assert.equal(leituras, 1);
  assert.equal(estado.lote.metricas.cartasCache, 1);
  assert.deepEqual(estado.lote.itens.map((item) => item.cotacao.centavos), [71, 133]);
  assert.equal(antecipadas.length, 2);
  assert.deepEqual(antecipadas.map((item) => item.numero), ["256", "256"]);
  assert.deepEqual(verificadas, ["normal", "foil"]);
  assert.ok(estado.lote.itens.every((item) => item.fullArt === false && item.cotacao.verificacaoScryfall === item.verificacaoScryfall));
  await controle.comando("iniciar", { colecao: { criterios: "variantes", variantes: [{ ...variante, acabamento: "normal" }] } });
  assert.equal(estado.lote.fase, "concluido");
  assert.deepEqual(verificadas, ["normal", "foil", "normal"]);
  assert.equal(navegacoes, 1);
});

test("imagens usam edicao numero, cache e ambas as faces sem trocar a impressao", async () => {
  const chamadas = [];
  const consultar = criarResolvedorImagens(async (url) => {
    chamadas.push(url);
    if (url.startsWith("https://cards.scryfall.io/")) return { ok: true, arrayBuffer: async () => new Uint8Array([255, 216, 255, 0]).buffer };
    return { ok: true, json: async () => ({ name: "Delver of Secrets // Insectile Aberration", set: "isd", collector_number: "51", card_faces: [
      { name: "Delver of Secrets", image_uris: { normal: "https://cards.scryfall.io/front.jpg" } },
      { name: "Insectile Aberration", image_uris: { normal: "https://cards.scryfall.io/back.jpg" } },
    ] }) };
  }, async () => {});
  const item = { nome: "Delver of Secrets", codigoEdicao: "isd", numero: "051" };
  const imagens = await consultar(item);
  assert.equal(chamadas[0], "https://api.scryfall.com/cards/isd/51");
  assert.equal(imagens.faces.length, 2);
  assert.equal(imagens.faces[0].extension, "jpeg");
  await consultar(item);
  assert.equal(chamadas.length, 3);
  await assert.rejects(consultar({ ...item, numero: "52" }), /edicao e ao numero/);
  await assert.rejects(consultar({ ...item, nome: "Outra carta" }), /nome da carta/);
});

test("imagem ilustrativa da colecao so aparece apos falha da impressao exata e informa a original", async () => {
  const item = { nome: "Quicksilver Fisher", codigoEdicao: "one", numero: "287", edicao: "Phyrexia: All Will Be One", acabamento: "normal" };
  const exata = { name: item.nome, set: "one", collector_number: "287", image_uris: { normal: "https://cards.scryfall.io/exata.jpg" } };
  const comum = { name: item.nome, set: "one", collector_number: "66", finishes: ["nonfoil"], full_art: false,
    image_uris: { normal: "https://cards.scryfall.io/comum.jpg" } };
  for (const falha of ["metadata", "download", "nenhuma"]) {
    const chamadas = [];
    const requisitar = async (url) => {
      chamadas.push(url);
      if (url.includes("/search?")) return { ok: true, json: async () => ({ data: [comum] }) };
      if (url.endsWith("/one/287")) return falha === "metadata" ? { ok: false, status: 404 }
        : { ok: true, json: async () => exata };
      if (url.endsWith("exata.jpg") && falha === "download") return { ok: false, status: 404 };
      return { ok: true, arrayBuffer: async () => new Uint8Array([255, 216, 255, 0]).buffer };
    };
    const resolver = criarResolvedorImagens(requisitar, async () => {}, { permitirIlustracao: true });
    const resultado = await resolver(item);
    assert.equal(chamadas[0], "https://api.scryfall.com/cards/one/287");
    assert.equal(Boolean(resultado.ilustrativa), falha !== "nenhuma");
    assert.equal(chamadas.some(url => url.includes("/search?")), falha !== "nenhuma");
    if (falha !== "nenhuma") {
      assert.match(resultado.referenciaOriginal, /Phyrexia: All Will Be One \/ ONE #287/);
      assert.equal(resultado.referencia, "ONE #66");
      const bytes = await gerarExcelColecao({ colecao: { criterios: "variantes" }, itens: [{ ...item, quantidade: 1, idioma: "en", status: "erro" }] },
        { resolverImagem: async () => resultado });
      const arquivo = new globalThis.ExcelJS.Workbook();
      await arquivo.xlsx.load(bytes);
      const folha = arquivo.getWorksheet("Colecao");
      assert.match(folha.getCell("A503").value, /Imagem apenas para ilustracao.*ONE #287.*ONE #66/);
      assert.equal(folha.getCell("A503").alignment.wrapText, true);
      assert.equal(folha.getImages().length, 1);
      assert.equal(folha.getColumn(11).hidden, true);
    }
  }
});

test("galeria do Excel incorpora faces na mesma aba com links bidirecionais e sem alterar a tabela", async () => {
  const buffer = Uint8Array.from(atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jz1cAAAAASUVORK5CYII="), (caractere) => caractere.charCodeAt(0));
  const cartas = [{ nome: "Delver of Secrets", centavos: 100 }, { nome: "Delver of Secrets", centavos: 200 }, { nome: "Indisponivel", centavos: 300 }];
  const progresso = [];
  const bytes = await gerarExcel(cartas, { onProgresso: (...valores) => progresso.push(valores), resolverImagem: async (item) => {
    if (item.nome === "Indisponivel") throw new Error("Falha de rede");
    return { referencia: "ISD #51", faces: [{ nome: "Frente", url: "frente", buffer, extension: "png" }, { nome: "Verso", url: "verso", buffer, extension: "png" }] };
  } });
  const arquivo = new globalThis.ExcelJS.Workbook();
  await arquivo.xlsx.load(bytes);
  const folha = arquivo.getWorksheet("Capturas");
  assert.equal(arquivo.worksheets.length, 1);
  assert.equal(folha.getTable("TabelaCapturas").table.tableRef, "A1:H4");
  assert.equal(folha.getColumn("H").hidden, true);
  assert.equal(folha.getCell("A500").value, "Galeria de cartas");
  assert.equal(folha.getCell("A2").value.hyperlink, "#A501");
  assert.equal(folha.getCell("A502").value.hyperlink, "#" + folha.getCell("H2").value);
  assert.equal(folha.getCell("B2").value, 1);
  assert.equal(folha.getCell("E2").value.result, "08 - Outros");
  for (const linha of [3, 4]) {
    const destino = Number(folha.getCell("A" + linha).value.hyperlink.match(/A(\d+)$/)[1]);
    assert.ok(destino > 501);
    assert.equal(folha.getCell("A" + (destino + 1)).value.hyperlink, "#" + folha.getCell("H" + linha).value);
    if (linha === 4) assert.match(folha.getCell("A" + (destino + 3)).value, /Imagem indisponivel: Falha de rede/);
  }
  assert.equal(arquivo.media.length, 2);
  assert.equal(folha.getImages().length, 4);
  assert.ok(folha.getImages().every((imagem) => imagem.range.tl.nativeRow >= 500));
  assert.deepEqual(progresso, [[1, 3], [2, 3], [3, 3]]);
  const referencias = lerExcel(bytes, { type: "array" }).Workbook.Names;
  assert.equal(referencias.length, cartas.length);
  const identidadePrimeira = folha.getCell("H2").value;
  const identidadeSegunda = folha.getCell("H3").value;
  assert.notEqual(identidadePrimeira, identidadeSegunda);
  for (const identidade of [identidadePrimeira, identidadeSegunda]) {
    const referencia = referencias.find((entrada) => entrada.Name === identidade);
    assert.equal(referencia.Sheet, 0);
    assert.equal(referencia.Ref, 'INDEX(\'Capturas\'!$A$2:$A$4,MATCH("' + identidade + '\",\'Capturas\'!$H$2:$H$4,0))');
  }
  const ordenadas = [2, 3, 4].map((linha) => folha.getRow(linha).values).sort((primeira, segunda) => segunda[2] - primeira[2]);
  for (const [indice, valores] of ordenadas.entries()) folha.getRow(indice + 2).values = valores;
  assert.equal(folha.getCell("H4").value, identidadePrimeira);
  assert.equal(folha.getCell("H3").value, identidadeSegunda);
  assert.equal(folha.getCell("A502").value.hyperlink, "#" + folha.getCell("H4").value);
  assert.equal(folha.getCell("B4").value, 1);
  folha.name = "Deck de drag\u00f5es";
  const destino = new globalThis.ExcelJS.Workbook();
  destino.addWorksheet("Outro deck");
  for (const imagem of arquivo.media) destino.addImage(imagem);
  const copiada = destino.addWorksheet("Deck transferido");
  copiada.model = { ...folha.model, name: copiada.name, id: copiada.id };
  const reaberto = new globalThis.ExcelJS.Workbook();
  await reaberto.xlsx.load(await destino.xlsx.writeBuffer());
  const transferida = reaberto.getWorksheet("Deck transferido");
  assert.equal(transferida.getCell("A4").value.hyperlink, "#A501");
  assert.equal(transferida.getCell("A502").value.hyperlink, "#" + transferida.getCell("H4").value);
  assert.equal(transferida.getImages().length, 4);
  assert.equal(reaberto.media.length, 2);
  for (const [endereco, celula] of Object.entries(lerExcel(await reaberto.xlsx.writeBuffer(), { type: "array" }).Sheets["Deck transferido"])) {
    if (!celula.l) continue;
    if (celula.l.Target.startsWith("#LM_")) continue;
    assert.match(celula.l.Target, /^#A\d+$/, endereco);
    assert.equal(transferida.getCell(celula.l.Target.slice(1)).value !== null, true, endereco);
  }
});

test("galeria do Excel com mais de 500 cartas comeca depois da tabela e preserva a variante cotada", async () => {
  const cartas = Array.from({ length: 500 }, (_, indice) => ({ nome: "Carta " + indice, centavos: 100 }));
  const bytes = await gerarExcel(cartas, { resolverImagem: async () => { throw new Error("Sem imagem"); } });
  const arquivo = new globalThis.ExcelJS.Workbook();
  await arquivo.xlsx.load(bytes);
  const folha = arquivo.getWorksheet("Capturas");
  assert.equal(folha.getTable("TabelaCapturas").table.tableRef, "A1:H501");
  assert.equal(folha.getCell("A100").value.text, "Carta 98");
  assert.equal(folha.getCell("A502").value, "Total Valor:");
  assert.equal(folha.getCell("B502").value.result, 500);
  assert.equal(folha.getCell("A521").value, "Galeria de cartas");
  assert.equal(folha.getCell("A2").value.hyperlink, "#A522");
  let referencia;
  await gerarExcelColecao({ colecao: { criterios: "variantes" }, itens: [{ nome: "Swamp", quantidade: 1, codigoEdicao: "eld", numero: "261", acabamento: "normal", idioma: "en", status: "ok",
    cotacao: { centavos: 20, codigoEdicao: "new", numero: "400" } }] }, { resolverImagem: async (item) => {
    referencia = item;
    throw new Error("Sem imagem");
  } });
  assert.equal(referencia.codigoEdicao, "eld");
  assert.equal(referencia.numero, "261");
});

test("Excel gera XLSX real nos tres relatorios com precos numericos e numero de carta textual", async () => {
  const cotacao = { nome: "Jace's Machinations", centavos: 573, quantidade: 2, tiposCarta: ["Instant"], categoriasTagger: ["Card Draw"],
    codigoEdicao: "fra", numero: "00337", edicao: "Reality Fracture", variante: "Normal", tipoPreco: "media_ligamagic", capturadoEm: "2026-11-02T10:00:00Z" };
  const abrir = (bytes, nome) => {
    assert.ok(bytes instanceof Uint8Array);
    assert.deepEqual([...bytes.slice(0, 4)], [80, 75, 3, 4]);
    const arquivo = lerExcel(bytes, { type: "array", cellNF: true, cellDates: true });
    assert.deepEqual(arquivo.SheetNames, [nome]);
    return arquivo.Sheets[nome];
  };
  const manual = abrir(await gerarExcel([cotacao, { ...cotacao, nome: "=1+1" }], { incluirImagens: false }), "Capturas");
  assert.equal(manual.B1.v, "Pre\u00e7o m\u00ednimo");
  assert.equal(manual.B2.t, "n");
  assert.equal(manual.B2.v, 5.73);
  assert.equal(manual.B2.z, '"R$" #,##0.00');
  assert.equal(manual.D2.v, "Card Draw");
  assert.equal(manual.A3.t, "s");
  assert.equal(manual.A3.f, undefined);
  assert.equal(manual.E1.v, "Filtro");
  assert.equal(manual.E2.v, "02 - Instants");
  assert.equal(manual.F1.v, "Data do preco (UTC)");
  assert.equal(manual.F2.v.toISOString(), cotacao.capturadoEm.replace("Z", ".000Z"));
  assert.equal(manual.G2.v, 2);
  assert.equal(manual.A100.v, "Total Valor:");
  assert.equal(manual.B100.v, 22.92);
  assert.equal(manual.B100.f, "SUMPRODUCT(B2:B3,G2:G3)");
  const cache = abrir(await gerarExcel([{ ...cotacao, capturadoEm: undefined, cachePrecoEm: Date.UTC(2025, 0, 1) }], { incluirImagens: false }), "Capturas");
  assert.equal(cache.F2.v.toISOString(), "2025-01-01T00:00:00.000Z");
  const lote = { colecao: { criterios: "variantes" }, itens: [
    { ...cotacao, acabamento: "normal", idioma: "en", status: "ok", cotacao },
    { nome: "Swamp", quantidade: 3, codigoEdicao: "eld", numero: "261", acabamento: "normal", idioma: "en", status: "pendente" },
  ] };
  const comum = abrir(await gerarExcelLote(lote, { incluirImagens: false }), "Lista");
  assert.equal(comum.B2.v, 5.73);
  assert.equal(comum.A102.v, "Swamp");
  assert.equal(comum.B102.v, "");
  assert.equal(comum.F102.v, "");
  assert.equal(comum.B100.v, 11.46);
  const colecao = abrir(await gerarExcelColecao(lote, { incluirImagens: false }), "Colecao");
  assert.deepEqual(Array.from({ length: 10 }, (_, indice) => colecao[String.fromCharCode(65 + indice) + "1"].v),
    ["Nome da carta", "Preco medio unitario (BRL)", "Subtotal estimado (BRL)", "Quantidade", "Acabamento", "Edicao", "Codigo edicao", "Numero", "Data do preco (UTC)", "URL da referencia"]);
  assert.equal(colecao.B2.t, "n");
  assert.equal(colecao.B2.v, 5.73);
  assert.equal(colecao.C2.v, 11.46);
  assert.equal(colecao.D2.v, 2);
  assert.equal(colecao.E2.v, "normal");
  assert.equal(colecao.H2.t, "s");
  assert.equal(colecao.H2.v, "00337");
  assert.equal(colecao.B3.v, "");
  assert.equal(colecao.I3.v, "");
  assert.equal(colecao.I2.v.toISOString(), "2026-11-02T10:00:00.000Z");
  assert.equal(colecao.B100.v, 11.46);
  assert.equal(colecao.B100.f, "SUM(C2:C3)");
  const antiga = abrir(await gerarExcelColecao({ ...lote, colecao: {} }, { incluirImagens: false }), "Colecao");
  assert.equal(antiga.B2.v, 5.73);
  assert.equal(antiga.C2.v, 11.46);
});

test("Filtro do Excel usa formula por linha com a prioridade dos tipos solicitada", async () => {
  const tipos = ["Artifact | Creature", "Instant", "Sorcery", "Artifact | Enchantment", "Enchantment", "Planeswalker", "Land", "Battle", "", "creature"];
  const capturas = tipos.map((tipo, indice) => ({ nome: "Carta " + indice, centavos: 100, tiposCarta: tipo ? [tipo] : [] }));
  const bytes = await gerarExcel(capturas, { incluirImagens: false });
  const arquivo = lerExcel(bytes, { type: "array" });
  const planilha = arquivo.Sheets.Capturas;
  const categorias = ["01 - Criaturas", "02 - Instants", "03 - Sorceries", "04 - Artefatos", "05 - Encantamentos", "06 - Planeswalkers", "07 - Terrenos", "08 - Outros", "08 - Outros", "01 - Criaturas"];
  for (const [indice, categoria] of categorias.entries()) {
    const linha = indice + 2;
    const celula = planilha["E" + linha];
    assert.equal(celula.v, categoria);
    assert.equal(celula.t, "s");
    assert.equal(celula.f, 'IF(ISNUMBER(SEARCH("Creature",$C' + linha + ')),"01 - Criaturas",IF(ISNUMBER(SEARCH("Instant",$C' + linha + ')),"02 - Instants",IF(ISNUMBER(SEARCH("Sorcery",$C' + linha + ')),"03 - Sorceries",IF(ISNUMBER(SEARCH("Artifact",$C' + linha + ')),"04 - Artefatos",IF(ISNUMBER(SEARCH("Enchantment",$C' + linha + ')),"05 - Encantamentos",IF(ISNUMBER(SEARCH("Planeswalker",$C' + linha + ')),"06 - Planeswalkers",IF(ISNUMBER(SEARCH("Land",$C' + linha + ')),"07 - Terrenos","08 - Outros")))))))');
  }
  const lote = { itens: [{ nome: "Criatura", status: "ok", cotacao: capturas[0] }, { nome: "Pendente", status: "pendente" }] };
  const bytesLista = await gerarExcelLote(lote, { incluirImagens: false });
  const lista = lerExcel(bytesLista, { type: "array" }).Sheets.Lista;
  assert.equal(lista.E1.v, "Filtro");
  assert.equal(lista.E2.v, "01 - Criaturas");
  assert.equal(lista.E3.v, "08 - Outros");
  assert.ok(lista.E3.f.includes('SEARCH("Creature",$C3)'));
  for (const [conteudo, nome, alcance] of [[bytes, "Capturas", "A1:G11"], [bytesLista, "Lista", "A1:G3"]]) {
    const tabelaArquivo = new globalThis.ExcelJS.Workbook();
    await tabelaArquivo.xlsx.load(conteudo);
    const folha = tabelaArquivo.getWorksheet(nome);
    const tabela = folha.getTable("Tabela" + nome).table;
    assert.equal(tabela.tableRef, alcance);
    assert.equal(tabela.autoFilterRef, alcance);
    assert.equal(tabela.style.theme, "TableStyleMedium2");
    assert.equal(tabela.style.showRowStripes, true);
    assert.equal(tabela.columns[4].name, "Filtro");
    assert.ok(tabela.columns.every((coluna) => coluna.filterButton));
    assert.equal(folha.views[0].ySplit, 1);
  }
  const vazia = lerExcel(await gerarExcel([]), { type: "array" }).Sheets.Capturas;
  assert.equal(vazia.E1.v, "Filtro");
  assert.equal(vazia["!ref"], "A1:G100");
  assert.equal(vazia.B100.v, 0);
});

function tag(slug, namespace = "card", status = "GOOD_STANDING", ancestorTags = []) {
  return { name: slug.replaceAll("-", " "), slug, namespace, status, ancestorTags };
}

function simular({ tags = [tag("mana-rock", "card", "GOOD_STANDING", [tag("mana-producer"), tag("ramp")])],
  carta = {}, resultado = {}, codigo = 200 } = {}) {
  const chamadas = [];
  const requisitar = async (url, opcoes) => {
    chamadas.push({ url, opcoes });
    if (url.startsWith("https://api.scryfall.com/")) return { ok: true, json: async () => ({ name: "Sol Ring", oracle_id: "oracle-sol",
      set: "cmm", collector_number: "410", type_line: "Artifact", layout: "normal", ...carta }) };
    throw new Error("Consulta HTTP inesperada.");
  };
  const consultarTags = async (url, variables) => {
    chamadas.push({ url, variables });
    if (codigo !== 200) throw new Error("Scryfall Tagger retornou HTTP " + codigo + ".");
    return { data: {
      front: { name: "Sol Ring", oracleId: "oracle-sol", taggings: tags.map((valor) => ({ status: "GOOD_STANDING", tag: valor })) },
    }, ...resultado };
  };
  return { chamadas, requisitar, consultarTags };
}

test("nomes de duas faces aceitam uma face exata sem misturar pares diferentes", () => {
  const frente = "Grist, Voracious Larva";
  const verso = "Grist, the Plague Swarm";
  const completo = frente + " // " + verso;
  for (const nome of [frente, verso, completo, frente + "//" + verso, "  GRIST, VORACIOUS LARVA  "]) {
    assert.ok(nomesEquivalentes(completo, nome), nome);
    assert.ok(nomesEquivalentes(nome, completo), nome);
    assert.doesNotThrow(() => conferirNome({ nome: frente }, nome === verso ? completo : nome));
  }
  for (const nome of ["Grist", "Grist, the Hunger Tide", frente + " // Sol Ring", "Sol Ring // " + verso]) {
    assert.ok(!nomesEquivalentes(completo, nome), nome);
    assert.throws(() => conferirNome({ nome: completo }, nome), /nao corresponde/);
  }
});

test("filtra somente as categorias permitidas e unifica protecao", () => {
  assert.deepEqual(filtrarCategorias(["gives shroud", "gives-protection", "gives indestructible", "gives-hexproof", "gives-flying", "unknown"]),
    ["Gives Flying", "Protection"]);
  assert.deepEqual(filtrarCategorias(["hexproof", "indestructible", "artwork"]), []);
  assert.deepEqual(filtrarCategorias(["counter-+1/+1", "counter--1/-1"]), ["+1/+1 Counters", "-1/-1 Counters"]);
  for (const nome of CATEGORIAS_PERMITIDAS) {
    assert.deepEqual(filtrarCategorias([nome], [], { textoOracle: "Sacrifice a creature: " }), [nome]);
  }
  assert.equal(new Set(CATEGORIAS_PERMITIDAS).size, CATEGORIAS_PERMITIDAS.length);
});

test("count matters reconhece automaticamente variantes e preserva funcoes independentes", () => {
  const variantes = ["artifact-count-matters", "creature-count-matters", "land-count-matters", "enchantment-count-matters",
    "card-type-count-matters", "artifact-count-matters-opponent", "future-type-count-matters"];
  for (const variante of variantes) {
    const esperado = variante.split("-").map((palavra) => palavra.charAt(0).toUpperCase() + palavra.slice(1)).join(" ");
    assert.deepEqual(filtrarCategorias([variante]), [esperado]);
  }
  assert.deepEqual(filtrarCategorias(["artifact-count-matters", "artifact count matters", "artifact_count_matters"]), ["Artifact Count Matters"]);
  assert.deepEqual(filtrarCategorias(["artifact-count-matters", "creature-count-matters", "draw"], [], {
    usosTags: { "artifact-count-matters": 10, "creature-count-matters": 100, draw: 50 },
  }), ["Creature Count Matters", "Draw Cards", "Artifact Count Matters"]);
  assert.deepEqual(filtrarCategorias(["count-matters", "artifact-count-matters"], [
    { tag: "artifact-count-matters", ancestrais: ["count-matters"] },
  ]), ["Artifact Count Matters"]);
  assert.deepEqual(filtrarCategorias(["artifact-count", "artifact-matters", "count-matters-", "discount-matters"]), []);
});

test("count matters respeita namespace, status, ancestrais e os dois CSVs", async () => {
  const casos = [
    { tags: [tag("artifact-count-matters")], esperado: ["Artifact Count Matters"] },
    { tags: [tag("unknown", "card", "GOOD_STANDING", [tag("artifact-count-matters")])], esperado: ["Artifact Count Matters"] },
    { tags: [tag("artifact-count-matters", "artwork"), tag("creature-count-matters", "card", "REJECTED")], esperado: [] },
    { tags: [{ ...tag("artifact-count-matters"), taggingCount: 100 }, { ...tag("draw"), taggingCount: 50 }],
      esperado: ["Artifact Count Matters", "Draw Cards"] },
  ];
  for (const caso of casos) {
    const { requisitar, consultarTags } = simular({ tags: caso.tags });
    const dados = await consultarTagger("Sol Ring", requisitar, consultarTags);
    assert.deepEqual(dados.categoriasTagger, caso.esperado);
    const cotacao = { nome: "Sol Ring", centavos: 123, ...dados };
    for (const csv of [gerarCsv([cotacao]), gerarCsvLote({ itens: [{ nome: cotacao.nome, status: "ok", cotacao }] })]) {
      assert.equal(csv.trimEnd().split("\r\n")[0].split(";").length, 4);
      assert.ok(csv.endsWith(";" + (caso.esperado.join(" | ") || "Land") + "\r\n"));
    }
  }
});

test("habilidades concedidas distinguem vigilance de haste e aceitam ambas", () => {
  assert.deepEqual(filtrarCategorias(["gives-vigilance"]), ["Gives Vigilance"]);
  assert.deepEqual(filtrarCategorias(["gives-haste"]), ["Gives Haste"]);
  assert.deepEqual(filtrarCategorias(["gives-haste", "gives-vigilance"]), ["Gives Haste", "Gives Vigilance"]);
  assert.deepEqual(filtrarCategorias(["vigilance", "haste"]), []);
});

test("todas as tags gives recebem categorias distintas sem cadastro individual", () => {
  const habilidades = ["first-strike", "double-strike", "reach", "menace", "ward", "infect", "trample", "flying",
    "deathtouch", "lifelink", "haste", "vigilance", "future-ability"];
  for (const habilidade of habilidades) {
    const esperado = "Gives " + habilidade.split("-").map((palavra) => palavra.charAt(0).toUpperCase() + palavra.slice(1)).join(" ");
    assert.deepEqual(filtrarCategorias(["gives-" + habilidade]), [esperado]);
    assert.ok(!filtrarCategorias([habilidade]).includes(esperado));
  }
  assert.deepEqual(filtrarCategorias(["gives first strike", "gives_first_strike", "gives-first-strike"]), ["Gives First Strike"]);
  assert.deepEqual(filtrarCategorias(["gives", "gives-", "forgives-ward", "unknown"]), []);
  assert.deepEqual(filtrarCategorias(["gives-ward", "gives-first-strike", "gives-shroud", "gives-protection", "gives-indestructible", "gives-hexproof"]),
    ["Protection", "Gives Ward", "Gives First Strike"]);
  assert.deepEqual(filtrarCategorias(["gives-menace", "evasion"], [{ tag: "gives-menace", ancestrais: ["evasion"] }]), ["Gives Menace"]);
});

test("novas tags gives respeitam validade e hierarquia e chegam aos dois CSVs", async () => {
  const { requisitar, consultarTags } = simular({ tags: [tag("gives-first-strike"), tag("gives-double-strike"),
    tag("gives-menace", "card", "GOOD_STANDING", [tag("evasion")]), tag("gives-ward"),
    tag("gives-shroud"), tag("gives-hexproof"), tag("gives-reach", "artwork"), tag("gives-infect", "card", "REJECTED")] });
  const dados = await consultarTagger("Sol Ring", requisitar, consultarTags);
  assert.deepEqual(dados.categoriasTagger, ["Protection", "Gives First Strike", "Gives Double Strike", "Gives Menace", "Gives Ward"]);
  const cotacao = { nome: "Sol Ring", centavos: 123, ...dados };
  for (const csv of [gerarCsv([cotacao]), gerarCsvLote({ itens: [{ nome: cotacao.nome, status: "ok", cotacao }] })]) {
    assert.equal(csv.trimEnd().split("\r\n")[0].split(";").length, 4);
    assert.ok(csv.endsWith(";" + dados.categoriasTagger.join(" | ") + "\r\n"));
  }
});

test("Embodiment of Insight preserva landfall e ambas as habilidades concedidas no CSV", async () => {
  const { requisitar, consultarTags } = simular({ carta: { name: "Embodiment of Insight" },
    tags: [tag("landfall"), tag("gives-haste"), tag("gives-vigilance")] });
  const dados = await consultarTagger("Embodiment of Insight", requisitar, consultarTags);
  assert.deepEqual(dados.categoriasTagger, ["Landfall", "Gives Haste", "Gives Vigilance"]);
  assert.ok(gerarCsv([{ nome: "Embodiment of Insight", centavos: 5, ...dados }])
    .endsWith(";Landfall | Gives Haste | Gives Vigilance\r\n"));
});

test("combina tags reais de remocao e nao trata ganho do oponente como beneficio proprio", () => {
  assert.deepEqual(filtrarCategorias(["removal"]), ["Removal"]);
  assert.deepEqual(filtrarCategorias(["removal-permanent"]), ["Permanent Removal"]);
  assert.deepEqual(filtrarCategorias(["removal-planeswalker"]), ["Planeswalker Removal"]);
  assert.deepEqual(filtrarCategorias(["removal-land"]), ["Land Removal"]);
  assert.deepEqual(filtrarCategorias(["removal", "removal-permanent", "removal-land"]), ["Land Removal"]);
  assert.deepEqual(filtrarCategorias(["removal", "removal-creature", "removal-destroy"]), ["Destroy Creature"]);
  assert.deepEqual(filtrarCategorias(["removal-exile", "removal-creature", "lifegain", "opponent-lifegain"]), ["Exile Creature"]);
  assert.deepEqual(filtrarCategorias(["sweeper", "removal-creature", "removal-destroy"]), ["Destroy Creature", "Board Wipe"]);
  assert.deepEqual(filtrarCategorias(["gives-indestructible", "gives-hexproof", "protection"]), ["Protection"]);
  assert.deepEqual(filtrarCategorias(["removal-sacrifice"]), ["Force Sacrifice"]);
  assert.deepEqual(filtrarCategorias(["card-advantage"]), ["Card Advantage"]);
});

test("removal cobre alvos e metodos sem duplicar categorias gerais", () => {
  const alvos = [
    ["creature", "Creature Removal", "Destroy Creature", "Exile Creature"],
    ["artifact", "Artifact Removal", "Destroy Artifact", "Exile Artifact"],
    ["enchantment", "Enchantment Removal", "Destroy Enchantment", "Exile Enchantment"],
    ["planeswalker", "Planeswalker Removal", "Destroy Planeswalker", "Exile Planeswalker"],
    ["land", "Land Removal", "Destroy Land", "Exile Land"],
    ["permanent", "Permanent Removal", "Destroy Permanent", "Exile Permanent"],
  ];
  for (const [alvo, generica, destruir, exilar] of alvos) {
    assert.deepEqual(filtrarCategorias(["removal-" + alvo]), [generica]);
    assert.deepEqual(filtrarCategorias(["removal", "removal-" + alvo, "removal-destroy"]), [destruir]);
    assert.deepEqual(filtrarCategorias(["removal", "removal-" + alvo, "removal-exile"]), [exilar]);
  }
  assert.deepEqual(filtrarCategorias(["removal-destroy"]), ["Destroy Removal"]);
  assert.deepEqual(filtrarCategorias(["removal-exile"]), ["Exile Removal"]);
  assert.deepEqual(filtrarCategorias(["removal-bounce", "bounce", "removal"]), ["Bounce to Hand"]);
  assert.deepEqual(filtrarCategorias(["removal-tuck", "removal"]), ["Tuck Removal"]);
  assert.deepEqual(filtrarCategorias(["removal-aura", "removal-equipment", "removal-battle", "removal"]),
    ["Removal Aura", "Removal Equipment", "Removal Battle"]);
  assert.deepEqual(filtrarCategorias(["removal-future-type", "removal"], [{ tag: "removal-future-type", ancestrais: ["removal"] }]), ["Removal Future Type"]);
  assert.deepEqual(filtrarCategorias(["removal-", "not-removal-creature"]), []);
});

test("variantes de removal respeitam uso, validade e os dois CSVs", async () => {
  const casos = [
    { tags: [tag("removal")], esperado: ["Removal"] },
    { tags: [tag("unknown", "card", "GOOD_STANDING", [tag("removal-creature")])], esperado: ["Creature Removal"] },
    { tags: [tag("removal-land"), tag("removal-destroy"), tag("removal"), tag("draw")], esperado: ["Destroy Land", "Draw Cards"] },
    { tags: [tag("removal-battle", "artwork"), tag("removal", "card", "REJECTED")], esperado: [] },
    { tags: [{ ...tag("removal-battle"), taggingCount: 20 }, { ...tag("draw"), taggingCount: 100 },
      { ...tag("removal"), taggingCount: 10000 }], esperado: ["Draw Cards", "Removal Battle"] },
  ];
  for (const caso of casos) {
    const { requisitar, consultarTags } = simular({ tags: caso.tags });
    const dados = await consultarTagger("Sol Ring", requisitar, consultarTags);
    assert.deepEqual(dados.categoriasTagger, caso.esperado);
    const cotacao = { nome: "Sol Ring", centavos: 123, ...dados };
    for (const csv of [gerarCsv([cotacao]), gerarCsvLote({ itens: [{ nome: cotacao.nome, status: "ok", cotacao }] })]) {
      assert.equal(csv.trimEnd().split("\r\n")[0].split(";").length, 4);
      assert.ok(csv.endsWith(";" + (caso.esperado.join(" | ") || "Land") + "\r\n"));
    }
  }
});

test("Group Slug / Direct Damage unifica aliases e aparece nos dois CSVs", async () => {
  assert.deepEqual(filtrarCategorias(["group-slug", "group slug", "group_slug"]), ["Group Slug / Direct Damage"]);
  assert.deepEqual(filtrarCategorias(["opponent-loses-life", "direct-damage"]), ["Group Slug / Direct Damage"]);
  assert.deepEqual(filtrarCategorias(["opponent-loses-life", "group-slug"]), ["Group Slug / Direct Damage"]);
  assert.deepEqual(filtrarCategorias(["burn"]), ["Group Slug / Direct Damage"]);
  assert.deepEqual(filtrarCategorias(["burn", "opponent-loses-life", "group-slug"]), ["Group Slug / Direct Damage"]);
  assert.deepEqual(filtrarCategorias(["burn", "draw"], [], { usosTags: { burn: 500, draw: 100 } }), ["Group Slug / Direct Damage", "Draw Cards"]);
  const casos = [
    { tags: [tag("group-slug")], esperado: ["Group Slug / Direct Damage"] },
    { tags: [tag("opponent-loses-life")], esperado: ["Group Slug / Direct Damage"] },
    { tags: [tag("burn")], esperado: ["Group Slug / Direct Damage"] },
    { tags: [tag("burn"), tag("group-slug"), tag("opponent-loses-life")], esperado: ["Group Slug / Direct Damage"] },
    { tags: [tag("burn", "artwork"), tag("burn", "card", "REJECTED")], esperado: [] },
    { tags: [tag("group-slug"), tag("opponent-loses-life")], esperado: ["Group Slug / Direct Damage"] },
    { tags: [tag("group-slug"), tag("drain-life")], esperado: ["Life Drain", "Group Slug / Direct Damage"] },
    { tags: [tag("group-slug", "artwork")], esperado: [] },
    { tags: [tag("group-slug", "card", "REJECTED")], esperado: [] },
  ];
  for (const caso of casos) {
    const { requisitar, consultarTags } = simular({ tags: caso.tags });
    const dados = await consultarTagger("Sol Ring", requisitar, consultarTags);
    assert.deepEqual(dados.categoriasTagger, caso.esperado);
    const cotacao = { nome: "Sol Ring", centavos: 123, ...dados };
    for (const csv of [gerarCsv([cotacao]), gerarCsvLote({ itens: [{ nome: cotacao.nome, status: "ok", cotacao }] })]) {
      assert.ok(csv.endsWith(";" + (caso.esperado.join(" | ") || "Land") + "\r\n"));
    }
  }
});

test("categorias especificas prevalecem sobre seus ancestrais", () => {
  assert.deepEqual(filtrarCategorias(["draw", "card-advantage"]), ["Draw Cards"]);
  assert.deepEqual(filtrarCategorias(["card-advantage"]), ["Card Advantage"]);
  assert.deepEqual(filtrarCategorias(["mill-self", "mill"], [{ tag: "mill-self", ancestrais: ["mill"] }]), ["Self-Mill"]);
  assert.deepEqual(filtrarCategorias(["menace", "evasion"], [{ tag: "menace", ancestrais: ["evasion"] }]), ["Evasion"]);
  assert.deepEqual(filtrarCategorias(["unblockable", "evasion"], [{ tag: "unblockable", ancestrais: ["evasion"] }]), ["Unblockable"]);
  assert.deepEqual(filtrarCategorias(["mill"]), ["Mill"]);
});

test("categorias redundantes sao removidas mesmo sem ancestrais e funcoes independentes ficam", () => {
  const casos = [
    { tags: ["draw", "card-advantage", "scry"], esperado: ["Draw Cards", "Scry"] },
    { tags: ["mill", "mill-self"], esperado: ["Self-Mill"] },
    { tags: ["reanimate-self", "reanimate", "recursion"], esperado: ["Self-Reanimate"] },
    { tags: ["regrowth", "recursion", "card-advantage"], esperado: ["Return to Hand"] },
    { tags: ["drain-life", "lifegain", "group-slug"], esperado: ["Life Drain", "Group Slug / Direct Damage"] },
    { tags: ["mana-rock", "mana-producer", "ramp", "draw", "card-advantage"], esperado: ["Draw Cards", "Mana Rock"] },
    { tags: ["unblockable", "evasion", "gives-flying"], esperado: ["Unblockable", "Gives Flying"] },
  ];
  for (const caso of casos) assert.deepEqual(filtrarCategorias(caso.tags), caso.esperado);
});

test("tokens reconhecem tags reais, priorizam tipos especificos e nao perdem outras funcoes", async () => {
  const casos = [
    { tags: [tag("repeatable-token-generator")], esperado: ["Create Tokens"] },
    { tags: [tag("repeatable-creature-tokens")], esperado: ["Create Creature Tokens"] },
    { tags: [tag("repeatable-noncreature-tokens")], esperado: ["Create Noncreature Tokens"] },
    { tags: [tag("repeatable-artifact-tokens")], esperado: ["Create Artifact Tokens"] },
    { tags: [tag("repeatable-token-generator"), tag("repeatable-creature-tokens"), tag("token-zombie")], esperado: ["Zombies"] },
    { tags: [tag("repeatable-treasures"), tag("repeatable-token-generator"), tag("repeatable-noncreature-tokens"), tag("repeatable-artifact-tokens")],
      esperado: ["Treasure Tokens"] },
    { tags: [tag("draw"), tag("card-advantage"), tag("repeatable-creature-tokens"), tag("scry")], esperado: ["Create Creature Tokens", "Draw Cards", "Scry"] },
    { tags: [tag("token-zombie"), tag("token-soldier"), tag("draw"), tag("scry"), tag("surveil")], esperado: ["Zombies", "Soldiers", "Draw Cards", "Scry", "Surveil"] },
    { tags: [tag("repeatable-creature-tokens", "artwork"), tag("repeatable-token-generator", "card", "REJECTED")], esperado: [] },
  ];
  for (const caso of casos) {
    const { requisitar, consultarTags } = simular({ tags: caso.tags });
    const dados = await consultarTagger("Sol Ring", requisitar, consultarTags);
    assert.deepEqual(dados.categoriasTagger, caso.esperado);
    const cotacao = { nome: "Sol Ring", centavos: 123, ...dados };
    for (const csv of [gerarCsv([cotacao]), gerarCsvLote({ itens: [{ nome: cotacao.nome, status: "ok", cotacao }] })]) {
      assert.equal(csv.trimEnd().split("\r\n")[0].split(";").length, 4);
      assert.ok(csv.endsWith(";" + (caso.esperado.join(" | ") || "Land") + "\r\n"));
    }
  }
});

test("tags repetidas mesclam ancestrais mais completos sem depender da ordem", async () => {
  const rasa = tag("unknown", "card", "GOOD_STANDING", [tag("repeatable-creature-tokens")]);
  const completa = tag("repeatable-creature-tokens", "card", "GOOD_STANDING", [tag("repeatable-token-generator")]);
  for (const tags of [[rasa, completa], [completa, rasa]]) {
    const { requisitar, consultarTags } = simular({ tags });
    const dados = await consultarTagger("Sol Ring", requisitar, consultarTags);
    assert.ok(dados.tagsTagger.includes("repeatable-token-generator"));
    assert.deepEqual(dados.categoriasTagger, ["Create Creature Tokens"]);
    assert.deepEqual(dados.categoriasTaggerSecundarias, ["Create Tokens"]);
  }
});

test("uso das tags ordena apenas categorias validas e tokens desempatem", () => {
  const tags = ["draw", "card-advantage", "scry", "repeatable-creature-tokens"];
  assert.deepEqual(filtrarCategorias(tags, [], { usosTags: { draw: 500, "card-advantage": 9000, scry: 100, "repeatable-creature-tokens": 200 } }),
    ["Draw Cards", "Create Creature Tokens", "Scry"]);
  assert.deepEqual(filtrarCategorias(tags, [], { usosTags: { draw: 100, "repeatable-creature-tokens": 100, scry: 50 } }),
    ["Create Creature Tokens", "Draw Cards", "Scry"]);
  assert.deepEqual(filtrarCategorias(["draw"], [], { usosTags: { draw: 1, "group-slug": 9999 } }), ["Draw Cards"]);
  assert.deepEqual(filtrarCategorias(["gives-ward", "scry"], [], { usosTags: { "gives-ward": 10, scry: 1 } }), ["Gives Ward", "Scry"]);
  assert.deepEqual(filtrarCategorias(["gives-shroud", "gives-hexproof", "scry"], [], { usosTags: { "gives-shroud": 10, "gives-hexproof": 10, scry: 15 } }),
    ["Scry", "Protection"]);
  assert.deepEqual(filtrarCategorias(["scry", "removal-creature", "removal-exile"], [], { usosTags: { scry: 100, "removal-creature": 500, "removal-exile": 200 } }),
    ["Exile Creature", "Scry"]);
  assert.deepEqual(filtrarCategorias(["draw", "scry", "repeatable-creature-tokens"], [], { usosTags: { draw: NaN, scry: -1, "repeatable-creature-tokens": "100" } }),
    ["Create Creature Tokens", "Draw Cards", "Scry"]);
});

test("consulta usa taggingCount de tags e ancestrais nas duas faces e preserva ordem no CSV", async () => {
  const compra = { ...tag("draw", "card", "GOOD_STANDING", [{ ...tag("card-advantage"), taggingCount: 9999 }]), taggingCount: 500 };
  const tokens = { ...tag("repeatable-creature-tokens", "card", "GOOD_STANDING", [{ ...tag("repeatable-token-generator"), taggingCount: 99999 }]), taggingCount: 200 };
  const scry = { ...tag("scry"), taggingCount: 100 };
  const ignorada = { ...tag("group-slug", "artwork"), taggingCount: 999999 };
  const rejeitada = { ...tag("gives-ward", "card", "REJECTED"), taggingCount: 999999 };
  const { requisitar, consultarTags } = simular({ tags: [compra, tokens, scry, ignorada, rejeitada] });
  const dados = await consultarTagger("Sol Ring", requisitar, consultarTags);
  assert.deepEqual(dados.categoriasTagger, ["Draw Cards", "Create Creature Tokens", "Scry"]);
  assert.equal(dados.usosTagsTagger.draw, 500);
  assert.equal(dados.usosTagsTagger["repeatable-creature-tokens"], 200);
  assert.ok(!Object.hasOwn(dados.usosTagsTagger, "group-slug"));
  assert.ok(!Object.hasOwn(dados.usosTagsTagger, "gives-ward"));
  const cotacao = { nome: "Sol Ring", centavos: 123, ...dados };
  for (const csv of [gerarCsv([cotacao]), gerarCsvLote({ itens: [{ nome: cotacao.nome, status: "ok", cotacao }] })]) {
    assert.ok(csv.endsWith(";Draw Cards | Create Creature Tokens | Scry\r\n"));
  }
  const dupla = simular({ carta: { layout: "transform" }, resultado: { data: {
    front: { oracleId: "oracle-sol", taggings: [{ status: "GOOD_STANDING", tag: compra }] },
    back: { oracleId: "oracle-sol", taggings: [{ status: "GOOD_STANDING", tag: { ...tokens, taggingCount: 800 } },
      { status: "GOOD_STANDING", tag: { ...compra, taggingCount: 400 } }] },
  } } });
  const duasFaces = await consultarTagger("Sol Ring", dupla.requisitar, dupla.consultarTags);
  assert.deepEqual(duasFaces.categoriasTagger, ["Create Creature Tokens", "Draw Cards"]);
  assert.equal(duasFaces.usosTagsTagger.draw, 500);
  const falha = await enriquecerCotacao(cotacao, async () => { throw new Error("Offline"); });
  assert.deepEqual(falha.usosTagsTagger, {});
  assert.equal(falha.centavos, 123);
});

test("untapper e trigger doubler reconhecem tags exatas sem duplicar categorias gerais", () => {
  assert.deepEqual(filtrarCategorias(["untapper"]), ["Untap"]);
  assert.deepEqual(filtrarCategorias(["untap", "untapper"]), ["Untap"]);
  assert.deepEqual(filtrarCategorias(["untapper", "untap-land"]), ["Untap Lands"]);
  assert.deepEqual(filtrarCategorias(["trigger-doubler"]), ["Trigger Doubler"]);
  assert.deepEqual(filtrarCategorias(["trigger-doubler", "etb-trigger"]), ["ETB Effects", "Trigger Doubler"]);
  assert.deepEqual(filtrarCategorias(["triggered-ability", "double", "untapped"]), []);
  assert.deepEqual(filtrarCategorias(["untapper", "trigger-doubler"], [], { usosTags: { untapper: 800, "trigger-doubler": 20 } }),
    ["Untap", "Trigger Doubler"]);
});

test("Cloud of Faeries com untapper e cartas com trigger doubler exportam as novas categorias", async () => {
  const casos = [
    { nome: "Cloud of Faeries", tags: [tag("untapper")], esperado: ["Untap"] },
    { nome: "Sol Ring", tags: [tag("trigger-doubler")], esperado: ["Trigger Doubler"] },
    { nome: "Sol Ring", tags: [tag("unknown", "card", "GOOD_STANDING", [tag("trigger-doubler")])], esperado: ["Trigger Doubler"] },
    { nome: "Sol Ring", tags: [tag("untapper", "artwork"), tag("trigger-doubler", "card", "REJECTED")], esperado: [] },
  ];
  for (const caso of casos) {
    const { requisitar, consultarTags } = simular({ carta: { name: caso.nome }, tags: caso.tags });
    const dados = await consultarTagger(caso.nome, requisitar, consultarTags);
    assert.deepEqual(dados.categoriasTagger, caso.esperado);
    const cotacao = { nome: caso.nome, centavos: 123, ...dados };
    for (const csv of [gerarCsv([cotacao]), gerarCsvLote({ itens: [{ nome: caso.nome, status: "ok", cotacao }] })]) {
      assert.equal(csv.trimEnd().split("\r\n")[0].split(";").length, 4);
      assert.ok(csv.endsWith(";" + (caso.esperado.join(" | ") || "Land") + "\r\n"));
    }
  }
});

test("phase manipulation e alternate win condition respeitam nomes exatos, ancestrais e uso", () => {
  assert.deepEqual(filtrarCategorias(["phase-manipulation"]), ["Phase Manipulation"]);
  assert.deepEqual(filtrarCategorias(["alternate-win-condition", "alternate win condition"]), ["Alternate Win Condition"]);
  assert.deepEqual(filtrarCategorias(["phase-manipulation", "extra-combat"], [
    { tag: "extra-combat", ancestrais: ["phase-manipulation"] },
  ]), ["Extra Combat"]);
  assert.deepEqual(filtrarCategorias(["phase-manipulation", "alternate-win-condition"], [], {
    usosTags: { "phase-manipulation": 10, "alternate-win-condition": 100 },
  }), ["Alternate Win Condition", "Phase Manipulation"]);
  assert.deepEqual(filtrarCategorias(["phasing", "win", "phase"]), []);
});

test("phase manipulation e alternate win condition entram nos dois CSVs apenas com tags validas", async () => {
  const casos = [
    { tags: [tag("phase-manipulation")], esperado: ["Phase Manipulation"] },
    { tags: [tag("alternate-win-condition")], esperado: ["Alternate Win Condition"] },
    { tags: [tag("unknown", "card", "GOOD_STANDING", [tag("alternate-win-condition")])], esperado: ["Alternate Win Condition"] },
    { tags: [tag("phase-manipulation", "artwork"), tag("alternate-win-condition", "card", "REJECTED")], esperado: [] },
  ];
  for (const caso of casos) {
    const { requisitar, consultarTags } = simular({ tags: caso.tags });
    const dados = await consultarTagger("Sol Ring", requisitar, consultarTags);
    assert.deepEqual(dados.categoriasTagger, caso.esperado);
    const cotacao = { nome: "Sol Ring", centavos: 123, ...dados };
    for (const csv of [gerarCsv([cotacao]), gerarCsvLote({ itens: [{ nome: cotacao.nome, status: "ok", cotacao }] })]) {
      assert.equal(csv.trimEnd().split("\r\n")[0].split(";").length, 4);
      assert.ok(csv.endsWith(";" + (caso.esperado.join(" | ") || "Land") + "\r\n"));
    }
  }
});

test("Hand Size Increase e independente de compra, respeita uso e exporta nos dois CSVs", async () => {
  assert.deepEqual(filtrarCategorias(["hand-size-increase", "hand size increase"]), ["Hand Size Increase"]);
  assert.deepEqual(filtrarCategorias(["draw", "card-advantage", "hand-size-increase"]), ["Draw Cards", "Hand Size Increase"]);
  assert.deepEqual(filtrarCategorias(["draw", "hand-size-increase"], [], { usosTags: { draw: 10, "hand-size-increase": 100 } }),
    ["Hand Size Increase", "Draw Cards"]);
  assert.deepEqual(filtrarCategorias(["hand-size", "hand-size-decrease"]), []);
  const casos = [
    { tags: [tag("hand-size-increase")], esperado: ["Hand Size Increase"] },
    { tags: [tag("unknown", "card", "GOOD_STANDING", [tag("hand-size-increase")])], esperado: ["Hand Size Increase"] },
    { tags: [tag("hand-size-increase", "artwork"), tag("hand-size-increase", "card", "REJECTED")], esperado: [] },
  ];
  for (const caso of casos) {
    const { requisitar, consultarTags } = simular({ tags: caso.tags });
    const dados = await consultarTagger("Sol Ring", requisitar, consultarTags);
    assert.deepEqual(dados.categoriasTagger, caso.esperado);
    const cotacao = { nome: "Sol Ring", centavos: 123, ...dados };
    for (const csv of [gerarCsv([cotacao]), gerarCsvLote({ itens: [{ nome: cotacao.nome, status: "ok", cotacao }] })]) {
      assert.equal(csv.trimEnd().split("\r\n")[0].split(";").length, 4);
      assert.ok(csv.endsWith(";" + (caso.esperado.join(" | ") || "Land") + "\r\n"));
    }
  }
});

test("ramp cobre subcategorias sem repetir ancestrais e sem inventar categorias", async () => {
  const ramp = tag("ramp");
  const produtor = tag("mana-producer", "card", "GOOD_STANDING", [ramp]);
  const casos = [
    { tags: [ramp], esperado: ["Ramp"] },
    { tags: [tag("unknown", "card", "GOOD_STANDING", [ramp])], esperado: ["Ramp"] },
    { tags: [produtor], esperado: ["Adds Mana"] },
    { tags: [tag("ritual", "card", "GOOD_STANDING", [produtor, ramp])], esperado: ["Ritual"] },
    { tags: [tag("mana-doubler", "card", "GOOD_STANDING", [ramp])], esperado: ["Mana Multiplier"] },
    { tags: [tag("mana-tripler", "card", "GOOD_STANDING", [ramp])], esperado: ["Mana Multiplier"] },
    { tags: [tag("mana-multiplier", "card", "GOOD_STANDING", [ramp])], esperado: ["Mana Multiplier"] },
    { tags: [tag("untap-land", "card", "GOOD_STANDING", [tag("untap"), ramp])], esperado: ["Untap Lands"] },
    { tags: [tag("mana-rock", "card", "GOOD_STANDING", [produtor, ramp])], esperado: ["Mana Rock"] },
    { tags: [tag("mana-dork", "card", "GOOD_STANDING", [produtor, ramp])], esperado: ["Mana Dork"] },
    { tags: [tag("tutor-land-to-battlefield", "card", "GOOD_STANDING", [tag("tutor-land"), ramp])], esperado: ["Put Land onto Battlefield"] },
    { tags: [tag("extra-land", "card", "GOOD_STANDING", [ramp])], esperado: ["Extra Land Drops"] },
    { tags: [tag("cost-reduction", "card", "GOOD_STANDING", [ramp])], esperado: ["Cost Reduction"] },
    { tags: [tag("repeatable-treasures", "card", "GOOD_STANDING", [ramp])], esperado: ["Treasure Tokens"] },
    { tags: [tag("mana-rock", "card", "GOOD_STANDING", [produtor, ramp]), tag("draw")], esperado: ["Draw Cards", "Mana Rock"] },
  ];
  for (const caso of casos) {
    const { requisitar, consultarTags } = simular({ tags: caso.tags });
    const dados = await consultarTagger("Sol Ring", requisitar, consultarTags);
    assert.deepEqual(dados.categoriasTagger, caso.esperado);
    assert.ok(gerarCsv([{ nome: "Sol Ring", centavos: 100, ...dados }]).endsWith(";" + caso.esperado.join(" | ") + "\r\n"));
  }
});

test("reanimacao propria, tesouros repetiveis e recursion reconhecem nomes canonicos", () => {
  assert.deepEqual(filtrarCategorias(["self-reanimate"]), ["Self-Reanimate"]);
  assert.deepEqual(filtrarCategorias(["reanimate-self"]), ["Self-Reanimate"]);
  assert.deepEqual(filtrarCategorias(["repeatable-treasures"]), ["Treasure Tokens"]);
  assert.deepEqual(filtrarCategorias(["treasure-token", "repeatable-treasures"]), ["Treasure Tokens"]);
  assert.deepEqual(filtrarCategorias(["recursion"]), ["Recursion"]);
  for (const tag of ["recursion-self", "recursion-any", "recursion-land", "restock", "restock-self"]) {
    assert.deepEqual(filtrarCategorias([tag]), ["Recursion"]);
  }
  assert.deepEqual(filtrarCategorias(["reanimate-self", "reanimate", "recursion"], [
    { tag: "reanimate-self", ancestrais: ["reanimate", "recursion"] },
  ]), ["Self-Reanimate"]);
});

test("Life Drain aceita drenagem e habilidades relacionadas sem confundir efeitos isolados", () => {
  for (const tag of ["drain", "drain-life", "life-drain", "blood-artist-ability"]) {
    assert.deepEqual(filtrarCategorias([tag]), ["Life Drain"]);
  }
  assert.deepEqual(filtrarCategorias(["blood-artist-ability", "drain-life", "lifegain"], [
    { tag: "blood-artist-ability", ancestrais: ["drain-life", "lifegain"] },
  ]), ["Life Drain"]);
  assert.ok(!filtrarCategorias(["opponent-loses-life"]).includes("Life Drain"));
  assert.ok(!filtrarCategorias(["lifegain"]).includes("Life Drain"));
});

test("familias preservam apenas categorias especificas e mantem CSV com quatro colunas", async () => {
  const casos = [
    { tags: [tag("mill-self", "card", "GOOD_STANDING", [tag("mill")])], esperado: ["Self-Mill"] },
    { tags: [tag("reanimate-self", "card", "GOOD_STANDING", [tag("reanimate"), tag("recursion"), tag("recursion-self")])],
      esperado: ["Self-Reanimate"] },
    { tags: [tag("repeatable-treasures")], esperado: ["Treasure Tokens"] },
    { tags: [tag("blood-artist-ability", "card", "GOOD_STANDING", [tag("drain-life"), tag("lifegain"), tag("death-trigger")])],
      esperado: ["Life Drain"] },
    { tags: [tag("regrowth-any", "card", "GOOD_STANDING", [tag("regrowth"), tag("recursion-any"), tag("recursion"), tag("card-advantage")])],
      esperado: ["Return to Hand"] },
  ];
  for (const caso of casos) {
    const { requisitar, consultarTags } = simular({ tags: caso.tags });
    const dados = await consultarTagger("Sol Ring", requisitar, consultarTags);
    assert.deepEqual(dados.categoriasTagger, caso.esperado);
    const cotacao = { nome: "Sol Ring", centavos: 100, ...dados };
    for (const csv of [gerarCsv([cotacao]), gerarCsvLote({ itens: [{ nome: cotacao.nome, status: "ok", cotacao }] })]) {
      assert.equal(csv.trimEnd().split("\r\n")[0].split(";").length, 4);
      assert.ok(csv.endsWith(";" + caso.esperado.join(" | ") + "\r\n"));
    }
  }
});

test("Sacrifice Outlet exige habilidade ativada barata e repetivel", () => {
  for (const texto of ["Sacrifice a creature: Scry 1.", "Pay 1 life, Sacrifice another creature: Draw a card.",
    "{B}, Sacrifice a creature: Destroy target creature.", "{1}, Sacrifice an artifact: Draw a card."]) {
    assert.ok(permiteSacrificioBarato(texto), texto);
    assert.deepEqual(filtrarCategorias(["sacrifice-outlet"], [], { textoOracle: texto }), ["Sacrifice Outlet"]);
  }
  for (const texto of ["As an additional cost to cast this spell, sacrifice a creature.\nDraw two cards.",
    "When this creature enters, each player sacrifices a creature.", "{1}{B}, Sacrifice a creature: Draw a card.",
    "{T}, Sacrifice a creature: Draw a card.", "Pay 2 life, Sacrifice a creature: Draw a card.",
    "Sacrifice this creature: Draw a card.", "Remove a counter, Sacrifice a creature: Draw a card.",
    "Sacrifice a creature: Draw a card. Activate only once each turn.", "Sacrifice two creatures: Draw a card.",
    "Discard a card, Sacrifice a creature: Draw a card."]) {
    assert.ok(!permiteSacrificioBarato(texto), texto);
    assert.deepEqual(filtrarCategorias(["sacrifice-outlet"], [], { textoOracle: texto }), []);
  }
  assert.deepEqual(filtrarCategorias(["sacrifice-outlet"]), []);
});

test("Force Sacrifice nao e suprimida pela prioridade de ancestrais nem vira outlet", () => {
  const tags = ["sacrifice-outlet", "removal-sacrifice"];
  const relacoes = [{ tag: "sacrifice-outlet", ancestrais: ["removal-sacrifice"] }];
  assert.deepEqual(filtrarCategorias(tags, relacoes, { textoOracle: "When this creature enters, each player sacrifices a creature." }), ["Force Sacrifice"]);
  assert.deepEqual(filtrarCategorias(tags, relacoes, { textoOracle: "Sacrifice a creature: Each opponent sacrifices a creature." }),
    ["Force Sacrifice", "Sacrifice Outlet"]);
});

test("consulta nome exato e tags funcionais validas incluindo ancestrais", async () => {
  const { chamadas, requisitar, consultarTags } = simular({ tags: [tag("mana-rock", "card", "GOOD_STANDING", [tag("mana-producer")]),
    tag("gives-hexproof", "artwork"), tag("gives-haste", "card", "REJECTED"), tag("unknown")] });
  const dados = await consultarTagger("Sol Ring", requisitar, consultarTags);
  assert.deepEqual(dados.categoriasTagger, ["Mana Rock"]);
  assert.deepEqual(dados.categoriasTaggerSecundarias, ["Adds Mana"]);
  assert.ok(dados.tagsTaggerDiretas.includes("mana-rock"));
  assert.ok(!dados.tagsTaggerDiretas.includes("mana-producer"));
  assert.ok(dados.tagsTaggerSecundarias.includes("mana-producer"));
  assert.deepEqual(dados.tiposCarta, ["Artifact"]);
  assert.equal(chamadas.length, 2);
  assert.equal(new URL(chamadas[0].url).searchParams.get("exact"), "Sol Ring");
  assert.equal(chamadas[1].variables.twoSided, false);
  assert.equal(dados.statusTagger, "ok");
});

test("Stitcher's Supplier descarta ancestrais da selecao mesmo com menos de tres categorias", async () => {
  const { requisitar, consultarTags } = simular({ tags: [tag("mill-self", "card", "GOOD_STANDING", [tag("mill")]),
    tag("life-and-death-trigger-self", "card", "GOOD_STANDING", [
      tag("death-trigger-self", "card", "GOOD_STANDING", [tag("death-trigger")]), tag("death-trigger"), tag("triggered-ability"),
    ])] });
  const dados = await consultarTagger("Sol Ring", requisitar, consultarTags);
  assert.deepEqual(dados.categoriasTagger, ["Self-Mill", "Self Death Triggers"]);
  assert.deepEqual(dados.categoriasTaggerSecundarias, ["Mill", "Death Triggers"]);
  assert.ok(dados.tagsTaggerSecundarias.includes("death-trigger-self"));
  const csv = gerarCsv([{ nome: "Stitcher's Supplier", centavos: 100, ...dados }]);
  assert.ok(csv.endsWith(";Self-Mill | Self Death Triggers\r\n"));
  assert.equal(csv.trimEnd().split("\r\n")[0].split(";").length, 4);
});

test("categorias nao sao preenchidas com ancestrais nem duplicadas ou truncadas", async () => {
  const casos = [
    { tags: [tag("draw", "card", "GOOD_STANDING", [tag("card-advantage")])], esperado: ["Draw Cards"] },
    { tags: [tag("evasion")], esperado: ["Evasion"] },
    { tags: [], esperado: [] },
    { tags: [tag("scry"), tag("surveil"), tag("draw"), tag("card-advantage")], esperado: ["Draw Cards", "Scry", "Surveil"] },
    { tags: [tag("unknown", "card", "GOOD_STANDING", [tag("sacrifice-outlet"), tag("card-advantage")])], esperado: ["Card Advantage"] },
  ];
  for (const caso of casos) {
    const { requisitar, consultarTags } = simular({ tags: caso.tags });
    const dados = await consultarTagger("Sol Ring", requisitar, consultarTags);
    assert.deepEqual(dados.categoriasTagger, caso.esperado);
    assert.equal(new Set(dados.categoriasTagger).size, dados.categoriasTagger.length);
  }
});

test("falhas HTTP e respostas incorretas preservam cotacoes", async () => {
  for (const opcoes of [{ codigo: 429 }, { resultado: { errors: [{ message: "Falhou" }] } },
    { resultado: { data: { front: { oracleId: "outra-carta", taggings: [] } } } }]) {
    const { requisitar, consultarTags } = simular(opcoes);
    const dados = await enriquecerCotacao({ nome: "Sol Ring", centavos: 123 }, requisitar, consultarTags);
    assert.equal(dados.centavos, 123);
    assert.equal(dados.statusTagger, "erro");
    assert.deepEqual(dados.categoriasTagger, []);
    assert.ok(dados.erroTagger);
  }
});

test("tipos e categorias incluem ambas as faces sem subtipos nem supertipos", async () => {
  const { requisitar, chamadas, consultarTags } = simular({ carta: { layout: "modal_dfc", type_line: "Legendary Creature \u2014 Human // Land",
    card_faces: [{ type_line: "Legendary Creature \u2014 Human" }, { type_line: "Land" }] }, resultado: { data: {
      front: { oracleId: "oracle-sol", taggings: [{ status: "GOOD_STANDING", tag: tag("gives-shroud") }] },
      back: { oracleId: "oracle-sol", taggings: [{ status: "GOOD_STANDING", tag: tag("mana-producer") }] },
    } } });
  const dados = await consultarTagger("Sol Ring", requisitar, consultarTags);
  assert.deepEqual(dados.tiposCarta, ["Creature", "Land"]);
  assert.deepEqual(dados.categoriasTagger, ["Adds Mana", "Protection"]);
  assert.equal(chamadas[1].variables.twoSided, true);
  const dupla = simular({ carta: { type_line: "Instant // Sorcery" } });
  assert.deepEqual((await consultarTagger("Sol Ring", dupla.requisitar, dupla.consultarTags)).tiposCarta, ["Instant", "Sorcery"]);
});

test("Grist aceita nome completo ou de uma face e combina tipos e tags das duas faces", async () => {
  const frente = "Grist, Voracious Larva";
  const verso = "Grist, the Plague Swarm";
  const completo = frente + " // " + verso;
  const carta = { name: completo, oracle_id: "oracle-grist", set: "mh3", collector_number: "189", layout: "transform",
    type_line: "Legendary Creature \u2014 Insect // Legendary Planeswalker \u2014 Grist",
    card_faces: [{ name: frente, type_line: "Legendary Creature \u2014 Insect" },
      { name: verso, type_line: "Legendary Planeswalker \u2014 Grist" }] };
  const resultado = { data: {
    front: { oracleId: "oracle-grist", taggings: [{ status: "GOOD_STANDING", tag: tag("mill-self", "card", "GOOD_STANDING", [tag("mill")]) }] },
    back: { oracleId: "oracle-grist", taggings: [{ status: "GOOD_STANDING", tag: tag("reanimate") },
      { status: "GOOD_STANDING", tag: tag("mill-self", "card", "GOOD_STANDING", [tag("mill")]) }] },
  } };
  for (const nome of [completo, frente, verso, frente + "//" + verso]) {
    const { requisitar, consultarTags, chamadas } = simular({ carta, resultado });
    const dados = await consultarTagger(nome, requisitar, consultarTags);
    assert.equal(new URL(chamadas[0].url).searchParams.get("exact"), nome === verso ? verso : frente);
    assert.equal(chamadas[1].variables.twoSided, true);
    assert.deepEqual(dados.tiposCarta, ["Creature", "Planeswalker"]);
    assert.deepEqual(dados.categoriasTagger, ["Reanimate", "Self-Mill"]);
    assert.deepEqual(dados.tagsTaggerDiretas, ["mill-self", "reanimate"]);
    assert.ok(gerarCsv([{ nome, centavos: 123, ...dados }]).endsWith(";Creature | Planeswalker;Reanimate | Self-Mill\r\n"));
  }
  for (const back of [undefined, { oracleId: "outra-carta", taggings: [] }, { oracleId: "oracle-grist" }]) {
    const { requisitar, consultarTags } = simular({ carta, resultado: { data: { ...resultado.data, back } } });
    await assert.rejects(consultarTagger(frente, requisitar, consultarTags), /verso/);
  }
  const { requisitar, consultarTags } = simular({ carta, resultado });
  await assert.rejects(consultarTagger(frente + " // Sol Ring", requisitar, consultarTags), /Carta exata nao encontrada/);
});

test("lista com nome completo de duas faces busca o nome completo e o preserva no CSV", async () => {
  const frente = "Grist, Voracious Larva";
  const completo = frente + " // Grist, the Plague Swarm";
  const estado = { capturas: [] };
  let urlAberta;
  const pagina = { nome: frente, url: "https://www.ligamagic.com.br/", capturadoEm: "2026-10-27",
    edicoes: [{ name: "Modern Horizons 3", price: { 0: { p: "1.23" } } }] };
  const api = {
    storage: { local: { get: async (chave) => structuredClone({ [chave]: estado[chave] }),
      set: async (dados) => Object.assign(estado, structuredClone(dados)) } },
    permissions: { contains: async () => true },
    tabs: { create: async () => ({ id: 7 }), get: async () => ({ status: "complete", url: urlAberta }),
      update: async (id, { url }) => { assert.equal(id, 7); urlAberta = url; } },
    scripting: { executeScript: async () => [{ result: pagina }] },
    alarms: { create: async () => {}, clear: async () => {} },
  };
  const controle = criarControlador(api, () => 0, async (cotacao) => ({ ...cotacao, statusTagger: "ok",
    tiposCarta: ["Creature", "Planeswalker"], categoriasTagger: ["Self-Mill", "Reanimate"] }));
  await controle.comando("iniciar", { texto: "2 " + completo, intervalo: 3 });
  assert.equal(new URL(urlAberta).searchParams.get("card"), completo);
  await controle.processar(7);
  assert.equal(estado.lote.fase, "concluido");
  assert.equal(estado.lote.itens[0].status, "ok");
  assert.equal(estado.lote.itens[0].nome, completo);
  assert.equal(estado.capturas[0].quantidade, 2);
  assert.ok(gerarCsvLote(estado.lote).includes(completo + ";1,23;Creature | Planeswalker;Self-Mill | Reanimate"));
});

test("captura serializada distingue resultados, precos e verificacao da LigaMagic", () => {
  const executar = new Function("document", "location", "return (" + capturarPagina.toString() + ")");
  const local = { href: "https://www.ligamagic.com.br/?view=cards/card&card=Woodwork+Prodigy&tipo=1" };
  const completo = "Woodwork Prodigy // Soul Tether";
  const destino = "https://www.ligamagic.com.br/?" + new URLSearchParams({ view: "cards/card", card: completo, tipo: "1" });
  const documento = { title: "LigaMagic", scripts: [], querySelector: () => null,
    querySelectorAll: () => [destino, destino, "https://outro.example/?view=cards/card&card=Woodwork+Prodigy", "javascript:void(0)", "#", " #zoom ", ""]
      .map((href) => ({ getAttribute: () => href })) };
  assert.deepEqual(executar(documento, local)(true), { resultados: [{ nome: completo, url: destino }] });
  assert.throws(() => executar(documento, local)(), /Carta indisponivel/);
  assert.match(executar({ ...documento, title: "Just a moment..." }, local)(true).erroCaptura, /verificacao/);
  const pagina = executar({ ...documento, scripts: [{ textContent: 'var cards_editions = [{"name":"Tarkir","price":{"0":{"p":"1.23"}}}];' }],
    querySelector: () => ({ content: completo + " | LigaMagic" }), querySelectorAll: () => { throw new Error("Nao deve procurar resultados"); } }, local)(true);
  assert.equal(pagina.nome, completo);
  assert.equal(pagina.edicoes[0].price[0].p, "1.23");
});

test("lista abre resultado exato de aventura, preserva quantidade e evita ambiguidades e ciclos", async () => {
  const nome = "Woodwork Prodigy // Soul Tether";
  const criarUrl = (carta, host = "www.ligamagic.com.br") => "https://" + host + "/?" + new URLSearchParams({ view: "cards/card", card: carta, tipo: "1" });
  const preco = { nome, capturadoEm: "2026-10-27", edicoes: [{ name: "Tarkir", price: { 0: { p: "1.23" } } }] };
  for (const cenario of ["exato", "ambiguo", "ausente", "ciclo", "verificacao", "frente"]) {
    let tempo = 0;
    const estado = { capturas: [], lote: { fase: "navegando", abaId: 7, indice: 0, prazo: 45000, intervalo: 3,
      itens: [{ nome, quantidade: 2, status: "pendente", erro: "" }] } };
    let url = criarUrl("Woodwork Prodigy");
    if (cenario === "frente") url = criarUrl(nome);
    let resposta = cenario === "verificacao" ? { erroCaptura: "Conclua a verificacao da LigaMagic." } : { resultados: [] };
    if (cenario === "exato") resposta.resultados = [{ url: criarUrl("Woodwork Prodigy", "outro.example") },
      { url: criarUrl("Soul Tether") }, { url: criarUrl(nome) }, { url: criarUrl(nome) }];
    if (cenario === "ambiguo") resposta.resultados = [{ url: criarUrl("Woodwork Prodigy") }, { url: criarUrl("Soul Tether") }];
    if (cenario === "ciclo") resposta.resultados = [{ url }];
    if (cenario === "ausente") resposta.resultados = [{ url: criarUrl("Woodwork Prodigy // Outra Aventura") }];
    const navegacoes = [];
    const api = {
      storage: { local: { get: async (chave) => structuredClone({ [chave]: estado[chave] }),
        set: async (dados) => Object.assign(estado, structuredClone(dados)) } },
      permissions: { contains: async () => true },
      tabs: { get: async () => ({ status: "complete", url }), update: async (id, dados) => { navegacoes.push(dados.url); url = dados.url; } },
      scripting: { executeScript: async ({ args }) => { assert.deepEqual(args, [true]); return [{ result: resposta }]; } },
      alarms: { create: async () => {}, clear: async () => {} },
    };
    const controle = criarControlador(api, () => tempo, async (cotacao) => ({ ...cotacao, statusTagger: "ok" }));
    await controle.processar(7);
    if (cenario !== "verificacao") {
      if (cenario !== "exato") {
        assert.equal(estado.lote.fase, "aguardando");
        assert.equal(navegacoes.length, 0);
        await controle.processar(7);
        assert.equal(navegacoes.length, 0);
        tempo = estado.lote.proxima;
        await controle.processar(7);
      }
      assert.equal(estado.lote.fase, "navegando");
      assert.equal(navegacoes[0], criarUrl(cenario === "frente" ? "Woodwork Prodigy" : nome));
      assert.equal(estado.capturas.length, 0);
      resposta = { ...preco, url };
      await controle.processar(7);
      assert.equal(estado.lote.fase, "concluido");
      assert.equal(estado.capturas[0].quantidade, 2);
      assert.ok(gerarCsvLote(estado.lote).includes(nome + ";1,23;"));
    } else {
      assert.equal(estado.lote.fase, "pausado", cenario);
      assert.equal(estado.capturas.length, 0);
      assert.equal(navegacoes.length, 0);
      assert.match(estado.lote.mensagem, /verificacao/);
    }
  }
});

test("Kyoshi retoma busca vazia ignorando Marketplace e captura pela frente sem perder resultados", async () => {
  let tempo = 0;
  const nome = "The Legend of Kyoshi // Avatar Kyoshi";
  const frente = "The Legend of Kyoshi";
  const criarUrl = (carta) => "https://www.ligamagic.com.br/?" + new URLSearchParams({ view: "cards/card", card: carta, tipo: "1" });
  let url = criarUrl(nome);
  const anterior = { id: "outra carta", nome: "Outra carta", centavos: 100, quantidade: 1 };
  const estado = { capturas: [anterior], lote: { fase: "pausado", abaId: 7, indice: 0, intervalo: 5,
    itens: [{ nome, quantidade: 1, status: "erro", erro: "A LigaMagic continua exibindo resultados." }] } };
  const navegacoes = [];
  const api = {
    storage: { local: { get: async (chave) => structuredClone({ [chave]: estado[chave] }),
      set: async (dados) => Object.assign(estado, structuredClone(dados)) } },
    permissions: { contains: async () => true },
    tabs: { get: async () => ({ status: "complete", url }), update: async (id, dados) => { navegacoes.push(dados.url); url = dados.url; } },
    scripting: { executeScript: async ({ func, args }) => {
      const detalhes = new URL(url).searchParams.get("card") === frente;
      const documento = { title: detalhes ? frente : "Busca: " + nome,
        scripts: detalhes ? [{ textContent: 'var cards_editions = [{"name":"Avatar: The Last Airbender","price":{"0":{"p":"1.23"}}}];' }] : [],
        querySelector: (seletor) => detalhes && seletor.includes("og:title") ? { content: frente + "  | Busca de Cartas" } : null,
        querySelectorAll: () => [{ getAttribute: () => "#" }] };
      const executar = new Function("document", "location", "return (" + func.toString() + ")");
      return [{ result: executar(documento, { href: url })(...args) }];
    } },
    alarms: { create: async () => {}, clear: async () => {} },
  };
  const controle = criarControlador(api, () => tempo, async (cotacao) => ({ ...cotacao, statusTagger: "ok" }));
  await controle.comando("retomar");
  assert.equal(estado.lote.fase, "aguardando");
  assert.equal(navegacoes.length, 0);
  tempo = 5000;
  await controle.processar(7);
  assert.deepEqual(navegacoes, [criarUrl(frente)]);
  assert.equal(estado.lote.fase, "navegando");
  assert.deepEqual(estado.capturas, [anterior]);
  await controle.processar(7);
  assert.equal(estado.lote.fase, "concluido");
  assert.equal(estado.lote.itens[0].status, "ok");
  assert.equal(estado.lote.itens[0].nome, nome);
  assert.equal(estado.capturas.length, 2);
  assert.deepEqual(estado.capturas[0], anterior);
  assert.equal(estado.capturas[1].centavos, 123);
  assert.ok(gerarCsvLote(estado.lote).includes(nome + ";1,23;"));
});

test("retry esgota seis buscas distintas, respeita intervalo e persiste entre controladores", async () => {
  const nome = "The Legend of Kyoshi // Avatar Kyoshi";
  const nomes = [nome, "The Legend of Kyoshi", "Avatar Kyoshi"];
  const urls = [true, false].flatMap((comTipo) => nomes.map((carta) => "https://www.ligamagic.com.br/?"
    + new URLSearchParams({ view: "cards/card", card: carta, ...(comTipo ? { tipo: "1" } : {}) })));
  for (const cenario of ["verso", "sem filtro", "vazio", "ciclo", "resultado repetido", "outra carta", "sem preco", "sem dados", "falha de leitura", "timeout", "verificacao", "permissao"]) {
    let tempo = 0;
    let url = urls[0];
    const navegacoes = [url];
    const anterior = { id: "outra", nome: "Outra", centavos: 100, quantidade: 1 };
    const estado = { capturas: [anterior], lote: { fase: "navegando", abaId: 7, indice: 0, prazo: 45000, intervalo: 5,
      itens: [{ nome, quantidade: 2, status: "pendente", erro: "" }] } };
    const sucesso = cenario === "verso" ? 2 : cenario === "sem filtro" ? 5 : -1;
    const api = {
      storage: { local: { get: async (chave) => structuredClone({ [chave]: estado[chave] }),
        set: async (dados) => Object.assign(estado, structuredClone(dados)) } },
      permissions: { contains: async () => cenario !== "permissao" },
      tabs: { get: async () => ({ status: cenario === "timeout" ? "loading" : "complete", url }),
        update: async (id, dados) => { navegacoes.push(dados.url); url = dados.url; } },
      scripting: { executeScript: async () => {
        if (navegacoes.length - 1 === sucesso) return [{ result: { nome: new URL(url).searchParams.get("card"), url,
          edicoes: [{ name: "Avatar", price: { 0: { p: "1.23" } } }] } }];
        if (cenario === "verificacao") return [{ result: { erroCaptura: "Conclua a verificacao da LigaMagic." } }];
        if (cenario === "falha de leitura") throw new Error("Dados incompletos na pagina.");
        if (cenario === "sem dados") return [];
        if (cenario === "outra carta") return [{ result: { nome: "The Legend of Kyoshi // Outra Face", edicoes: [{ price: { 0: { p: "0.01" } } }] } }];
        if (cenario === "sem preco") return [{ result: { nome, edicoes: [] } }];
        if (cenario === "resultado repetido") return [{ result: { resultados: [{ url: urls[0].replace("www.", "") + "#resultado" }] } }];
        return [{ result: { resultados: cenario === "ciclo" ? [{ url }] : [] } }];
      } },
      alarms: { create: async () => {}, clear: async () => {} },
    };
    const criar = () => criarControlador(api, () => tempo, async (cotacao) => ({ ...cotacao, statusTagger: "ok" }));
    let controle = criar();
    for (let tentativa = 0; tentativa < 8; tentativa++) {
      if (cenario === "timeout") tempo = estado.lote.prazo;
      await controle.processar(7);
      assert.deepEqual(estado.capturas[0], anterior, cenario);
      if (estado.lote.fase !== "aguardando") break;
      assert.equal(estado.capturas.length, 1, cenario);
      assert.equal(estado.lote.proxima, tempo + 5000, cenario);
      const total = navegacoes.length;
      await controle.processar(7);
      assert.equal(navegacoes.length, total, cenario);
      tempo = estado.lote.proxima;
      controle = criar();
      await controle.processar(7);
      assert.equal(estado.lote.fase, "navegando", cenario);
    }
    if (sucesso >= 0) {
      assert.deepEqual(navegacoes, urls.slice(0, sucesso + 1), cenario);
      assert.equal(estado.lote.fase, "concluido", cenario);
      assert.equal(estado.capturas[1].centavos, 123, cenario);
      assert.equal(estado.capturas[1].quantidade, 2, cenario);
      assert.ok(gerarCsvLote(estado.lote).includes(nome + ";1,23;"), cenario);
    } else {
      const bloqueado = ["verificacao", "permissao"].includes(cenario);
      assert.deepEqual(navegacoes, bloqueado ? urls.slice(0, 1) : urls, cenario);
      assert.equal(estado.lote.fase, "pausado", cenario);
      assert.equal(estado.lote.itens[0].status, "erro", cenario);
      assert.equal(estado.capturas.length, 1, cenario);
      if (!bloqueado) assert.match(estado.lote.mensagem, /Todas as 6 buscas foram esgotadas/, cenario);
      await controle.processar(7);
      assert.equal(navegacoes.length, bloqueado ? 1 : 6, cenario);
      if (cenario === "vazio") {
        await controle.comando("retomar");
        assert.equal(estado.lote.fase, "aguardando");
        assert.equal(estado.lote.itens[0].proximaBusca, urls[0]);
        assert.deepEqual(estado.capturas, [anterior]);
      }
    }
  }
});

test("CSV ManaBox agrega quantidades, le aspas e exclui proxies sem usar preco de compra", async () => {
  const { importarCsvColecao, interpretarLista } = await import("./core.mjs");
  const csv = '\uFEFFName,Quantity,Set code,Foil,Purchase price,Proxy\r\n"Grist, Voracious Larva // Grist, the Plague Swarm",2,MH3,normal,900,false\r\n"Grist, Voracious Larva // Grist, the Plague Swarm",3,MH3,foil,800,false\r\n"Card with ""quotes""",1,ABC,normal,100,false\r\nProxy Card,4,ABC,normal,999,true\r\n';
  const dados = importarCsvColecao(csv);
  assert.equal(dados.distintas, 2);
  assert.equal(dados.quantidade, 6);
  assert.equal(dados.proxiesExcluidos, 4);
  assert.deepEqual(interpretarLista(dados.texto, true).map(({ nome, quantidade }) => ({ nome, quantidade })), [
    { nome: "Grist, Voracious Larva // Grist, the Plague Swarm", quantidade: 5 },
    { nome: 'Card with "quotes"', quantidade: 1 },
  ]);
  assert.equal(importarCsvColecao('Quantity;Name\n2;"Card; Name"').texto, "2 Card; Name");
  for (const invalido of ["Name,Price\nSol Ring,1", 'Name,Quantity\n"Sol Ring,1', "Name,Quantity\nSol Ring,0", "Name,Quantity\nSol Ring,1.5", "Name,Quantity\n,1", "Name,Quantity,Proxy\nSol Ring,1,maybe", "Name,Quantity\nSol Ring,9999\nSol Ring,1"]) {
    assert.throws(() => importarCsvColecao(invalido));
  }
  const grande = "Name,Quantity\n" + Array.from({ length: 546 }, (_, indice) => "Carta " + indice + ",1").join("\n");
  assert.equal(importarCsvColecao(grande).distintas, 546);
  assert.throws(() => interpretarLista(importarCsvColecao(grande).texto), /500/);
});

test("estimativa de colecao multiplica copias, separa pendencias e exporta relatorio auditavel", async () => {
  const { resumirColecao, gerarCsvColecao } = await import("./core.mjs");
  const lote = { itens: [
    { nome: "Sol Ring", quantidade: 3, status: "ok", cotacao: { centavos: 123, quantidade: 1, edicao: "Commander", variante: "Normal", capturadoEm: "2026-11-02" } },
    { nome: "Pendente", quantidade: 2, status: "pendente" },
    { nome: "Sem preco", quantidade: 1, status: "pulado", erro: "Sem ofertas" },
  ] };
  assert.deepEqual(resumirColecao(lote), { distintas: 3, unidades: 6, avaliadas: 3, pendentes: 3, centavos: 369 });
  const csv = gerarCsvColecao(lote);
  assert.ok(csv.includes("Sol Ring;3;1,23;3,69;Avaliada;Commander;Normal;2026-11-02"));
  assert.ok(csv.includes("Pendente;2;;;Pendente;"));
  assert.ok(csv.includes("Sem preco;1;;;Ignorada;"));
  assert.ok(csv.includes("sem ajuste por edicao"));
  assert.ok(csv.startsWith("\uFEFF"));
  assert.ok(gerarCsvColecao({ itens: [{ nome: "=Formula", quantidade: 1, status: "erro", erro: "+cmd" }] }).includes("'=Formula"));
});

test("controlador aceita colecao com 546 nomes e preserva modo ao reiniciar", async () => {
  const { importarCsvColecao } = await import("./core.mjs");
  const dados = importarCsvColecao("Name,Quantity\n" + Array.from({ length: 546 }, (_, indice) => "Carta " + indice + ",1").join("\n"));
  const estado = {};
  const api = {
    storage: { local: { get: async (chave) => structuredClone({ [chave]: estado[chave] }), set: async (valores) => Object.assign(estado, structuredClone(valores)) } },
    permissions: { contains: async () => true },
    tabs: { create: async () => ({ id: 7 }), update: async () => {} },
    alarms: { create: async () => {}, clear: async () => {} },
  };
  await criarControlador(api, () => 0).comando("iniciar", { texto: dados.texto, intervalo: 5, colecao: { arquivo: "minhas cartas.csv", proxiesExcluidos: 0 } });
  assert.equal(estado.lote.itens.length, 546);
  assert.equal(estado.lote.colecao.arquivo, "minhas cartas.csv");
  await criarControlador(api, () => 0).reiniciar();
  assert.equal(estado.lote.fase, "pausado");
  assert.equal(estado.lote.colecao.arquivo, "minhas cartas.csv");
  assert.equal(estado.lote.itens.length, 546);
});

test("CSV preserva edicao acabamento e idioma sem misturar variantes da colecao", async () => {
  const { importarCsvColecao } = await import("./core.mjs");
  const dados = importarCsvColecao("Name,Quantity,Set code,Set name,Collector number,Foil,Language\nSol Ring,2,CMM,Commander Masters,410,normal,en\nSol Ring,3,CMM,Commander Masters,410,foil,en\nSol Ring,1,CMM,Commander Masters,410,normal,pt\nSol Ring,1,CMM,Commander Masters,410,normal,en");
  assert.equal(dados.quantidade, 7);
  assert.equal(dados.distintas, 1);
  assert.equal(dados.variantes.length, 3);
  assert.deepEqual(dados.variantes[0], { nome: "Sol Ring", quantidade: 3, codigoEdicao: "cmm", edicao: "Commander Masters", numero: "410", acabamento: "normal", idioma: "en" });
  assert.equal(dados.variantes[1].acabamento, "foil");
  assert.equal(dados.variantes[2].idioma, "pt");
});

test("media da colecao usa valor amarelo da edicao numero e acabamento sem depender de ofertas legiveis", async () => {
  const { importarCsvColecao, prepararColecao, extrairCotacaoColecao } = await import("./core.mjs");
  const [item] = prepararColecao(importarCsvColecao("Name,Quantity,Set code,Set name,Collector number,Foil,Language\nMeteor Golem,2,FDN,Foundations,256,normal,en").variantes);
  const oferta = { sellType: 1, idEdicao: 480601, num: "256", idioma: 2, extras: 0, precoFinal: "2.00" };
  const pagina = { nome: item.nome, edicoes: [{ id: 480601, code: "fdn", num: "256", name: "Foundations", price: { 0: { p: "0.13", m: "0.71", g: "2.50" }, 2: { p: "0.50", m: "1.33", g: "2.99" }, 31: { m: "2.45" } } }],
    idiomas: [{ id: 2, acron: "EN" }, { id: 8, acron: "PT" }, { id: 11, acron: "PTEN" }], extras: [{ id: 2, label: "Foil" }, { id: 31, label: "Foil Especial / Foil Etched" }],
    ofertas: [oferta, { ...oferta, precoFinal: "4.00" }, { ...oferta, idioma: 8, precoFinal: "90.00" }, { ...oferta, idioma: 11, precoFinal: "90.00" },
      { ...oferta, extras: 2, precoFinal: "80.00" }, { ...oferta, num: "999", precoFinal: "80.00" }, { ...oferta, idEdicao: 1, precoFinal: "80.00" },
      { ...oferta, is_graded: 1, precoFinal: "80.00" }, { ...oferta, sellType: 2, precoFinal: "80.00" }] };
  const cotacao = extrairCotacaoColecao(pagina, item);
  assert.equal(cotacao.centavos, 71);
  assert.equal(cotacao.idiomaPreco, "geral");
  assert.equal(cotacao.id, item.idColecao);
  assert.equal(cotacao.tipoPreco, "media_ligamagic");
  assert.equal(extrairCotacaoColecao(pagina, { ...item, acabamento: "foil" }).centavos, 133);
  assert.equal(extrairCotacaoColecao(pagina, { ...item, acabamento: "etched" }).centavos, 245);
  const portugues = extrairCotacaoColecao(pagina, { ...item, idioma: "pt" });
  assert.equal(portugues.centavos, 71);
  assert.equal(portugues.idioma, "pt");
  assert.equal(extrairCotacaoColecao({ ...pagina, ofertas: [{ ...oferta, precoFinal: undefined, precoCss: "imagem" }] }, item).centavos, 71);
  assert.throws(() => extrairCotacaoColecao(pagina, { ...item, codigoEdicao: "m20" }), /Nenhuma outra impressao/);
  assert.throws(() => extrairCotacaoColecao(pagina, { ...item, numero: "999" }), /Nenhuma outra impressao/);
  assert.equal(extrairCotacaoColecao({ ...pagina, ofertas: [], idiomas: [] }, item).centavos, 71);
  for (const m of [undefined, "", "0.00", "-1", "abc", "1.234"]) {
    assert.throws(() => extrairCotacaoColecao({ ...pagina, edicoes: [{ ...pagina.edicoes[0], price: { 0: { p: "0.13", g: "2.50", m } } }] }, item), /amarelo indisponivel/);
  }
  assert.throws(() => prepararColecao([{ nome: "Sol Ring", quantidade: 1 }]), /Reimporte/);
});

test("edicao exata precede siglas prefixadas confirmadas na cotacao da colecao", async () => {
  const { extrairCotacaoColecao } = await import("./core.mjs");
  const item = { nome: "Cemetery Recruitment", quantidade: 1, codigoEdicao: "fdn", edicao: "Foundations", numero: "517", acabamento: "normal", idioma: "en",
    tratamento: "padrao", referenciasTerreno: [{ codigoEdicao: "fdn", numero: "517", tratamento: "padrao" }] };
  const edicao = (code, m) => ({ code, num: "517", name: "Foundations", price: { 0: { m } } });
  const pagina = { nome: item.nome, edicoes: [edicao("fdnx", "9.00"), edicao("bbfdn", "2.00"), edicao("fdn", "1.00")] };
  assert.equal(extrairCotacaoColecao(pagina, item).centavos, 100);
  assert.equal(extrairCotacaoColecao({ ...pagina, edicoes: pagina.edicoes.slice(0, 2) }, item).centavos, 200);
  assert.throws(() => extrairCotacaoColecao({ ...pagina, edicoes: pagina.edicoes.slice(0, 1) }, item), /Nenhuma outra impressao/);
});

test("terreno comum usa lancamento mais recente e foil full art preservam versao antes de aproximar", async () => {
  const { extrairCotacaoColecao } = await import("./core.mjs");
  const item = { nome: "Swamp", quantidade: 6, codigoEdicao: "eld", edicao: "Throne of Eldraine", numero: "261", acabamento: "normal", idioma: "en", fullArt: false };
  const edicoes = [
    { code: "eld", num: "261", name: "Throne of Eldraine", dataLancamento: "2019-10-04", fullArt: false, price: { 0: { m: "0.60" }, 2: { m: "2.00" } } },
    { code: "new", num: "1", name: "Latest", dataLancamento: "2026-01-01", fullArt: false, price: { 0: { m: "0.20" }, 2: { m: "1.00" } } },
    { code: "art", num: "2", name: "Full Art", dataLancamento: "2026-02-01", fullArt: true, price: { 0: { m: "8.00" }, 2: { m: "10.00" } } },
    { code: "future", num: "3", name: "Future", dataLancamento: "2099-01-01", fullArt: false, price: { 0: { m: "9.00" } } },
  ];
  const pagina = { nome: "Swamp", edicoes, extras: [{ id: 2, label: "Foil" }] };
  assert.equal(extrairCotacaoColecao(pagina, item).codigoEdicao, "new");
  assert.equal(extrairCotacaoColecao(pagina, item).centavos, 20);
  assert.equal(extrairCotacaoColecao(pagina, { ...item, acabamento: "foil" }).centavos, 200);
  assert.equal(extrairCotacaoColecao(pagina, { ...item, fullArt: true, codigoEdicao: "art", numero: "2" }).centavos, 800);
  assert.equal(extrairCotacaoColecao({ ...pagina, edicoes: edicoes.slice(1) }, { ...item, acabamento: "foil" }).centavos, 100);
  assert.equal(extrairCotacaoColecao(pagina, { ...item, fullArt: true, codigoEdicao: "missing", numero: "99" }).centavos, 800);
  assert.equal(extrairCotacaoColecao({ nome: item.nome, edicoes: [{ code: "starter", num: "517", name: "Beginner Foundations Box", price: { 0: { m: "0.71" } } }] },
    { ...item, nome: "Swamp", codigoEdicao: "fdn", edicao: "Foundations", numero: "517", fullArt: undefined }).centavos, 71);
});

test("Scryfall confirma full art e fornece datas reais sem usar cadastro da LigaMagic como lancamento", async () => {
  const { consultarReferenciasTerreno } = await import("./core.mjs");
  const chamadas = [];
  const carta = { name: "Swamp", set: "eld", collector_number: "261", full_art: false, finishes: ["nonfoil"], released_at: "2019-10-04" };
  const requisitar = async (url) => { chamadas.push(url); return { ok: true, json: async () => url.includes("/search?") ? { data: [
    { ...carta, set: "new", collector_number: "1", released_at: "2026-01-01" }, { ...carta, full_art: true }, { ...carta, digital: true },
  ] } : carta }; };
  const cache = new Map();
  const item = { nome: "Swamp", codigoEdicao: "eld", numero: "261", acabamento: "normal" };
  const dados = await consultarReferenciasTerreno(item, requisitar, cache);
  assert.equal(dados.fullArt, false);
  assert.equal(dados.referenciasTerreno.length, 2);
  assert.equal(dados.referenciasTerreno[1].dataLancamento, "2026-01-01");
  await consultarReferenciasTerreno(item, requisitar, cache);
  assert.equal(chamadas.length, 2);
});

test("edicao e numero identificam arte de cartas nao terreno sem um limite numerico universal", async () => {
  const { consultarReferenciasColecao, extrairCotacaoColecao } = await import("./core.mjs");
  const item = { nome: "Jace's Machinations", codigoEdicao: "fra", edicao: "Reality Fracture", numero: "337", acabamento: "normal", idioma: "en", quantidade: 1,
    scryfallId: "aabdc0a6-5d60-44c4-96cb-692eeeac0630" };
  const carta = { name: item.nome, set: "fra", collector_number: "337", full_art: false, border_color: "borderless", frame_effects: ["showcase"],
    finishes: ["nonfoil"], released_at: "2026-10-02" };
  const comum = { ...carta, collector_number: "100", border_color: "black", frame_effects: [] };
  const alternativa = { ...carta, set: "promo", collector_number: "10" };
  const chamadas = [];
  const requisitar = async (url) => { chamadas.push(url); return { ok: true, json: async () => url.includes("/search?") ? { data: [comum, alternativa] } : carta }; };
  const dados = await consultarReferenciasColecao(item, requisitar);
  assert.equal(new URL(chamadas[0]).pathname, "/cards/fra/337");
  assert.equal(dados.fullArt, false);
  assert.equal(dados.tratamento, "borderless");
  const edicao = (code, num, m) => ({ code, num, name: code, price: { 0: { m } } });
  const pagina = { nome: item.nome, edicoes: [edicao("fra", "100", "0.10"), edicao("promo", "10", "4.00"), edicao("fra", "337", "5.73")] };
  assert.equal(extrairCotacaoColecao(pagina, { ...item, ...dados }).numero, "337");
  assert.throws(() => extrairCotacaoColecao({ ...pagina, edicoes: pagina.edicoes.slice(0, 2) }, { ...item, ...dados }), /Nenhuma outra impressao/);
  assert.throws(() => extrairCotacaoColecao({ ...pagina, edicoes: pagina.edicoes.slice(0, 1) }, { ...item, ...dados }), /indisponivel/);
  const dadosComuns = await consultarReferenciasColecao({ ...item, codigoEdicao: "other" }, async (url) => ({ ok: true,
    json: async () => url.includes("/search?") ? { data: [] } : { ...comum, set: "other", collector_number: "337" } }));
  assert.equal(dadosComuns.fullArt, false);
  assert.equal(dadosComuns.tratamento, "padrao");
  for (const [tratamento, campos] of [["fullart", { full_art: true }], ["extendedart", { frame_effects: ["extendedart"] }]]) {
    const dadosArte = await consultarReferenciasColecao(item, async (url) => ({ ok: true,
      json: async () => url.includes("/search?") ? { data: [] } : { ...comum, collector_number: "337", ...campos } }));
    assert.equal(dadosArte.tratamento, tratamento);
  }
  await assert.rejects(consultarReferenciasColecao(item, async () => ({ ok: true, json: async () => comum })), /numero diferente/);
});

test("colecao associa CBZNR SSFRA e SCFDN as impressoes confirmadas no Scryfall", async () => {
  const { consultarReferenciasColecao, extrairCotacaoColecao } = await import("./core.mjs");
  for (const caso of [
    { nome: "Canyon Jerboa", codigo: "znr", alias: "cbznr", numero: "290", comum: "7", media: "1.09",
      arte: { border_color: "black", frame_effects: ["showcase", "inverted"] } },
    { nome: "Jace's Machinations", codigo: "fra", alias: "ssfra", numero: "337", comum: "32", media: "24.53",
      arte: { border_color: "borderless", frame_effects: ["showcase"] } },
    { nome: "Chart a Course", codigo: "fdn", alias: "scfdn", numero: "586", comum: "999", media: "0.98",
      arte: { border_color: "black" }, finishes: ["nonfoil"] },
    { nome: "Linden, the Steadfast Queen", codigo: "fdn", alias: "scfdn", numero: "577", comum: "737", media: "2.09",
      arte: { border_color: "black", frame_effects: ["legendary"] }, finishes: ["nonfoil"] },
  ]) {
    const item = { nome: caso.nome, codigoEdicao: caso.codigo, edicao: "Edicao CSV", numero: caso.numero,
      acabamento: "normal", idioma: "en", quantidade: 1 };
    const carta = { name: item.nome, set: caso.codigo, collector_number: caso.numero, full_art: false,
      finishes: caso.finishes || ["nonfoil", "foil"], ...caso.arte };
    const dados = await consultarReferenciasColecao(item, async (url) => ({ ok: true,
      json: async () => url.includes("/search?") ? { data: [] } : carta }));
    const edicao = (code, num, m) => ({ code, num, name: code, price: caso.alias === "scfdn" ? [{ m }] : { 0: { m } } });
    const pagina = { nome: item.nome, edicoes: [edicao(caso.codigo, caso.comum, "0.01"),
      edicao(caso.alias, caso.numero, caso.media), edicao("x" + caso.codigo + "x", caso.numero, "0.02")] };
    const preco = extrairCotacaoColecao(pagina, { ...item, ...dados });
    assert.equal(preco.codigoEdicao, caso.alias);
    assert.equal(preco.numero, caso.numero);
    assert.equal(preco.centavos, Math.round(Number(caso.media) * 100));
    assert.throws(() => extrairCotacaoColecao({ ...pagina, edicoes: [pagina.edicoes[0], pagina.edicoes[2]] },
      { ...item, ...dados }), /Nenhuma outra impressao/);
    assert.throws(() => extrairCotacaoColecao(pagina, { ...item, ...dados, referenciasTerreno: [] }), /Nenhuma outra impressao/);
  }
});

test("colecao aceita prefixos arbitrarios apenas com impressao e tratamento confirmados", async () => {
  const { extrairCotacaoColecao } = await import("./core.mjs");
  const item = { nome: "Chart a Course", quantidade: 1, codigoEdicao: "fdn", edicao: "Foundations", numero: "586",
    acabamento: "normal", idioma: "en", fullArt: false, tratamento: "padrao",
    referenciasTerreno: [{ codigoEdicao: "fdn", numero: "586", fullArt: false, tratamento: "padrao" }] };
  const edicao = (code, num = "586") => ({ code, num, name: "Foundations (Variante)", price: { 0: { m: "0.98" } } });
  for (const prefixo of ["bb", "cb", "bl", "ss", "bi", "sc", "wd", "novo", "tratamento2027"]) {
    const pagina = { nome: item.nome, edicoes: [edicao(prefixo + "fdn")] };
    assert.equal(extrairCotacaoColecao(pagina, item).centavos, 98, prefixo);
    assert.throws(() => extrairCotacaoColecao(pagina, { ...item, referenciasTerreno: [] }), /Nenhuma outra impressao/);
    assert.throws(() => extrairCotacaoColecao(pagina, { ...item, tratamento: "borderless" }), /Nenhuma outra impressao/);
    assert.throws(() => extrairCotacaoColecao({ ...pagina, edicoes: [edicao(prefixo + "fdn", "587")] }, item), /Nenhuma outra impressao/);
  }
  for (const code of ["fdnextra", "scfdnextra", "mb2", "scm21"]) {
    assert.throws(() => extrairCotacaoColecao({ nome: item.nome, edicoes: [edicao(code)] }, item), /Nenhuma outra impressao/);
  }
  const outraImpressao = { codigoEdicao: "pfdn", numero: "586", fullArt: false, tratamento: "padrao" };
  assert.throws(() => extrairCotacaoColecao({ nome: item.nome, edicoes: [edicao("pfdn")] },
    { ...item, referenciasTerreno: [...item.referenciasTerreno, outraImpressao] }), /Nenhuma outra impressao/);
});

test("colecao preserva Tranquil Cove FDN 270 e Swiftwater Cliffs FDN 268 sem regra de terreno basico", async () => {
  const { consultarReferenciasColecao, extrairCotacaoColecao, terrenoBasico } = await import("./core.mjs");
  for (const caso of [
    { nome: "Tranquil Cove", numero: "270", minimo: "0.10", media: "0.56", foil: "0.98" },
    { nome: "Swiftwater Cliffs", numero: "268", minimo: "0.09", media: "0.57", foil: "0.99" },
  ]) {
    const item = { nome: caso.nome, codigoEdicao: "fdn", edicao: "Foundations", numero: caso.numero,
      acabamento: "normal", idioma: "en", quantidade: 1 };
    assert.equal(terrenoBasico(item.nome), false);
    const carta = { name: item.nome, set: "fdn", collector_number: caso.numero, full_art: false,
      border_color: "black", finishes: ["nonfoil", "foil"] };
    const dados = await consultarReferenciasColecao(item, async (url) => ({ ok: true, json: async () => {
      if (url.includes("/search?")) return { data: [] };
      assert.equal(new URL(url).pathname, "/cards/fdn/" + caso.numero);
      return carta;
    } }));
    const pagina = { nome: item.nome, extras: [{ id: 2, label: "Foil" }], edicoes: [
      { code: "other", num: caso.numero, name: "Outra edicao", price: { 0: { m: "0.01" } } },
      { id: 480601, code: "fdn", num: caso.numero, name: "Foundations",
        price: { 0: { p: caso.minimo, m: caso.media }, 2: { m: caso.foil } } },
    ] };
    const preco = extrairCotacaoColecao(pagina, { ...item, ...dados });
    assert.equal(preco.codigoEdicao, "fdn");
    assert.equal(preco.numero, caso.numero);
    assert.equal(preco.centavos, Math.round(Number(caso.media) * 100));
    assert.equal(preco.referenciaSubstituida, false);
    assert.equal(extrairCotacaoColecao(pagina, { ...item, ...dados, acabamento: "foil" }).centavos, Math.round(Number(caso.foil) * 100));
    assert.throws(() => extrairCotacaoColecao({ ...pagina, edicoes: pagina.edicoes.slice(0, 1) }, { ...item, ...dados }), /Nenhuma outra impressao/);
  }
});

test("colecao nao substitui An Offer You Can't Refuse FDN 160 foil por MB2 261", async () => {
  const { extrairCotacaoColecao } = await import("./core.mjs");
  const item = { nome: "An Offer You Can't Refuse", codigoEdicao: "fdn", edicao: "Foundations", numero: "160",
    acabamento: "foil", idioma: "en", quantidade: 1, fullArt: false, tratamento: "padrao",
    referenciasTerreno: [{ codigoEdicao: "fdn", numero: "160", fullArt: false, tratamento: "padrao" },
      { codigoEdicao: "mb2", numero: "261", fullArt: false, tratamento: "padrao" }] };
  const errada = { code: "mb2", num: "261", name: "Mystery Booster 2", price: { 2: { m: "99.00" } } };
  const correta = { code: "fdn", num: "160", name: "Foundations", price: { 0: { m: "1.00" }, 2: { m: "25.00" } } };
  const pagina = { nome: item.nome, extras: [{ id: 2, label: "Foil" }], edicoes: [errada, correta] };
  const preco = extrairCotacaoColecao(pagina, item);
  assert.equal(preco.codigoEdicao, "fdn");
  assert.equal(preco.numero, "160");
  assert.equal(preco.centavos, 2500);
  assert.equal(preco.referenciaSubstituida, false);
  for (const edicoes of [[errada], [errada, { ...correta, price: { 0: { m: "1.00" } } }],
    [errada, { ...correta, num: "161" }]]) {
    assert.throws(() => extrairCotacaoColecao({ ...pagina, edicoes }, item), /FDN #160 \/ Foil.*Nenhuma outra impressao/);
  }
});

test("colecao confirma Quicksilver Fisher ONE 287 Showcase Ichor sem confundir normal com arte padrao", async () => {
  const { consultarReferenciasColecao, extrairCotacaoColecao } = await import("./core.mjs");
  const item = { nome: "Quicksilver Fisher", codigoEdicao: "one", edicao: "Phyrexia: All Will Be One", numero: "287",
    acabamento: "normal", idioma: "en", quantidade: 1 };
  const carta = { name: item.nome, set: "one", collector_number: "287", full_art: false, border_color: "black",
    frame_effects: ["showcase", "inverted"], finishes: ["nonfoil", "foil"] };
  const comum = { ...carta, collector_number: "66", frame_effects: [] };
  const dados = await consultarReferenciasColecao(item, async (url) => ({ ok: true, json: async () => {
    if (url.includes("/search?")) return { data: [comum] };
    assert.equal(new URL(url).pathname, "/cards/one/287");
    return carta;
  } }));
  assert.equal(dados.tratamento, "showcase:inverted,showcase");
  const edicao = (code, num, name, m) => ({ code, num, name, price: { 0: { m } } });
  const pagina = { nome: item.nome, edicoes: [edicao("one", "66", item.edicao, "0.10"),
    edicao("bione", "287", item.edicao + " (Borderless Ichor)", "0.54")] };
  const preco = extrairCotacaoColecao(pagina, { ...item, ...dados });
  assert.equal(preco.codigoEdicao, "bione");
  assert.equal(preco.numero, "287");
  assert.equal(preco.centavos, 54);
  assert.equal(preco.variante, "Normal");
  assert.throws(() => extrairCotacaoColecao({ ...pagina, edicoes: pagina.edicoes.slice(0, 1) }, { ...item, ...dados }), /indisponivel/);
  await assert.rejects(consultarReferenciasColecao({ ...item, acabamento: "etched" }, async () => ({ ok: true, json: async () => carta })), /acabamento/);
});

test("terreno de lista comum consulta por nome sem exigir edicao e nao herda Full Art da resposta", async () => {
  const { consultarReferenciasTerreno } = await import("./core.mjs");
  const chamadas = [];
  const carta = { name: "Swamp", full_art: true, set: "art", collector_number: "1", finishes: ["nonfoil"], released_at: "2026-01-01" };
  const requisitar = async (url) => {
    chamadas.push(url);
    return { ok: true, json: async () => url.includes("/search?") ? { data: [
      { ...carta, full_art: false, set: "new", collector_number: "400" }, carta,
    ] } : carta };
  };
  const dados = await consultarReferenciasTerreno({ nome: "Swamp" }, requisitar);
  assert.equal(new URL(chamadas[0]).pathname, "/cards/named");
  assert.equal(new URL(chamadas[0]).searchParams.get("exact"), "Swamp");
  assert.ok(new URL(chamadas[1]).searchParams.get("q").includes("-is:fullart"));
  assert.equal(dados.fullArt, false);
  assert.deepEqual(dados.referenciasTerreno, [{ codigoEdicao: "new", numero: "400", fullArt: false, dataLancamento: "2026-01-01" }]);
});

test("lote de variantes preserva criterios, calcula media e segue com pendencias sem capturas antigas", async () => {
  const { importarCsvColecao, resumirColecao, gerarCsvColecao } = await import("./core.mjs");
  const dados = importarCsvColecao("Name,Quantity,Set code,Set name,Collector number,Foil,Language\nMeteor Golem,2,FDN,Foundations,256,normal,en\nMeteor Golem,3,FDN,Foundations,256,foil,pt");
  const estado = { capturas: [{ id: "anterior", centavos: 100, quantidade: 1 }] };
  let url;
  let instante = 0;
  const pagina = { nome: "Meteor Golem", url: "https://www.ligamagic.com.br/", edicoes: [{ id: 1, code: "fdn", num: "256", name: "Foundations", price: { 0: { p: "0.13", m: "0.71", g: "2.50" } } }],
    idiomas: [{ id: 2, acron: "EN" }, { id: 8, acron: "PT" }], extras: [{ id: 2, label: "Foil" }],
    ofertas: [{ sellType: 1, idEdicao: 1, num: "256", idioma: 2, extras: 0, precoFinal: "2.00" }, { sellType: 1, idEdicao: 1, num: "256", idioma: 2, extras: 0, precoFinal: "4.00" }] };
  const api = {
    storage: { local: { get: async (chave) => structuredClone({ [chave]: estado[chave] }), set: async (valores) => Object.assign(estado, structuredClone(valores)) } },
    permissions: { contains: async () => true },
    tabs: { create: async () => ({ id: 7 }), get: async () => ({ status: "complete", url }), update: async (id, dados) => { url = dados.url; } },
    scripting: { executeScript: async ({ args }) => { assert.deepEqual(args, [true, true]); return [{ result: pagina }]; } },
    alarms: { create: async () => {}, clear: async () => {} },
  };
  const controle = criarControlador(api, () => instante, async () => { throw new Error("Colecao nao precisa consultar categorias"); },
    async () => ({ fullArt: false, tratamento: "padrao", referenciasTerreno: [{ codigoEdicao: "fdn", numero: "256", fullArt: false, tratamento: "padrao" }] }));
  await controle.comando("iniciar", { texto: dados.texto, intervalo: 5, somenteNormal: true, colecao: { arquivo: "cartas.csv", criterios: "variantes", variantes: dados.variantes } });
  assert.equal(estado.lote.itens.length, 2);
  assert.equal(estado.lote.somenteNormal, false);
  await controle.processar(7);
  assert.equal(estado.lote.itens[0].cotacao.centavos, 71);
  instante = 5000;
  await controle.processar();
  await controle.processar(7);
  assert.equal(estado.lote.fase, "concluido");
  assert.match(estado.lote.itens[1].erro, /amarelo indisponivel/);
  assert.deepEqual(resumirColecao(estado.lote), { distintas: 2, unidades: 5, avaliadas: 2, pendentes: 3, centavos: 142 });
  assert.deepEqual(estado.capturas, []);
  const relatorio = gerarCsvColecao(estado.lote);
  assert.ok(relatorio.includes("Preco medio unitario (BRL)"));
  assert.ok(relatorio.includes("Meteor Golem;2;Foundations;fdn;256;normal;en;0,71;1,42;media_ligamagic;Avaliada"));
  assert.ok(relatorio.includes("sem filtro por idioma"));
  assert.ok(relatorio.includes("Meteor Golem;3;Foundations;fdn;256;foil;pt;;;;Pendente"));
  const normalSalva = structuredClone(estado.lote.itens[0]);
  pagina.edicoes[0].price[2] = { m: "1.33" };
  await controle.comando("retentar");
  instante += 5000;
  await controle.processar();
  await controle.processar(7);
  assert.equal(estado.lote.fase, "concluido");
  assert.deepEqual(estado.lote.itens[0], normalSalva);
  assert.equal(estado.lote.itens[1].status, "ok");
  assert.equal(estado.lote.itens[1].cotacao.centavos, 133);
  assert.deepEqual(resumirColecao(estado.lote), { distintas: 2, unidades: 5, avaliadas: 5, pendentes: 0, centavos: 541 });
  assert.deepEqual(estado.capturas, []);
  await assert.rejects(controle.comando("retentar"), /Nenhuma pendencia/);
});

test("retentativa preserva cartas validas intercaladas em listas concluidas canceladas ou pausadas", async () => {
  for (const fase of ["concluido", "cancelado", "pausado"]) {
    let tempo = 0;
    let url;
    const pronta = (nome) => ({ nome, quantidade: 2, status: "ok", cotacao: { id: nome, nome, centavos: 200, quantidade: 2 } });
    const itens = [{ nome: "Falta A", quantidade: 1, status: "erro", erro: "Falhou" }, pronta("Pronta A"),
      { nome: "Falta B", quantidade: 3, status: "pulado", buscasTentadas: ["antiga"] }, pronta("Pronta B")];
    const validas = structuredClone([itens[1], itens[3]]);
    const estado = { lote: { fase, indice: 4, abaId: null, intervalo: 4, somenteNormal: true, itens }, capturas: validas.map((item) => item.cotacao) };
    const navegacoes = [];
    const api = {
      storage: { local: { get: async (chave) => structuredClone({ [chave]: estado[chave] }), set: async (dados) => Object.assign(estado, structuredClone(dados)) } },
      permissions: { contains: async () => true },
      tabs: { create: async () => ({ id: 7 }), get: async () => ({ status: "complete", url }), update: async (id, dados) => {
        url = dados.url;
        navegacoes.push(new URL(url).searchParams.get("card"));
      } },
      scripting: { executeScript: async () => [{ result: { nome: new URL(url).searchParams.get("card"), url,
        edicoes: [{ name: "Set", price: { 0: { p: "1.00" } } }] } }] },
      alarms: { create: async () => {}, clear: async () => {} },
    };
    const controle = criarControlador(api, () => tempo, async (cotacao) => ({ ...cotacao, statusTagger: "ok" }));
    await controle.comando("retentar");
    await assert.rejects(controle.comando("retentar"), /Pause ou conclua/);
    await controle.processar(7);
    assert.equal(estado.lote.indice, 2);
    assert.deepEqual(navegacoes, ["Falta A"]);
    tempo = 4000;
    await controle.processar();
    await controle.processar(7);
    assert.equal(estado.lote.fase, "concluido", fase);
    assert.deepEqual(navegacoes, ["Falta A", "Falta B"]);
    assert.deepEqual([estado.lote.itens[1], estado.lote.itens[3]], validas);
    assert.equal(estado.capturas.length, 4);
    assert.equal(estado.capturas.find((item) => item.nome === "Falta B").quantidade, 3);
  }
});

test("lote confirma arte de FRA 337 pelo numero e preserva a identificacao ao reiniciar", async () => {
  const { consultarReferenciasColecao } = await import("./core.mjs");
  const item = { nome: "Jace's Machinations", codigoEdicao: "fra", edicao: "Reality Fracture", numero: "337", acabamento: "normal", idioma: "en", quantidade: 1 };
  const estado = {};
  let url;
  let consultas = 0;
  let falharLeitura = true;
  const api = {
    storage: { local: { get: async (chave) => structuredClone({ [chave]: estado[chave] }), set: async (dados) => Object.assign(estado, structuredClone(dados)) } },
    permissions: { contains: async () => true },
    tabs: { create: async () => ({ id: 7 }), get: async () => ({ status: "complete", url }), update: async (id, dados) => { url = dados.url; } },
    scripting: { executeScript: async () => {
      if (falharLeitura) throw new Error("Falha temporaria de leitura");
      return [{ result: { nome: item.nome, url, edicoes: [
        { code: "fra", num: "100", name: item.edicao, price: { 0: { m: "0.10" } } },
        { code: "fra", num: "337", name: item.edicao, price: { 0: { m: "5.73" } } },
      ] } }];
    } },
    alarms: { create: async () => {}, clear: async () => {} },
  };
  const resolver = async (criterio) => {
    consultas++;
    return consultarReferenciasColecao(criterio, async (endereco) => ({ ok: true, json: async () => endereco.includes("/search?") ? { data: [] }
      : { name: item.nome, set: "fra", collector_number: "337", full_art: false, border_color: "borderless", frame_effects: ["showcase"], finishes: ["nonfoil"] } }));
  };
  const controle = criarControlador(api, () => 0, undefined, resolver);
  await controle.comando("iniciar", { intervalo: 5, colecao: { criterios: "variantes", variantes: [item] } });
  await controle.processar(7);
  assert.equal(estado.lote.itens[0].fullArt, false);
  assert.equal(estado.lote.itens[0].tratamento, "borderless");
  falharLeitura = false;
  const retomado = criarControlador(api, () => 0, undefined, resolver);
  await retomado.comando("retomar");
  await retomado.processar(7);
  assert.equal(consultas, 1);
  assert.equal(estado.lote.fase, "concluido");
  assert.equal(estado.lote.itens[0].cotacao.numero, "337");
  assert.equal(estado.lote.itens[0].cotacao.centavos, 573);
  assert.equal(estado.lote.itens[0].cotacao.fullArt, false);
  assert.equal(estado.lote.itens[0].cotacao.tratamento, "borderless");
});

test("captura de Swamp extrai URL ed num do cardsjson e preserva URL especifica na pagina de precos", () => {
  const executar = new Function("document", "location", "return (" + capturarPagina.toString() + ")");
  const carta = { nEN: "Swamp (#261)", sSigla: "eld", idE: 479771, sN: "261", p1b: "1.03" };
  const documento = { title: "Pantano | LigaMagic", scripts: [{ textContent: "var cardsjson = " + JSON.stringify([carta]) + ";" }],
    querySelector: () => null, querySelectorAll: () => [] };
  const resultado = executar(documento, { href: "https://www.ligamagic.com.br/?view=cards/card&card=Swamp" })(true, true);
  const referencia = resultado.edicoesResultados[0];
  const url = new URL(referencia.url);
  assert.equal(url.searchParams.get("ed"), "479771");
  assert.equal(url.searchParams.get("num"), "261");
  assert.equal(referencia.price[0].m, "1.03");
  const detalhes = { ...documento, scripts: [{ textContent: 'var cards_editions = [{"code":"eld","num":"261","price":{"0":{"m":"1.03"}}}];' }],
    querySelector: () => ({ content: "Pantano / Swamp | Busca de Cartas" }) };
  assert.equal(executar(detalhes, { href: url.href })(true, true).url, url.href);
});

test("colecao navega da galeria para FDN exata de Tranquil Cove e Swiftwater Cliffs", async () => {
  const executar = new Function("document", "location", "return (" + capturarPagina.toString() + ")");
  for (const caso of [{ nome: "Tranquil Cove", numero: "270", media: "0.56" },
    { nome: "Swiftwater Cliffs", numero: "268", media: "0.57" }]) {
    for (const { acabamento, prefixo } of ["normal", "foil"].flatMap(acabamento => ["", "novo"].map(prefixo => ({ acabamento, prefixo })))) {
      const estado = {};
      let tempo = 0;
      let url;
      const navegacoes = [];
      const item = { nome: caso.nome, quantidade: 1, codigoEdicao: "fdn", edicao: "Foundations", numero: caso.numero, acabamento, idioma: "en" };
      const endereco = (ed, num) => "https://www.ligamagic.com.br/?" + new URLSearchParams({ view: "cards/card", card: caso.nome, ed, num });
      const correta = endereco("480601", caso.numero);
      const errada = endereco("999999", "999");
      const cartas = [{ nEN: caso.nome, sSigla: prefixo + "fdn", idE: 480601, sN: caso.numero },
        { nEN: caso.nome, sSigla: "other", idE: 999999, sN: "999", p1b: "0.01" }];
      const paginaPrecos = { code: prefixo + "fdn", num: caso.numero, name: "Foundations", price: { 0: { m: caso.media }, 2: { m: "0.99" } } };
      const api = {
        storage: { local: { get: async (chave) => structuredClone({ [chave]: estado[chave] }),
          set: async (dados) => Object.assign(estado, structuredClone(dados)) } },
        permissions: { contains: async () => true },
        tabs: { create: async () => ({ id: 7 }), get: async () => ({ status: "complete", url }),
          update: async (id, dados) => { url = dados.url; navegacoes.push(url); } },
        scripting: { executeScript: async ({ args }) => {
          const detalhes = new URL(url).searchParams.has("ed");
          if (detalhes) assert.equal(url, correta);
          const documento = { title: caso.nome + " | LigaMagic",
            scripts: [{ textContent: detalhes ? "var cards_editions = " + JSON.stringify([paginaPrecos]) + "; var dataExtras = "
              + JSON.stringify([{ id: 2, label: "Foil" }]) + ";" : "var cardsjson = " + JSON.stringify(cartas) + ";" }],
            querySelector: () => detalhes ? { content: caso.nome + " | Busca de Cartas" } : null,
            querySelectorAll: () => detalhes ? [] : [correta, errada].map(href => ({ getAttribute: () => href })) };
          return [{ result: executar(documento, { href: url })(...args) }];
        } },
        alarms: { create: async () => {}, clear: async () => {} },
      };
      const controle = criarControlador(api, () => tempo, undefined, async () => ({ fullArt: false, tratamento: "padrao",
        referenciasTerreno: [{ codigoEdicao: "fdn", numero: caso.numero, fullArt: false, tratamento: "padrao" }] }));
      await controle.comando("iniciar", { colecao: { criterios: "variantes", variantes: [item] } });
      await controle.processar(7);
      assert.equal(estado.lote.itens[0].proximaBusca, correta);
      tempo = 4000;
      await controle.processar();
      assert.equal(url, correta);
      await controle.processar(7);
      assert.equal(estado.lote.fase, "concluido");
      assert.equal(estado.lote.itens[0].status, "ok");
      assert.equal(estado.lote.itens[0].cotacao.codigoEdicao, prefixo + "fdn");
      assert.equal(estado.lote.itens[0].cotacao.numero, caso.numero);
      assert.equal(estado.lote.itens[0].cotacao.centavos, acabamento === "normal" ? Math.round(Number(caso.media) * 100) : 99);
      assert.equal(estado.lote.itens[0].cotacao.url, correta);
      assert.equal(navegacoes.length, 2);
      assert.equal(navegacoes.includes(errada), false);
    }
  }
});

test("lote de Swamp abre impressao numerada recente, retoma e preserva referencia original", async () => {
  const estado = {};
  let tempo = 0;
  let url;
  let resposta;
  let consultas = 0;
  const item = { nome: "Swamp", quantidade: 6, codigoEdicao: "eld", edicao: "Throne of Eldraine", numero: "261", acabamento: "normal", idioma: "en" };
  const imprimir = (code, num, m) => ({ nome: "Swamp (#" + num + ")", code, num, name: code, price: { 0: { m } },
    url: "https://www.ligamagic.com.br/?" + new URLSearchParams({ view: "cards/card", card: "Swamp (#" + num + ")", ed: "100", num }) });
  const antiga = imprimir("eld", "261", "0.60");
  const recente = imprimir("new", "400", "0.20");
  const api = {
    storage: { local: { get: async (chave) => structuredClone({ [chave]: estado[chave] }), set: async (valores) => Object.assign(estado, structuredClone(valores)) } },
    permissions: { contains: async () => true },
    tabs: { create: async () => ({ id: 7 }), get: async () => ({ status: "complete", url }), update: async (id, dados) => { url = dados.url; } },
    scripting: { executeScript: async () => [{ result: resposta }] }, alarms: { create: async () => {}, clear: async () => {} },
  };
  const resolver = async () => { consultas++; return { fullArt: false, referenciasTerreno: [
    { codigoEdicao: "eld", numero: "261", fullArt: false, dataLancamento: "2019-10-04" },
    { codigoEdicao: "new", numero: "400", fullArt: false, dataLancamento: "2026-01-01" },
  ] }; };
  const controle = criarControlador(api, () => tempo, undefined, resolver);
  await controle.comando("iniciar", { texto: "6 Swamp", intervalo: 3.5, colecao: { criterios: "variantes", variantes: [item] } });
  resposta = { resultados: [{ url }], edicoesResultados: [antiga, recente] };
  await controle.processar(7);
  assert.equal(estado.lote.fase, "aguardando");
  assert.equal(estado.lote.itens[0].proximaBusca, recente.url);
  tempo = 3499;
  const retomado = criarControlador(api, () => tempo, undefined, resolver);
  await retomado.processar();
  assert.notEqual(url, recente.url);
  tempo = 3500;
  await retomado.processar();
  assert.equal(url, recente.url);
  assert.equal(estado.lote.fase, "navegando");
  resposta = { nome: "Swamp (#400)", url, edicoes: [recente] };
  await retomado.processar(7);
  assert.equal(estado.lote.fase, "concluido");
  assert.equal(estado.lote.itens[0].numero, "261");
  assert.equal(estado.lote.itens[0].cotacao.numero, "400");
  assert.equal(estado.lote.itens[0].cotacao.centavos, 20);
  assert.equal(consultas, 1);
  const { gerarCsvColecao } = await import("./core.mjs");
  const relatorio = gerarCsvColecao(estado.lote);
  assert.ok(relatorio.includes("Edicao usada;Sigla usada;Numero usado;Referencia substituida;Motivo da referencia;URL da referencia"));
  assert.ok(relatorio.includes(";new;400;Sim;Terreno basico comum"));
  assert.equal(nomesEquivalentes("Swamp (#400)", "Swamp"), true);
  assert.equal(nomesEquivalentes("Other Card (#400)", "Other Card"), false);
});

test("lista comum de terrenos usa URL recente e preco verde sem confundir media foil ou full art", async () => {
  const estado = {};
  let tempo = 0;
  let url;
  let pagina;
  const referencia = (code, num, p, m) => ({ code, num, name: code, price: { 0: { p, m } },
    url: "https://www.ligamagic.com.br/?" + new URLSearchParams({ view: "cards/card", card: "Swamp (#" + num + ")", ed: "100", num }) });
  const antiga = referencia("eld", "261", "0.01", "0.50");
  const recente = referencia("new", "400", "0.20", "9.99");
  const arte = referencia("art", "500", "0.01", "1.00");
  const api = {
    storage: { local: { get: async (chave) => structuredClone({ [chave]: estado[chave] }), set: async (dados) => Object.assign(estado, structuredClone(dados)) } },
    permissions: { contains: async () => true },
    tabs: { create: async () => ({ id: 7 }), get: async () => ({ status: "complete", url }), update: async (id, dados) => { url = dados.url; } },
    scripting: { executeScript: async ({ args }) => { assert.deepEqual(args, [true]); return [{ result: pagina }]; } },
    alarms: { create: async () => {}, clear: async () => {} },
  };
  const resolver = async (item) => { assert.equal(item.nome, "Swamp"); return { fullArt: false, referenciasTerreno: [
    { codigoEdicao: "eld", numero: "261", fullArt: false, dataLancamento: "2019-10-04" },
    { codigoEdicao: "new", numero: "400", fullArt: false, dataLancamento: "2026-01-01" },
    { codigoEdicao: "art", numero: "500", fullArt: true, dataLancamento: "2026-02-01" },
  ] }; };
  let enriquecimentos = 0;
  const enriquecer = async (cotacao) => { enriquecimentos++; return { ...cotacao, statusTagger: "ok", tiposCarta: ["Land"] }; };
  const controle = criarControlador(api, () => tempo, enriquecer, resolver);
  await controle.comando("iniciar", { texto: "6 Swamp", intervalo: 3.5, somenteNormal: false });
  pagina = { resultados: [{ url }], edicoesResultados: [antiga, arte, recente] };
  await controle.processar(7);
  assert.equal(estado.lote.fase, "aguardando");
  assert.equal(estado.lote.proxima, 3500);
  tempo = 3500;
  await controle.processar();
  assert.equal(url, recente.url);
  pagina = { nome: "Swamp", url, edicoes: [{ ...recente, price: { ...recente.price, 2: { p: "0.01", m: "0.50" } } }], extras: [{ id: 2, label: "Foil" }] };
  await criarControlador(api, () => tempo, enriquecer, resolver).processar(7);
  assert.equal(estado.lote.fase, "concluido");
  assert.equal(estado.capturas[0].centavos, 20);
  assert.equal(estado.capturas[0].quantidade, 6);
  assert.equal(estado.capturas[0].variante, "Normal");
  assert.equal(estado.capturas[0].codigoEdicao, "new");
  assert.equal(estado.capturas[0].url, recente.url);
  assert.equal(enriquecimentos, 1);
});

test("limpar colecao remove lote rascunhos e capturas sem apagar o cache", async () => {
  for (const fase of ["pausado", "concluido", "cancelado", "navegando"]) {
    const estado = { lote: { fase, colecao: { criterios: "variantes" }, itens: [], indice: 0 },
      rascunhoLista: "2 Sol Ring", rascunhoColecao: { arquivo: "cartas.csv" },
      capturas: [{ nome: "Captura avulsa", centavos: 100, quantidade: 1 }] };
    const original = structuredClone(estado);
    const cache = criarCachePersistente({ banco: null });
    await cache.set("carta-guardada", { centavos: 100 }, VALIDADE.precos);
    const alarmes = [];
    const api = { storage: { local: { get: async (chave) => structuredClone({ [chave]: estado[chave] }),
      set: async (dados) => Object.assign(estado, structuredClone(dados)) } },
      alarms: { clear: async (nome) => alarmes.push(nome) } };
    const controle = criarControlador(api, Date.now, undefined, undefined, { cache });
    if (fase === "navegando") {
      await assert.rejects(controle.comando("limpar"), /Pause/);
      assert.deepEqual(estado, original);
      assert.equal(alarmes.length, 0);
    } else {
      await controle.comando("limpar");
      assert.equal(estado.lote, null);
      assert.equal(estado.rascunhoLista, "");
      assert.equal(estado.rascunhoColecao, null);
      assert.deepEqual(estado.capturas, []);
      assert.equal(alarmes.length, 1);
    }
    assert.deepEqual(await cache.get("carta-guardada"), { centavos: 100 });
  }
});

test("checkbox de colecao e intervalo escolhido sao restaurados e enviados pelo popup", async () => {
  const { importarCsvColecao } = await import("./core.mjs");
  const documentoAnterior = globalThis.document;
  const chromeAnterior = globalThis.chrome;
  const janelaAnterior = globalThis.window;
  const intervalos = new Set();
  let aoFechar;
  let bytesSalvos = 0;
  const dados = importarCsvColecao("Name,Quantity,Set code,Set name,Collector number,Foil,Language\nSol Ring,2,CMM,Commander Masters,410,normal,en");
  const estado = { rascunhoLista: dados.texto, somenteNormal: true,
    rascunhoColecao: { arquivo: "cartas.csv", texto: dados.texto, variantes: dados.variantes, proxiesExcluidos: 0 } };
  const comandos = [];
  const criarElemento = () => ({ checked: false, value: "", dataset: {}, eventos: new Map(),
    addEventListener(tipo, funcao) { this.eventos.set(tipo, funcao); },
    checkValidity() { return true; },
    showModal() { this.open = true; }, close() { this.open = false; },
    setAttribute() {}, append() {}, prepend() {}, replaceChildren() {},
    emitir(tipo) { return this.eventos.get(tipo)({ currentTarget: this }); },
  });
  try {
    globalThis.window = {
      setInterval(acao, atraso) {
        assert.equal(atraso, 5000);
        const intervalo = { acao };
        intervalos.add(intervalo);
        return intervalo;
      },
      clearInterval(intervalo) { intervalos.delete(intervalo); },
      addEventListener(tipo, acao) { if (tipo === "pagehide") aoFechar = acao; },
    };
    globalThis.chrome = { storage: { local: {
      QUOTA_BYTES: 1000,
      getBytesInUse: async () => bytesSalvos,
      get: async (padrao) => structuredClone({ ...padrao, ...estado }),
      set: async (dados) => Object.assign(estado, structuredClone(dados)),
    }, onChanged: { addListener() {} } }, permissions: { request: async () => true }, runtime: {
      sendMessage: async (comando) => {
        comandos.push(comando);
        if (comando.comando === "iniciar") estado.lote = { fase: "aguardando", indice: 0, itens: [{ nome: "Sol Ring", quantidade: 2, status: "pendente" }],
          intervalo: comando.dados.intervalo, texto: comando.dados.texto, comandantes: comando.dados.comandante ? [comando.dados.comandante] : [] };
        else if (comando.comando === "limpar") Object.assign(estado, { lote: null, capturas: [], rascunhoLista: "", rascunhoColecao: null, rascunhoComandante: "" });
        else estado.lote.fase = "aguardando";
        return { ok: true };
      },
    } };
    for (const etapa of ["ativar", "reabrir", "desativar", "reabrir-desativado", "intervalo-salvo", "pendencias", "comandante", "reabrir-comandante", "limpar-colecao"]) {
      if (etapa === "comandante") { estado.lote = null; estado.consultaColecao = false; }
      if (etapa === "intervalo-salvo") estado.lote = { fase: "concluido", itens: [], indice: 0, intervalo: 7.5, texto: dados.texto };
      if (etapa === "pendencias") estado.lote = { fase: "concluido", itens: [{ nome: "Sol Ring", quantidade: 2, status: "erro", erro: "Sem preco" }],
        indice: 1, intervalo: 8.5, texto: dados.texto };
      if (etapa === "limpar-colecao") {
        estado.consultaColecao = true;
        estado.lote = { fase: "pausado", colecao: { criterios: "variantes", arquivo: "cartas.csv" },
          itens: [{ ...dados.variantes[0], status: "pendente" }], indice: 0, intervalo: 4, texto: dados.texto };
      }
      const elementos = new Map();
      globalThis.document = { body: { dataset: {} }, createElement: criarElemento,
        getElementById(id) { if (!elementos.has(id)) elementos.set(id, criarElemento()); return elementos.get(id); } };
      await import("./popup.mjs?teste-checkbox=" + etapa);
      await new Promise((resolver) => setImmediate(resolver));
      assert.equal(elementos.get("memory-warning").hidden, true);
      if (etapa === "ativar") {
        const statusAnterior = elementos.get("status").textContent;
        const intervalo = [...intervalos][0];
        for (const [bytes, critico] of [[850, false], [950, true], [0, false]]) {
          bytesSalvos = bytes;
          await intervalo.acao();
          const aviso = elementos.get("memory-warning");
          assert.equal(aviso.hidden, bytes === 0);
          assert.equal(aviso.dataset.critical, String(critico));
          if (bytes) assert.match(aviso.textContent, /Armazenamento das capturas quase cheio/);
          assert.equal(elementos.get("status").textContent, statusAnterior);
        }
      }
      aoFechar();
      assert.equal(intervalos.size, 0);
      await elementos.get("tab-batch").emitir("click");
      const checkbox = elementos.get("collection-mode");
      if (etapa === "comandante") {
        const campo = elementos.get("batch-commander");
        assert.equal(campo.disabled, false);
        campo.value = "Grist, Voracious Larva";
        await campo.emitir("change");
        assert.equal(estado.rascunhoComandante, campo.value);
        await elementos.get("batch-start").emitir("click");
        assert.equal(comandos.at(-1).dados.comandante, campo.value);
        assert.equal(campo.disabled, true);
        continue;
      }
      if (etapa === "reabrir-comandante") {
        assert.equal(elementos.get("batch-commander").value, "Grist, Voracious Larva");
        assert.equal(elementos.get("batch-commander").disabled, true);
        continue;
      }
      if (etapa === "limpar-colecao") {
        assert.equal(elementos.get("clear").disabled, false);
        assert.equal(elementos.get("confirm-clear").disabled, false);
        await elementos.get("clear").emitir("click");
        assert.equal(elementos.get("clear-title").textContent, "Limpar a lista atual?");
        assert.equal(elementos.get("clear-dialog").open, true);
        await elementos.get("confirm-clear").emitir("click");
        assert.equal(comandos.at(-1).comando, "limpar");
        assert.equal(estado.lote, null);
        assert.equal(elementos.get("batch-text").value, "");
        assert.equal(elementos.get("batch-commander").value, "");
        assert.equal(estado.rascunhoComandante, "");
        assert.equal(elementos.get("clear-dialog").open, false);
        assert.equal(elementos.get("clear").disabled, true);
        assert.equal(elementos.get("import-list").disabled, false);
        assert.equal(elementos.get("batch-start").disabled, true);
        continue;
      }
      assert.equal(elementos.get("normal").checked, Boolean(estado.somenteNormal) && !estado.consultaColecao);
      assert.equal(checkbox.checked, Boolean(estado.consultaColecao), etapa);
      assert.equal(checkbox.disabled, false);
      assert.equal(Number(elementos.get("batch-interval").value), estado.lote?.intervalo ?? 4);
      assert.equal(elementos.get("batch-interval").disabled, false);
      if (etapa === "ativar" || etapa === "desativar") {
        const esperado = etapa === "ativar";
        checkbox.checked = esperado;
        await checkbox.emitir("change");
        assert.equal(checkbox.checked, esperado, etapa);
        assert.equal(estado.consultaColecao, esperado, etapa);
        assert.equal(elementos.get("normal").checked, false);
        assert.equal(estado.somenteNormal, false);
        assert.equal(elementos.get("normal").disabled, esperado);
        assert.equal(elementos.get("batch-text").readOnly, esperado);
        assert.equal(elementos.get("batch-start").textContent, esperado ? "Estimar colecao" : "Iniciar lista");
        assert.equal(elementos.get("batch-start").disabled, false);
        assert.equal(elementos.get("batch-text").value, dados.texto);
        assert.match(elementos.get("status").textContent, esperado ? /Modo colecao ativado/ : /Modo colecao desativado/);
      }
      if (etapa === "intervalo-salvo") {
        elementos.get("batch-interval").value = "8.5";
        await elementos.get("batch-start").emitir("click");
        assert.equal(comandos.length, 1);
        assert.equal(comandos[0].comando, "iniciar");
        assert.equal(comandos[0].dados.intervalo, 8.5);
        assert.equal(elementos.get("batch-interval").disabled, true);
      }
      if (etapa === "pendencias") {
        assert.equal(elementos.get("batch-retry").hidden, false);
        await elementos.get("batch-retry").emitir("click");
        assert.equal(comandos.at(-1).comando, "retentar");
        assert.equal(elementos.get("batch-retry").hidden, true);
      }
    }
  } finally {
    if (documentoAnterior === undefined) delete globalThis.document;
    else globalThis.document = documentoAnterior;
    if (chromeAnterior === undefined) delete globalThis.chrome;
    else globalThis.chrome = chromeAnterior;
    if (janelaAnterior === undefined) delete globalThis.window;
    else globalThis.window = janelaAnterior;
  }
});

test("erros de rede identificam a API e preservam precos sem consultar o Tagger", async () => {
  for (const erro of [new TypeError("Failed to fetch"), new DOMException("timeout", "TimeoutError")]) {
    let consultouTags = false;
    const dados = await enriquecerCotacao({ nome: "Sol Ring", centavos: 123 }, async () => { throw erro; }, async () => { consultouTags = true; });
    assert.equal(dados.centavos, 123);
    assert.equal(dados.statusTagger, "erro");
    assert.match(dados.erroTagger, /api\.scryfall\.com/);
    assert.equal(consultouTags, false);
    if (erro.name === "TimeoutError") assert.match(dados.erroTagger, /demorou/);
    else assert.match(dados.erroTagger, /certificado HTTPS/);
  }
});

test("funcao injetada serializa falhas do Tagger sem perder o diagnostico", async () => {
  const executar = new Function("document", "fetch", "return (" + lerTagsNaPagina.toString() + ")");
  const documento = { querySelector: () => ({ content: "csrf-teste" }) };
  const casos = [
    { requisitar: async () => { throw new TypeError("Failed to fetch"); }, esperado: /tagger\.scryfall\.com.*certificado HTTPS/ },
    { requisitar: async () => { throw new DOMException("timeout", "TimeoutError"); }, esperado: /tagger\.scryfall\.com demorou/ },
    { requisitar: async () => ({ ok: false, status: 429 }), esperado: /HTTP 429/ },
    { requisitar: async () => ({ ok: true, json: async () => { throw new SyntaxError("JSON invalido"); } }), esperado: /JSON invalido/ },
  ];
  for (const caso of casos) {
    const resultado = await executar(documento, caso.requisitar)("query Teste", {});
    assert.match(resultado.erroTagger, caso.esperado);
  }
  const semSessao = await executar({ querySelector: () => null }, async () => { throw new Error("Nao deveria consultar"); })("query Teste", {});
  assert.match(semSessao.erroTagger, /Sessao.*indisponivel/);
  const sucesso = await executar(documento, async () => ({ ok: true, json: async () => ({ data: { front: {} } }) }))("query Teste", {});
  assert.deepEqual(sucesso, { data: { front: {} } });
});

test("consulta GraphQL dentro de aba auxiliar e sempre fecha a aba criada", async () => {
  for (const falhar of [false, true, "rede"]) {
    const ouvintes = new Set();
    const removidas = [];
    const evento = { addListener: (funcao) => ouvintes.add(funcao), removeListener: (funcao) => ouvintes.delete(funcao) };
    const api = { tabs: { onUpdated: evento, onRemoved: evento, create: async ({ active }) => {
      assert.equal(active, false);
      return { id: 8 };
    }, get: async () => ({ status: "complete" }), remove: async (id) => removidas.push(id) },
    scripting: { executeScript: async ({ target, args, world }) => {
      assert.equal(target.tabId, 8);
      assert.equal(world, "MAIN");
      assert.ok(args[0].startsWith("query "));
      assert.match(args[0], /status taggingCount ancestorTags/);
      if (falhar === "rede") return [{ result: { erroTagger: "Falha de conexao com tagger.scryfall.com." } }];
      if (falhar) throw new Error("Bloqueado");
      return [{ result: { data: { front: {} } } }];
    } } };
    if (falhar) await assert.rejects(consultarTagsNavegador("https://tagger.scryfall.com/card/cmm/410", {}, api), falhar === "rede" ? /tagger\.scryfall\.com/ : /Bloqueado/);
    else assert.ok((await consultarTagsNavegador("https://tagger.scryfall.com/card/cmm/410", {}, api)).data);
    assert.deepEqual(removidas, [8]);
    assert.equal(ouvintes.size, 0);
  }
});

test("CSV preserva quatro colunas e exporta apenas categorias Tagger filtradas", () => {
  const cotacao = { nome: "Sol Ring", centavos: 123, quantidade: 2, tiposCarta: ["Artifact"],
    categoriasTagger: ["Adds Mana", "Mana Rock"], tagsTagger: ["activated-ability", "mana-rock"],
    categoriaArchidekt: "Antiga", categoriasOracle: ["ignorar"] };
  for (const csv of [gerarCsv([cotacao]), gerarCsvLote({ itens: [{ nome: cotacao.nome, status: "ok", cotacao }] })]) {
    assert.equal(csv, "\uFEFFNome da carta;Pre\u00e7o m\u00ednimo;Tipos;Categorias (principal)\r\n"
      + "Sol Ring;1,23;Artifact;Adds Mana | Mana Rock\r\n");
    for (const linha of csv.trimEnd().split("\r\n")) assert.equal(linha.split(";").length, 4);
  }
  assert.ok(gerarCsv([{ ...cotacao, categoriasTagger: [] }]).endsWith("Sol Ring;1,23;Artifact;Land\r\n"));
  assert.ok(gerarCsvLote({ itens: [{ nome: "Ausente", status: "erro" }] }).endsWith("Ausente;;;Land\r\n"));
  assert.ok(gerarCsv([{ ...cotacao, nome: '=nome; "carta"' }]).includes('"\'=nome; ""carta"""'));
  assert.throws(() => gerarCsv([{ ...cotacao, centavos: -1 }]), /Preco invalido/);
});

test("controlador salva categorias do Tagger e preserva preco se faltar rede ou permissao", async () => {
  for (const cenario of ["ok", "offline", "sem permissao"]) {
    const pagina = { nome: "Sol Ring", nomePt: "Anel Solar", url: "https://www.ligamagic.com.br/?view=cards/card&card=Sol+Ring",
      capturadoEm: "2026-10-27", edicoes: [{ name: "Commander", price: { 0: { p: "1.23" } } }] };
    const estado = { capturas: [], lote: { fase: "navegando", abaId: 7, indice: 0, prazo: 45000, intervalo: 3,
      itens: [{ nome: "Sol Ring", quantidade: 3, status: "pendente", erro: "" }] } };
    const api = {
      storage: { local: { get: async (chave) => structuredClone({ [chave]: estado[chave] }),
        set: async (dados) => Object.assign(estado, structuredClone(dados)) } },
      permissions: { contains: async ({ origins }) => !origins.includes("https://tagger.scryfall.com/*") || cenario !== "sem permissao" },
      tabs: { get: async () => ({ status: "complete", url: pagina.url }) },
      scripting: { executeScript: async () => [{ result: pagina }] },
      alarms: { create: async () => {}, clear: async () => {} },
    };
    let chamadas = 0;
    const enriquecer = async (preco) => {
      chamadas++;
      return cenario === "ok" ? { ...preco, tiposCarta: ["Artifact"], categoriasTagger: ["Adds Mana", "Mana Rock"], statusTagger: "ok" }
        : enriquecerCotacao(preco, async () => { throw new Error("Offline"); });
    };
    await criarControlador(api, () => 0, enriquecer).processar(7);
    assert.equal(estado.lote.fase, "concluido");
    assert.equal(estado.lote.itens[0].status, "ok");
    assert.equal(estado.capturas[0].quantidade, 3);
    assert.equal(estado.capturas[0].centavos, 123);
    assert.equal(chamadas, cenario === "sem permissao" ? 0 : 1);
    assert.equal(estado.capturas[0].statusTagger, cenario === "ok" ? "ok" : "erro");
    if (cenario === "ok") assert.ok(gerarCsvLote(estado.lote).endsWith("Sol Ring;1,23;Artifact;Adds Mana | Mana Rock\r\n"));
  }
});