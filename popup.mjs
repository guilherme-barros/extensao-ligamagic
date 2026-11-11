import { capturarPagina, extrairCotacao, adicionarCaptura, totalCentavos, reais, interpretarLista, prepararListaDeck, foraDoTotalDeck, importarCsvColecao, prepararColecao, resumirColecao } from "./core.mjs";
import { gerarExcel, gerarExcelLote, gerarExcelColecao } from "./excel.mjs";
import { ORIGENS_IMAGENS } from "./imagens.mjs";
import { ORIGENS_CONSULTA, loteAberto } from "./lote.mjs";
import { enriquecerCotacao, ORIGENS_TAGGER } from "./tagger.mjs";
import { verificarMemoria } from "./memoria.mjs";

const elementos = Object.fromEntries(["capture", "normal", "status", "memory-warning", "count", "items", "empty", "total", "total-label", "units", "export", "export-label", "clear", "clear-dialog", "cancel-clear", "confirm-clear"]
  .map((id) => [id, document.getElementById(id)]));
const lista = Object.fromEntries(["tab-manual", "tab-batch", "manual-panel", "batch-panel", "batch-text", "batch-commander", "commander-field", "import-list", "list-file", "collection-mode", "collection-summary", "collection-basis", "batch-interval", "batch-refresh", "batch-start", "batch-pause", "batch-resume", "batch-retry", "batch-skip", "batch-tab", "batch-cancel", "batch-progress", "batch-meter", "batch-state", "batch-total", "batch-export", "batch-errors"]
  .map((id) => [id, document.getElementById(id)]));
let capturas = [];
let lote = null;
let colecao = null;
let consultaColecao = false;
let ocupado = true;
let pronto = false;
let verificandoMemoria = false;

function mensagem(texto, tipo = "") {
  elementos.status.textContent = texto;
  elementos.status.dataset.kind = tipo;
}

async function atualizarAvisoMemoria() {
  if (verificandoMemoria) return;
  verificandoMemoria = true;
  try {
    const { avisos, critico } = await verificarMemoria();
    const texto = avisos.length ? (critico ? "Memoria em nivel critico.\n" : "Aviso de memoria.\n") + avisos.join("\n") : "";
    const aviso = elementos["memory-warning"];
    if (aviso.textContent !== texto) aviso.textContent = texto;
    aviso.dataset.critical = String(critico);
    aviso.hidden = !avisos.length;
  } finally {
    verificandoMemoria = false;
  }
}

function icone(nome) {
  const imagem = document.createElement("img");
  imagem.src = "assets/" + nome + ".svg";
  imagem.alt = "";
  return imagem;
}

function renderizar() {
  const avaliacao = document.body.dataset.mode === "batch" && lote?.colecao ? resumirColecao(lote) : null;
  const visiveis = avaliacao ? lote.itens.filter((item) => item.status === "ok" && item.cotacao)
    .map((item) => ({ ...item.cotacao, quantidade: item.quantidade })) : capturas;
  elementos.items.replaceChildren();
  for (const item of visiveis) {
    const linha = document.createElement("div");
    linha.className = "item";
    const detalhes = document.createElement("div");
    detalhes.className = "card-details";
    const nome = document.createElement("a");
    const url = new URL(item.url);
    if (url.protocol === "https:" && ["ligamagic.com.br", "www.ligamagic.com.br"].includes(url.hostname)) {
      nome.href = url.href;
      nome.target = "_blank";
      nome.rel = "noopener noreferrer";
    }
    nome.textContent = item.nome;
    const descricao = document.createElement("small");
    descricao.textContent = item.edicao + " / " + item.variante + (item.idioma ? " / " + item.idioma.toUpperCase() : "");
    if (item.codigoEdicao && item.numero) descricao.textContent += " / " + item.codigoEdicao.toUpperCase() + " #" + item.numero;
    const data = document.createElement("small");
    const instante = new Date(item.capturadoEm).toLocaleString("pt-BR", { dateStyle: "short", timeStyle: "short" });
    data.textContent = instante + (item.tipoPreco === "media_ligamagic" ? " / Media LigaMagic (amarelo), sem filtro de idioma"
      : item.tipoPreco === "media_ofertas" ? ` / Media de ${item.ofertasConsideradas} ofertas`
      : item.somenteNormal ? " / Normal" : " / Todas as variantes");
    detalhes.append(nome, descricao, data);
    if (item.referenciaSubstituida) {
      const referencia = document.createElement("small");
      referencia.textContent = "Referencia substituta: " + item.motivoReferencia;
      detalhes.append(referencia);
    }
    if (item.categoriasTagger?.length) {
      const categorias = document.createElement("small");
      categorias.textContent = item.categoriasTagger.join(" / ");
      detalhes.append(categorias);
    }
    if (item.statusTagger === "erro") {
      const aviso = document.createElement("small");
      aviso.textContent = "Scryfall Tagger: " + item.erroTagger;
      detalhes.append(aviso);
    }
    const quantidade = document.createElement("input");
    quantidade.className = "quantity";
    quantidade.type = "number";
    quantidade.min = "1";
    quantidade.max = "9999";
    quantidade.step = "1";
    quantidade.value = item.quantidade;
    quantidade.disabled = ocupado || loteAberto(lote) || Boolean(avaliacao);
    quantidade.setAttribute("aria-label", "Quantidade de " + item.nome);
    quantidade.addEventListener("change", () => executar(async () => {
      if (!quantidade.checkValidity()) throw new Error("Informe uma quantidade inteira entre 1 e 9999.");
      await salvar(capturas.map((atual) => atual.id === item.id ? { ...atual, quantidade: Number(quantidade.value) } : atual));
      mensagem("Quantidade atualizada.", "success");
    }));
    const preco = document.createElement("div");
    preco.className = "price";
    const subtotal = document.createElement("strong");
    subtotal.textContent = reais(item.centavos * item.quantidade);
    const unitario = document.createElement("small");
    unitario.textContent = reais(item.centavos) + " / un.";
    preco.append(subtotal, unitario);
    const remover = document.createElement("button");
    remover.className = "icon-button remove danger";
    remover.title = "Remover " + item.nome;
    remover.setAttribute("aria-label", remover.title);
    remover.disabled = ocupado || loteAberto(lote) || Boolean(avaliacao);
    remover.append(icone("x"));
    remover.addEventListener("click", () => executar(async () => {
      await salvar(capturas.filter((atual) => atual.id !== item.id));
      mensagem(item.nome + " removida.");
    }));
    linha.append(detalhes, quantidade, preco, remover);
    elementos.items.append(linha);
  }
  elementos.empty.hidden = visiveis.length > 0;
  const distintas = avaliacao ? avaliacao.distintas : capturas.length;
  elementos.count.textContent = distintas + (avaliacao && lote.colecao.criterios === "variantes" ? (distintas === 1 ? " variante" : " variantes") : distintas === 1 ? " carta" : " cartas");
  const cartasTotal = capturas.filter((item) => !foraDoTotalDeck(item, lote?.comandantes));
  elementos.total.textContent = reais(avaliacao ? avaliacao.centavos : totalCentavos(cartasTotal));
  elementos["total-label"].textContent = avaliacao ? (avaliacao.pendentes ? "Estimativa parcial da colecao" : "Estimativa da colecao")
    : cartasTotal.length < capturas.length ? "Total Commander 500" : "Total capturado";
  const unidades = capturas.reduce((soma, item) => soma + item.quantidade, 0);
  elementos.units.textContent = avaliacao ? `${avaliacao.avaliadas}/${avaliacao.unidades} unidades avaliadas / sem frete`
    : unidades + (unidades === 1 ? " unidade" : " unidades") + " / sem frete";
  elementos.capture.disabled = ocupado || !pronto || loteAberto(lote);
  elementos.normal.disabled = ocupado || !pronto || loteAberto(lote) || (document.body.dataset.mode === "batch" && consultaColecao);
  const exportaLote = document.body.dataset.mode === "batch" && lote;
  elementos.export.disabled = ocupado || (!exportaLote && !capturas.length);
  elementos["export-label"].textContent = avaliacao ? "Exportar estimativa Excel" : exportaLote ? "Exportar lista Excel" : "Exportar Excel";
  const limparLote = document.body.dataset.mode === "batch" && Boolean(lote || colecao || lista["batch-text"].value || lista["batch-commander"].value);
  const bloqueado = loteAberto(lote) && (!limparLote || lote.fase !== "pausado");
  elementos.clear.disabled = ocupado || bloqueado || (!limparLote && !capturas.length);
  elementos["confirm-clear"].disabled = ocupado || bloqueado;
  renderizarLote();
}

function renderizarLote() {
  const aberto = loteAberto(lote);
  const pausado = lote?.fase === "pausado";
  lista["collection-mode"].checked = consultaColecao;
  lista["collection-mode"].disabled = ocupado || aberto || !pronto;
  lista["collection-mode"].title = aberto ? "Conclua ou cancele a consulta atual antes de alterar o modo." : "Ativar preco medio amarelo por edicao e acabamento do CSV";
  lista["batch-text"].readOnly = consultaColecao;
  lista["commander-field"].hidden = consultaColecao;
  lista["batch-commander"].disabled = ocupado || aberto || !pronto || consultaColecao;
  for (const id of ["batch-text", "batch-interval", "batch-refresh", "import-list"]) lista[id].disabled = ocupado || aberto || !pronto;
  lista["batch-start"].hidden = aberto;
  lista["batch-start"].disabled = ocupado || !pronto || (consultaColecao && !colecao?.variantes?.length);
  lista["batch-pause"].hidden = !aberto || pausado;
  lista["batch-resume"].hidden = !pausado;
  lista["batch-retry"].hidden = !lote || !["concluido", "cancelado", "pausado"].includes(lote.fase) || !lote.itens.some((item) => item.status !== "ok");
  lista["batch-skip"].hidden = !pausado;
  lista["batch-cancel"].hidden = !aberto;
  lista["batch-tab"].hidden = !lote?.abaId;
  for (const id of ["batch-pause", "batch-resume", "batch-retry", "batch-skip", "batch-cancel", "batch-tab"]) lista[id].disabled = ocupado;
  lista["batch-progress"].hidden = !lote;
  const origem = aberto ? lote.colecao : colecao;
  lista["collection-summary"].hidden = !origem && !consultaColecao;
  lista["collection-basis"].hidden = !consultaColecao && !lote?.colecao;
  lista["collection-basis"].textContent = consultaColecao || lote?.colecao?.criterios === "variantes"
    ? "Preco medio amarelo. Terreno basico comum: versao recente; Foil/Full Art: versao exata ou alternativa compativel identificada. Siglas relacionadas incluem BB. Sem filtro de idioma, conservacao ou frete."
    : "Base: menor preco por nome, sem ajuste por edicao, acabamento, idioma ou estado. Sem frete; nao representa oferta de recompra.";
  lista["batch-start"].textContent = consultaColecao ? "Estimar colecao" : "Iniciar lista";
  lista["batch-start"].prepend(icone("play"));
  if (!origem && consultaColecao) lista["collection-summary"].textContent = "Nenhuma colecao CSV importada.";
  if (origem) {
    try {
      const resumo = resumirColecao(aberto ? lote : { itens: consultaColecao ? prepararColecao(origem.variantes) : interpretarLista(lista["batch-text"].value, true) });
      lista["collection-summary"].textContent = `${origem.arquivo} / ${resumo.distintas} ${consultaColecao ? "variantes" : "nomes"} / ${resumo.unidades} unidades`
        + (origem.proxiesExcluidos ? ` / ${origem.proxiesExcluidos} proxies excluidos` : "");
    } catch {
      lista["collection-summary"].textContent = origem.arquivo + (consultaColecao ? " / Reimporte o CSV completo com edicao, acabamento e idioma." : " / Quantidades invalidas.");
    }
  }
  if (!lote) return;
  lista["batch-meter"].max = lote.itens.length;
  lista["batch-meter"].value = lote.indice;
  const fases = { aguardando: "Aguardando", navegando: "Abrindo", lendo: "Lendo", pausado: "Pausada", concluido: "Concluida", cancelado: "Cancelada" };
  const atual = lote.itens[lote.indice]?.nome;
  lista["batch-state"].textContent = `${lote.indice}/${lote.itens.length} / ${fases[lote.fase]}${atual ? " / " + atual : ""}. ${lote.mensagem}`;
  if (lote.metricas) lista["batch-state"].textContent += " / " + (lote.metricas.cartasCache || 0) + " em cache / " + (lote.metricas.navegacoes || 0) + " acessos";
  lista["batch-state"].dataset.paused = String(pausado);
  const validas = lote.itens.filter((item) => item.status === "ok").map((item) => ({ ...item.cotacao, quantidade: item.quantidade }));
  const parcial = validas.length !== lote.itens.length;
  lista["batch-total"].textContent = (lote.colecao ? (parcial ? "Estimativa parcial: " : "Estimativa da colecao: ")
    : parcial ? "Total parcial Commander 500: " : "Total Commander 500: ")
    + reais(totalCentavos(lote.colecao ? validas : validas.filter((item) => !foraDoTotalDeck(item, lote.comandantes))));
  lista["batch-errors"].replaceChildren();
  for (const item of lote.itens.filter((item) => item.erro)) {
    const paragrafo = document.createElement("p");
    paragrafo.textContent = item.nome + (item.idColecao ? ` / ${item.codigoEdicao.toUpperCase()} #${item.numero} / ${item.acabamento} / ${item.idioma.toUpperCase()}` : "") + ": " + item.erro;
    lista["batch-errors"].append(paragrafo);
  }
  for (const item of lote.itens.filter((item) => item.cotacao?.statusTagger === "erro")) {
    const paragrafo = document.createElement("p");
    paragrafo.textContent = item.nome + " / Scryfall Tagger: " + item.cotacao.erroTagger;
    lista["batch-errors"].append(paragrafo);
  }
}

function alternarModo(modo) {
  document.body.dataset.mode = modo;
  for (const [id, escolhido] of [["tab-manual", modo === "manual"], ["tab-batch", modo === "batch"]]) {
    lista[id].setAttribute("aria-selected", String(escolhido));
    lista[id].tabIndex = escolhido ? 0 : -1;
  }
  lista["manual-panel"].hidden = modo !== "manual";
  lista["batch-panel"].hidden = modo !== "batch";
  renderizar();
}

for (const [id, modo] of [["tab-manual", "manual"], ["tab-batch", "batch"]]) {
  lista[id].addEventListener("click", () => alternarModo(modo));
  lista[id].addEventListener("keydown", (evento) => {
    if (["ArrowLeft", "ArrowRight"].includes(evento.key)) {
      evento.preventDefault();
      const outro = modo === "manual" ? "batch" : "manual";
      alternarModo(outro);
      lista[outro === "batch" ? "tab-batch" : "tab-manual"].focus();
    }
  });
}

async function enviarComando(comando, dados) {
  const resposta = await chrome.runtime.sendMessage({ tipo: "lote", comando, dados });
  if (!resposta?.ok) throw new Error(resposta?.erro || "O controlador da lista nao respondeu.");
  const salvo = await chrome.storage.local.get("lote");
  lote = salvo.lote;
}

lista["batch-start"].addEventListener("click", () => executar(async () => {
  if (!lista["batch-interval"].checkValidity()) throw new Error("Intervalo deve estar entre 3 e 300 segundos, em passos de 0,1 segundo.");
  if (consultaColecao) prepararColecao(colecao?.variantes);
  else prepararListaDeck(lista["batch-text"].value, lista["batch-commander"].value, Boolean(colecao));
  const permitido = await chrome.permissions.request({ origins: ORIGENS_CONSULTA });
  if (!permitido) throw new Error("Permissao recusada. A captura manual continua disponivel.");
  await enviarComando("iniciar", { texto: lista["batch-text"].value, intervalo: Number(lista["batch-interval"].value), somenteNormal: elementos.normal.checked,
    atualizarPrecos: lista["batch-refresh"].checked,
    comandante: consultaColecao ? "" : lista["batch-commander"].value,
    importacaoCsv: Boolean(colecao), colecao: consultaColecao ? { ...colecao, criterios: "variantes" } : null });
  mensagem("Lista iniciada.", "success");
}));

for (const [id, comando] of [["batch-pause", "pausar"], ["batch-resume", "retomar"], ["batch-retry", "retentar"], ["batch-skip", "pular"], ["batch-cancel", "cancelar"]]) {
  lista[id].addEventListener("click", () => executar(async () => {
    if (["retomar", "retentar"].includes(comando) && !await chrome.permissions.request({ origins: ORIGENS_CONSULTA })) throw new Error("Permissao recusada.");
    await enviarComando(comando);
    mensagem("Comando aplicado.");
  }));
}

lista["batch-tab"].addEventListener("click", () => executar(async () => {
  const aba = await chrome.tabs.update(lote.abaId, { active: true });
  await chrome.windows.update(aba.windowId, { focused: true });
}));

lista["collection-mode"].addEventListener("change", (evento) => {
  const ativar = evento.currentTarget.checked;
  return executar(async () => {
    consultaColecao = ativar;
    if (consultaColecao) elementos.normal.checked = false;
    if (consultaColecao && colecao?.texto) lista["batch-text"].value = colecao.texto;
    await chrome.storage.local.set({ consultaColecao, somenteNormal: elementos.normal.checked, rascunhoLista: lista["batch-text"].value });
    mensagem(consultaColecao ? "Modo colecao ativado: preco medio amarelo por edicao e acabamento do CSV." : "Modo colecao desativado: consulta comum pelo menor preco.");
  });
});

lista["import-list"].addEventListener("click", () => lista["list-file"].click());
lista["list-file"].addEventListener("change", () => executar(async () => {
  const arquivo = lista["list-file"].files[0];
  if (!arquivo) return;
  const csv = /\.csv$/i.test(arquivo.name) || arquivo.type === "text/csv";
  if (arquivo.size > (csv ? 2000000 : 200000)) throw new Error(csv ? "O CSV deve ter no maximo 2 MB." : "O TXT deve ter no maximo 200 KB.");
  const texto = await arquivo.text();
  const dados = csv ? importarCsvColecao(texto) : null;
  const itens = interpretarLista(dados ? dados.texto : texto, csv);
  colecao = dados ? { arquivo: arquivo.name, proxiesExcluidos: dados.proxiesExcluidos, variantes: dados.variantes, texto: dados.texto } : null;
  lista["batch-text"].value = dados ? dados.texto : texto;
  await chrome.storage.local.set({ rascunhoLista: lista["batch-text"].value, rascunhoColecao: colecao });
  mensagem(itens.length + " cartas distintas importadas" + (dados ? ` / ${dados.quantidade} unidades.` : "."), "success");
  lista["list-file"].value = "";
}));
lista["batch-text"].addEventListener("change", () => {
  chrome.storage.local.set({ rascunhoLista: lista["batch-text"].value, rascunhoColecao: colecao })
    .then(() => renderizar()).catch((erro) => mensagem(erro.message, "error"));
});
lista["batch-commander"].addEventListener("change", () => {
  chrome.storage.local.set({ rascunhoComandante: lista["batch-commander"].value })
    .then(() => renderizar()).catch((erro) => mensagem(erro.message, "error"));
});

chrome.storage.onChanged.addListener((mudancas, area) => {
  if (area !== "local") return;
  atualizarAvisoMemoria().catch(console.error);
  if (mudancas.capturas) capturas = mudancas.capturas.newValue || [];
  if (mudancas.lote) lote = mudancas.lote.newValue || null;
  if (mudancas.capturas || mudancas.lote) renderizar();
});

async function salvar(proximas) {
  totalCentavos(proximas);
  await chrome.storage.local.set({ capturas: proximas });
  capturas = proximas;
}

async function executar(acao) {
  if (ocupado) return;
  ocupado = true;
  renderizar();
  atualizarAvisoMemoria().catch(console.error);
  try {
    await acao();
  } catch (erro) {
    mensagem(erro.message || "Nao foi possivel concluir a operacao.", "error");
  } finally {
    ocupado = false;
    renderizar();
  }
}

elementos.capture.addEventListener("click", () => executar(async () => {
  const permitido = await chrome.permissions.request({ origins: ORIGENS_TAGGER });
  mensagem("Lendo a carta aberta...");
  const [aba] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!aba?.id) throw new Error("Nenhuma aba ativa disponivel.");
  const resultados = await chrome.scripting.executeScript({ target: { tabId: aba.id }, func: capturarPagina });
  const pagina = resultados[0]?.result;
  if (!pagina) throw new Error("Nao foi possivel ler a carta. Abra a pagina, conclua a verificacao e tente novamente.");
  const preco = extrairCotacao(pagina, elementos.normal.checked);
  mensagem("Consultando categorias no Scryfall Tagger...");
  const cotacao = permitido ? await enriquecerCotacao(preco)
    : { ...preco, categoriasTagger: [], statusTagger: "erro", erroTagger: "Acesso ao Scryfall Tagger nao autorizado." };
  const existente = capturas.some((item) => item.id === cotacao.id);
  await salvar(adicionarCaptura(capturas, cotacao));
  mensagem(cotacao.nome + (existente ? " atualizada: " : " capturada: ") + reais(cotacao.centavos)
    + (cotacao.statusTagger === "erro" ? ". Aviso: " + cotacao.erroTagger : "."), cotacao.statusTagger === "erro" ? "" : "success");
}));

elementos.normal.addEventListener("change", () => executar(async () => {
  await chrome.storage.local.set({ somenteNormal: elementos.normal.checked });
  mensagem("Filtro aplicado \u00e0s pr\u00f3ximas capturas.");
}));

function baixarExcel(conteudo, prefixo) {
  const blob = new Blob([conteudo], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = prefixo + new Date().toISOString().slice(0, 10) + ".xlsx";
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}

async function prepararGaleria() {
  if (!await chrome.permissions.request({ origins: ORIGENS_IMAGENS })) {
    throw new Error("Autorize o acesso ao Scryfall e as imagens para exportar a galeria.");
  }
  mensagem("Preparando galeria de imagens...");
  return {
    comandantes: lote?.comandantes || [],
    onProgresso: (atual, total) => mensagem("Incorporando imagens: " + atual + " de " + total + "."),
    onMetricas: (metricas) => {
      console.info("LigaMagic: desempenho da exportacao", metricas);
      return chrome.storage.local.set({ metricasExportacao: { ...metricas, data: new Date().toISOString() } });
    },
  };
}

elementos.export.addEventListener("click", () => executar(async () => {
  const exportaLote = document.body.dataset.mode === "batch" && lote;
  const opcoes = await prepararGaleria();
  baixarExcel(await (exportaLote ? (lote.colecao ? gerarExcelColecao(lote, opcoes) : gerarExcelLote(lote, opcoes)) : gerarExcel(capturas, opcoes)),
    exportaLote ? (lote.colecao ? "ligamagic-colecao-" : "ligamagic-lista-") : "ligamagic-");
  mensagem("Exporta\u00e7\u00e3o Excel iniciada.", "success");
}));

lista["batch-export"].addEventListener("click", () => executar(async () => {
  const opcoes = await prepararGaleria();
  baixarExcel(await (lote.colecao ? gerarExcelColecao(lote, opcoes) : gerarExcelLote(lote, opcoes)), lote.colecao ? "ligamagic-colecao-" : "ligamagic-lista-");
  mensagem("Relatorio Excel da lista exportado.", "success");
}));

elementos.clear.addEventListener("click", () => {
  document.getElementById("clear-title").textContent = document.body.dataset.mode === "batch" ? "Limpar a lista atual?" : "Limpar todas as capturas?";
  elementos["clear-dialog"].showModal();
});
elementos["cancel-clear"].addEventListener("click", () => elementos["clear-dialog"].close());
elementos["confirm-clear"].addEventListener("click", () => executar(async () => {
  if (document.body.dataset.mode === "batch") {
    await enviarComando("limpar");
    colecao = null;
    lista["batch-text"].value = "";
    lista["batch-commander"].value = "";
  } else await salvar([]);
  elementos["clear-dialog"].close();
  mensagem("Lista limpa.");
}));

const intervaloMemoria = window.setInterval(() => atualizarAvisoMemoria().catch(console.error), 5000);
window.addEventListener("pagehide", () => window.clearInterval(intervaloMemoria), { once: true });
atualizarAvisoMemoria().catch(console.error);

try {
  const dados = await chrome.storage.local.get({ capturas: [], somenteNormal: false, lote: null, rascunhoLista: "", rascunhoComandante: "", rascunhoColecao: null, consultaColecao: false });
  totalCentavos(dados.capturas);
  capturas = dados.capturas;
  lote = dados.lote;
  lista["batch-interval"].value = lote?.intervalo ?? 4;
  lista["batch-text"].value = loteAberto(lote) ? lote.texto : dados.rascunhoLista || lote?.texto || "";
  lista["batch-commander"].value = loteAberto(lote) ? (lote.comandantes || []).join(" + ") : dados.rascunhoComandante || (lote?.comandantes || []).join(" + ");
  colecao = loteAberto(lote) ? lote.colecao || null : dados.rascunhoLista ? dados.rascunhoColecao : lote?.colecao || null;
  if (loteAberto(lote) && lote.colecao?.criterios === "variantes") colecao = dados.rascunhoColecao || colecao;
  consultaColecao = loteAberto(lote) ? lote.colecao?.criterios === "variantes" : Boolean(dados.consultaColecao);
  alternarModo(lote ? "batch" : "manual");
  elementos.normal.checked = Boolean(dados.somenteNormal) && !consultaColecao;
  pronto = true;
  mensagem(capturas.length ? "Capturas locais carregadas." : "Nenhuma captura salva.");
} catch (erro) {
  mensagem("Falha ao carregar capturas: " + erro.message, "error");
} finally {
  ocupado = false;
  renderizar();
}