import { parse } from "./assets/csv-parse.mjs";
import { consultarJsonScryfall } from "./cache.mjs";

export function capturarPagina(lerResultados = false, colecao = false) {
  const endereco = new URL(location.href);
  if (endereco.protocol !== "https:" || !["ligamagic.com.br", "www.ligamagic.com.br"].includes(endereco.hostname)
      || endereco.searchParams.get("view") !== "cards/card") {
    throw new Error("Abra uma pagina de carta da LigaMagic nesta aba.");
  }
  const titulo = document.querySelector('meta[property="og:title"]')?.content;
  if (lerResultados && ![...document.scripts].some((script) => /\b(?:var|let|const)\s+cards_editions\s*=/.test(script.textContent))) {
    if (/just a moment|security verification|verifica[c\u00e7][a\u00e3]o de seguran[c\u00e7]a/i.test(document.title || "")
        || document.querySelector("#challenge-running, #challenge-form")) {
      return { erroCaptura: "Conclua a verificacao da LigaMagic na aba da lista e retome." };
    }
    const resultados = new Map();
    for (const link of document.querySelectorAll("a[href]")) {
      try {
        const href = link.getAttribute("href")?.trim();
        if (!href || href.startsWith("#")) continue;
        const url = new URL(href, endereco);
        const nome = url.searchParams.get("card")?.trim();
        if (url.protocol !== "https:" || !["ligamagic.com.br", "www.ligamagic.com.br"].includes(url.hostname)
            || url.searchParams.get("view") !== "cards/card" || !nome) continue;
        url.hash = "";
        resultados.set(url.href, { nome, url: url.href });
      } catch {}
    }
    const edicoesResultados = lerJson("cardsjson", true).flatMap((carta) => {
      if (!carta.nEN || !carta.sSigla || !carta.idE || !carta.sN) return [];
      const url = new URL("/", endereco);
      url.search = new URLSearchParams({ view: "cards/card", card: carta.nEN, ed: String(carta.idE), num: String(carta.sN) }).toString();
      return [{ nome: carta.nEN, code: carta.sSigla, num: String(carta.sN), name: carta.sEdicao || "", url: url.href,
        price: { 0: { p: carta.p1a ?? carta.precoMenor, m: carta.p1b } } }];
    });
    return { resultados: [...resultados.values()], ...(edicoesResultados.length ? { edicoesResultados } : {}) };
  }
  if (!titulo || !titulo.includes(" | ")) {
    throw new Error("Carta indisponivel. Conclua a verificacao no navegador e aguarde a pagina carregar.");
  }
  function lerJson(variavel, opcional = false) {
    for (const script of document.scripts) {
      const texto = script.textContent;
      const match = new RegExp("\\b(?:var|let|const)\\s+" + variavel + "\\s*=\\s*").exec(texto);
      if (!match) continue;
      const inicio = match.index + match[0].length;
      if (!["[", "{"].includes(texto[inicio])) throw new Error("Formato de dados inesperado: " + variavel);
      let nivel = 0;
      let emString = false;
      let escape = false;
      for (let posicao = inicio; posicao < texto.length; posicao++) {
        const caractere = texto[posicao];
        if (emString) {
          if (escape) escape = false;
          else if (caractere === "\\") escape = true;
          else if (caractere === '"') emString = false;
        } else if (caractere === '"') emString = true;
        else if (caractere === "[" || caractere === "{") nivel++;
        else if ((caractere === "]" || caractere === "}") && --nivel === 0) {
          return JSON.parse(texto.slice(inicio, posicao + 1));
        }
      }
      throw new Error("Dados incompletos na pagina: " + variavel);
    }
    if (opcional) return [];
    throw new Error("Precos ausentes. Aguarde a carta carregar ou conclua a verificacao do site.");
  }
  const nomes = titulo.split(" | ")[0].split(" / ").map((nome) => nome.trim());
  const nome = nomes[1] || nomes[0];
  if (!nome) throw new Error("Nome da carta nao encontrado.");
  const url = new URL("https://" + endereco.hostname + "/");
  url.search = new URLSearchParams({ view: "cards/card", card: nome, tipo: "1" }).toString();
  return {
    nome,
    nomePt: nomes[0],
    url: endereco.searchParams.has("ed") && endereco.searchParams.has("num") ? endereco.href : url.href,
    capturadoEm: new Date().toISOString(),
    edicoes: lerJson("cards_editions"),
    extras: lerJson("dataExtras", true),
    ...(colecao ? { ofertas: lerJson("cards_stock", true), idiomas: lerJson("dataLanguage", true) } : {}),
  };
}

export function normalizarNome(nome) {
  return nome.normalize("NFC").trim().replace(/\s+/g, " ").toLocaleLowerCase("pt-BR");
}

export function nomesEquivalentes(primeiro, segundo) {
  const comparar = (nome) => {
    const base = nome.replace(/\s+\(#[\w-]+\)$/, "");
    return normalizarNome(terrenoBasico(base) ? base : nome).replace(/\u00e6/g, "ae").replace(/\s*\/\/\s*/g, " // ");
  };
  const nomePrimeiro = comparar(primeiro);
  const nomeSegundo = comparar(segundo);
  if (nomePrimeiro === nomeSegundo) return true;
  const facesPrimeiro = nomePrimeiro.split(" // ");
  const facesSegundo = nomeSegundo.split(" // ");
  return (facesPrimeiro.length === 2 && facesSegundo.length === 1 && facesPrimeiro.includes(nomeSegundo))
    || (facesSegundo.length === 2 && facesPrimeiro.length === 1 && facesSegundo.includes(nomePrimeiro));
}

export function importarCsvColecao(texto) {
  let delimitador;
  let colunas;
  for (const candidato of [",", ";", "\t"]) {
    try {
      const cabecalho = parse(texto, { bom: true, delimiter: candidato, to: 1, skip_empty_lines: true })[0] || [];
      const nomes = cabecalho.map((nome) => nome.trim().toLowerCase());
      if (nomes.includes("name") && nomes.includes("quantity")) {
        delimitador = candidato;
        colunas = nomes;
        break;
      }
    } catch {}
  }
  if (!delimitador) throw new Error("CSV de colecao deve conter as colunas Name e Quantity (ManaBox).");
  if (new Set(colunas).size !== colunas.length) throw new Error("CSV contem colunas duplicadas.");
  let registros;
  try {
    registros = parse(texto, { bom: true, delimiter: delimitador, skip_empty_lines: true, trim: true });
  } catch {
    throw new Error("CSV invalido: confira as aspas e o numero de colunas.");
  }
  const linhas = [];
  const variantes = new Map();
  let proxiesExcluidos = 0;
  for (const [indice, registro] of registros.slice(1).entries()) {
    const nome = registro[colunas.indexOf("name")].trim();
    const valor = registro[colunas.indexOf("quantity")].trim();
    const quantidade = Number(valor);
    if (!nome || /[\r\n]/.test(nome) || !/^\d+$/.test(valor) || !Number.isSafeInteger(quantidade) || quantidade < 1 || quantidade > 9999) {
      throw new Error("Nome ou quantidade invalida no registro " + (indice + 2) + " do CSV.");
    }
    const proxy = colunas.includes("proxy") ? registro[colunas.indexOf("proxy")].trim().toLowerCase() : "false";
    if (!["", "false", "true"].includes(proxy)) throw new Error("Valor Proxy invalido no registro " + (indice + 2) + " do CSV.");
    if (proxy === "true") {
      proxiesExcluidos += quantidade;
      continue;
    }
    linhas.push(quantidade + " " + nome);
    const campo = (coluna) => colunas.includes(coluna) ? registro[colunas.indexOf(coluna)].trim() : "";
    const variante = { nome, codigoEdicao: campo("set code").toLowerCase(), edicao: campo("set name"),
      numero: campo("collector number"), acabamento: campo("foil").toLowerCase(), idioma: campo("language").toLowerCase() };
    if (campo("scryfall id")) variante.scryfallId = campo("scryfall id");
    const arte = campo("full art") || campo("full_art");
    if (arte) {
      if (!["true", "false"].includes(arte.toLowerCase())) throw new Error("Full Art invalido no CSV.");
      variante.fullArt = arte.toLowerCase() === "true";
    }
    const chave = JSON.stringify({ ...variante, nome: normalizarNome(nome) });
    const anterior = variantes.get(chave);
    const soma = quantidade + (anterior?.quantidade || 0);
    if (soma > 9999) throw new Error("Quantidade acima de 9999 para " + nome + ".");
    variantes.set(chave, { ...variante, quantidade: soma });
  }
  if (!linhas.length) throw new Error("Nenhuma carta nao-proxy encontrada no CSV.");
  const itens = interpretarLista(linhas.join("\n"), true);
  return {
    texto: itens.map((item) => item.quantidade + " " + item.nome).join("\n"),
    distintas: itens.length,
    quantidade: itens.reduce((total, item) => total + item.quantidade, 0),
    proxiesExcluidos,
    variantes: [...variantes.values()],
  };
}

export function prepararColecao(variantes) {
  if (!Array.isArray(variantes) || !variantes.length || variantes.length > 5000) throw new Error("Importe um CSV com ate 5000 variantes de cartas.");
  const itens = new Map();
  for (const variante of variantes) {
    const campos = ["nome", "codigoEdicao", "edicao", "numero", "acabamento", "idioma"];
    if (campos.some((campo) => typeof variante[campo] !== "string")
        || ["nome", "codigoEdicao", "numero", "idioma"].some((campo) => !variante[campo].trim())
        || !["normal", "foil", "etched"].includes(variante.acabamento)
        || !Number.isSafeInteger(variante.quantidade) || variante.quantidade < 1 || variante.quantidade > 9999) {
      throw new Error("O modo colecao exige Name, Set code, Collector number, Foil, Language e Quantity validos. Reimporte o CSV completo.");
    }
    const dados = Object.fromEntries(campos.map((campo) => [campo, variante[campo].trim()]));
    dados.codigoEdicao = dados.codigoEdicao.toLowerCase();
    dados.idioma = dados.idioma.toLowerCase();
    if (variante.scryfallId) dados.scryfallId = String(variante.scryfallId);
    if (typeof variante.fullArt === "boolean") dados.fullArt = variante.fullArt;
    if (Array.isArray(variante.referenciasTerreno)) dados.referenciasTerreno = variante.referenciasTerreno;
    if (variante.tratamento) dados.tratamento = variante.tratamento;
    const idColecao = "colecao:" + JSON.stringify([normalizarNome(dados.nome), dados.codigoEdicao, dados.numero, dados.acabamento, dados.idioma]);
    const quantidade = variante.quantidade + (itens.get(idColecao)?.quantidade || 0);
    if (quantidade > 9999) throw new Error("Quantidade acima de 9999 para " + dados.nome + ".");
    itens.set(idColecao, { ...dados, idColecao, quantidade, status: "pendente", erro: "" });
  }
  return [...itens.values()];
}

export function terrenoBasico(nome) {
  return /^(?:(?:snow-covered )?(?:plains|island|swamp|mountain|forest)|wastes|planicie|ilha|pantano|montanha|floresta|ermos)$/.test(
    normalizarNome(nome).normalize("NFD").replace(/\p{Diacritic}/gu, ""));
}

export async function consultarReferenciasColecao(item, requisitar = fetch, cache = new Map()) {
  const basico = terrenoBasico(item.nome);
  if (!basico && (!item.codigoEdicao || !item.numero)) return {};
  const obter = async (caminho) => {
    if (!cache.has(caminho)) cache.set(caminho, consultarJsonScryfall("https://api.scryfall.com" + caminho, requisitar,
      requisitar === globalThis.fetch ? undefined : async () => {}));
    try { return await cache.get(caminho); }
    catch (erro) { cache.delete(caminho); throw erro; }
  };
  const porNome = !item.codigoEdicao || !item.numero;
  const caminho = !porNome ? "/cards/" + encodeURIComponent(item.codigoEdicao) + "/" + encodeURIComponent(item.numero)
    : /^[\da-f-]{36}$/i.test(item.scryfallId || "") ? "/cards/" + item.scryfallId
      : "/cards/named?" + new URLSearchParams({ exact: item.nome });
  const carta = await obter(caminho);
  if (![carta.name, carta.printed_name].filter(Boolean).some((nome) => nomesEquivalentes(nome, item.nome)) || typeof carta.full_art !== "boolean") {
    throw new Error("Scryfall nao confirmou a identidade e a arte da carta do CSV.");
  }
  const numero = (valor) => String(valor).toLowerCase().replace(/^0+(?=\d)/, "");
  if (!porNome && (carta.set?.toLowerCase() !== item.codigoEdicao.toLowerCase() || numero(carta.collector_number) !== numero(item.numero))) {
    throw new Error("Scryfall retornou uma edicao ou numero diferente do CSV.");
  }
  const tratamento = (valor) => valor.full_art ? "fullart" : valor.border_color === "borderless" ? "borderless"
    : valor.frame_effects?.includes("extendedart") ? "extendedart"
      : valor.frame_effects?.includes("showcase") ? "showcase:" + [...valor.frame_effects].sort().join(",") : "padrao";
  const fullArt = porNome ? false : carta.full_art;
  const acabamento = item.acabamento || "normal";
  if (!porNome && !carta.finishes?.includes(acabamento === "normal" ? "nonfoil" : acabamento)) {
    throw new Error("Scryfall nao confirmou o acabamento da carta do CSV.");
  }
  const consulta = new URLSearchParams({ q: '!"' + carta.name + '" game:paper ' + (basico ? "t:basic " + (fullArt ? "is:fullart" : "-is:fullart") : ""), unique: "prints", order: "released", dir: "desc" });
  const catalogo = await obter("/cards/search?" + consulta);
  if (!Array.isArray(catalogo.data)) throw new Error("Catalogo de impressoes indisponivel no Scryfall.");
  const referenciasTerreno = [carta, ...catalogo.data].filter((valor) => valor.name === carta.name && valor.full_art === fullArt
    && (basico || tratamento(valor) === tratamento(carta))
    && !valor.digital && valor.finishes?.includes(acabamento === "normal" ? "nonfoil" : acabamento))
    .map((valor) => ({ codigoEdicao: valor.set, numero: valor.collector_number, fullArt: valor.full_art, dataLancamento: valor.released_at,
      ...(!basico ? { tratamento: tratamento(valor) } : {}) }));
  return { fullArt, referenciasTerreno, ...(!basico ? { tratamento: tratamento(carta) } : {}) };
}

export async function consultarReferenciasTerreno(item, requisitar = fetch, cache = new Map()) {
  return terrenoBasico(item.nome) ? consultarReferenciasColecao(item, requisitar, cache) : {};
}

export function ordenarEdicoesColecao(edicoes, criterio, extra = null, campoPreco = "m") {
  const normalizar = (valor) => String(valor ?? "").toLowerCase().normalize("NFD").replace(/\p{Diacritic}/gu, "");
  const numero = (valor) => normalizar(valor).replace(/^0+(?=\d)/, "");
  const correspondeCodigo = (codigo, base) => Boolean(base) && codigo.endsWith(normalizar(base));
  const nomeEdicao = normalizar(criterio.edicao || (criterio.codigoEdicao === "fdn" ? "Foundations" : ""));
  const comum = terrenoBasico(criterio.nome) && criterio.acabamento === "normal" && criterio.fullArt === false;
  const candidatos = edicoes.flatMap((edicao) => {
    const codigo = normalizar(edicao.code);
    const nome = normalizar(edicao.name);
    const relacao = codigo === criterio.codigoEdicao ? 0 : correspondeCodigo(codigo, criterio.codigoEdicao) ? 1
      : codigo.includes(criterio.codigoEdicao) ? 2 : nomeEdicao && nome && (nome.includes(nomeEdicao) || nomeEdicao.includes(nome)) ? 3 : 4;
    const mesmoNumero = numero(edicao.num) === numero(criterio.numero);
    if (!terrenoBasico(criterio.nome) && (relacao > 1 || !mesmoNumero)) return [];
    const referencias = criterio.referenciasTerreno?.filter((referencia) => numero(referencia.numero) === numero(edicao.num)) || [];
    const meta = referencias.find((referencia) => codigo === normalizar(referencia.codigoEdicao))
      || referencias.find((referencia) => normalizar(referencia.codigoEdicao) === criterio.codigoEdicao && correspondeCodigo(codigo, referencia.codigoEdicao))
      || referencias.find((referencia) => correspondeCodigo(codigo, referencia.codigoEdicao));
    if (!terrenoBasico(criterio.nome) && relacao === 1
        && (!criterio.tratamento || meta?.tratamento !== criterio.tratamento || normalizar(meta?.codigoEdicao) !== criterio.codigoEdicao)) return [];
    const impressaoExata = relacao === 0 && mesmoNumero;
    if (criterio.tratamento && !impressaoExata && meta?.tratamento !== criterio.tratamento) return [];
    const fullArt = meta?.fullArt ?? edicao.fullArt ?? (/full[ -]?art/i.test(edicao.name || "") ? true : undefined);
    if (criterio.fullArt === true && (fullArt === false || fullArt !== true && !(relacao <= 3 && mesmoNumero))) return [];
    if (criterio.fullArt === false && fullArt === true) return [];
    if (terrenoBasico(criterio.nome) && criterio.fullArt === false && !comum && fullArt !== false && !(relacao <= 3 && mesmoNumero)) return [];
    const data = meta?.dataLancamento || edicao.dataLancamento || edicao.released_at;
    const lancamento = /^\d{4}-\d{2}-\d{2}$/.test(data || "") ? Date.parse(data) : NaN;
    if (comum && (fullArt !== false || !Number.isFinite(lancamento) || lancamento > Date.now())) return [];
    if (extra !== null) {
      const valor = String(edicao.price?.[extra]?.[campoPreco] ?? "");
      if (!/^\d+(?:\.\d{1,2})?$/.test(valor) || Number(valor) <= 0 || !Number.isSafeInteger(Math.round(Number(valor) * 100))) return [];
    }
    const prioridade = relacao * 2 + (mesmoNumero ? 0 : 1);
    return [{ ...edicao, fullArt, lancamento, prioridade, motivoReferencia: comum ? "Terreno basico comum: versao mais recente disponivel"
      : relacao === 0 && mesmoNumero ? "Edicao e numero exatos" : relacao <= 3 ? "Referencia relacionada a " + criterio.codigoEdicao.toUpperCase()
        : "Edicao substituta com acabamento compativel" }];
  });
  return candidatos.sort((primeira, segunda) => comum ? segunda.lancamento - primeira.lancamento || primeira.prioridade - segunda.prioridade
    : primeira.prioridade - segunda.prioridade || (Number.isFinite(segunda.lancamento) && Number.isFinite(primeira.lancamento) ? segunda.lancamento - primeira.lancamento : 0));
}

export function extrairCotacaoColecao(pagina, item) {
  const [criterio] = prepararColecao([item]);
  conferirNome(pagina, criterio.nome);
  const rotulo = { normal: "Normal", foil: "Foil", etched: "Foil Especial / Foil Etched" }[criterio.acabamento];
  const extra = criterio.acabamento === "normal" ? 0 : (pagina.extras || []).find((valor) => valor.label === rotulo)?.id;
  if (extra === undefined) throw new Error("Acabamento do CSV nao identificado: " + criterio.acabamento + ".");
  const [edicao] = ordenarEdicoesColecao(pagina.edicoes || [], criterio, extra);
  if (!edicao) throw new Error("Preco medio amarelo indisponivel para " + (terrenoBasico(criterio.nome) ? "uma edicao compativel com o CSV."
    : criterio.codigoEdicao.toUpperCase() + " #" + criterio.numero + " / " + rotulo + ". Nenhuma outra impressao foi usada."));
  const valor = String(edicao.price?.[extra]?.m ?? "");
  if (!/^\d+(?:\.\d{1,2})?$/.test(valor)) throw new Error("Preco medio amarelo indisponivel para a edicao e acabamento do CSV.");
  const [inteiro, fracao = ""] = valor.split(".");
  const centavos = Number(inteiro) * 100 + Number(fracao.padEnd(2, "0"));
  if (!Number.isSafeInteger(centavos) || centavos <= 0) throw new Error("Preco medio amarelo indisponivel para a edicao e acabamento do CSV.");
  return { id: criterio.idColecao, nome: pagina.nome, nomePt: pagina.nomePt, url: pagina.url, capturadoEm: pagina.capturadoEm,
    centavos, quantidade: 1, edicao: edicao.name, variante: rotulo, idioma: criterio.idioma, idiomaPreco: "geral",
    codigoEdicao: edicao.code, numero: String(edicao.num), tipoPreco: "media_ligamagic",
    ...(typeof criterio.fullArt === "boolean" ? { fullArt: criterio.fullArt } : {}),
    ...(criterio.tratamento ? { tratamento: criterio.tratamento } : {}),
    referenciaSubstituida: String(edicao.code).toLowerCase() !== criterio.codigoEdicao || String(edicao.num) !== criterio.numero,
    motivoReferencia: edicao.motivoReferencia };
}

export function interpretarLista(texto, colecao = false) {
  const itens = new Map();
  for (const [indice, original] of texto.replace(/^\uFEFF/, "").split(/\r?\n/).entries()) {
    const linha = original.trim();
    if (!linha || linha.startsWith("#") || linha.startsWith("//")) continue;
    const partes = /^(\d+)(?:[xX]\s*|\s+[xX]\s+|\s+)(.+)$/.exec(linha);
    const nome = (partes ? partes[2] : linha).trim();
    const quantidade = partes ? Number(partes[1]) : 1;
    if (!nome || !Number.isSafeInteger(quantidade) || quantidade < 1 || quantidade > 9999 || /^\d+$/.test(nome)) {
      throw new Error("Quantidade ou nome invalido na linha " + (indice + 1) + ".");
    }
    const id = normalizarNome(nome);
    const anterior = itens.get(id);
    const soma = quantidade + (anterior?.quantidade || 0);
    if (soma > 9999) throw new Error("Quantidade acima de 9999 para " + nome + ".");
    itens.set(id, { nome: anterior?.nome || nome, quantidade: soma, status: "pendente", erro: "" });
    const limite = colecao ? 5000 : 500;
    if (itens.size > limite) throw new Error("Importe no maximo " + limite + " cartas distintas por lista.");
  }
  if (!itens.size) throw new Error("Informe pelo menos uma carta.");
  return [...itens.values()];
}

export function conferirNome(pagina, solicitado) {
  if (![pagina.nome, pagina.nomePt].filter(Boolean).some((nome) => nomesEquivalentes(nome, solicitado))) {
    throw new Error("A pagina aberta nao corresponde a " + solicitado + ". Confira o nome exato em portugues ou ingles.");
  }
}

export function prepararListaDeck(texto, comandante = "", importacaoCsv = false) {
  if (typeof comandante !== "string") throw new Error("Informe o nome do comandante.");
  const escolhidos = comandante.trim() ? interpretarLista(comandante.split(/\s*\+\s*|\r?\n/).join("\n")) : [];
  if (escolhidos.length > 2 || escolhidos.some((item) => item.quantidade !== 1)) {
    throw new Error("Informe um comandante ou dois parceiros, com uma copia de cada.");
  }
  const comandantes = escolhidos.map((item) => item.nome);
  const itens = interpretarLista(texto.trim() ? texto : comandantes.join("\n"), importacaoCsv);
  for (const comandante of escolhidos) {
    const existente = itens.find((item) => nomesEquivalentes(item.nome, comandante.nome));
    if (existente && existente.quantidade !== 1) throw new Error("O comandante deve ter uma unica copia: " + existente.nome + ".");
    if (!existente) itens.push(comandante);
  }
  const limite = importacaoCsv ? 5000 : 500;
  if (itens.length > limite) throw new Error("Importe no maximo " + limite + " cartas distintas por lista.");
  return { itens, comandantes };
}

export function foraDoTotalDeck(item, comandantes = []) {
  return terrenoBasico(item.nome) || comandantes.some((comandante) => nomesEquivalentes(item.nome, comandante));
}

export function extrairCotacao(pagina, somenteNormal = false) {
  if (!Array.isArray(pagina.edicoes)) throw new Error("Lista de edicoes invalida.");
  let menor = null;
  for (const edicao of pagina.edicoes) {
    for (const [extra, precos] of Object.entries(edicao?.price || {})) {
      if (somenteNormal && extra !== "0") continue;
      const valor = String(precos?.p ?? "");
      const partes = /^(\d+)(?:\.(\d{1,2}))?$/.exec(valor);
      if (!partes) continue;
      const centavos = Number(partes[1]) * 100 + Number((partes[2] || "").padEnd(2, "0"));
      if (!Number.isSafeInteger(centavos) || centavos <= 0 || (menor && centavos >= menor.centavos)) continue;
      const variante = extra === "0" ? "Normal"
        : (pagina.extras || []).find((item) => String(item.id) === extra)?.label || "Extra " + extra;
      menor = { centavos, edicao: String(edicao.name || "Edicao desconhecida"), variante: String(variante),
        ...(edicao.code && edicao.num !== undefined ? { codigoEdicao: String(edicao.code).toLowerCase(), numero: String(edicao.num) } : {}) };
    }
  }
  if (!menor) throw new Error("Nenhum preco verde positivo encontrado para o filtro selecionado.");
  return {
    ...menor,
    id: normalizarNome(pagina.nome),
    nome: pagina.nome,
    nomePt: pagina.nomePt,
    url: pagina.url,
    capturadoEm: pagina.capturadoEm,
    somenteNormal,
    quantidade: 1,
  };
}

export function adicionarCaptura(capturas, nova) {
  const existente = capturas.find((item) => item.id === nova.id);
  const atualizada = { ...nova, quantidade: existente?.quantidade ?? 1 };
  return existente ? capturas.map((item) => item.id === nova.id ? atualizada : item) : [...capturas, atualizada];
}

export function totalCentavos(capturas) {
  let total = 0;
  for (const item of capturas) {
    if (!Number.isSafeInteger(item.quantidade) || item.quantidade < 1 || item.quantidade > 9999
        || !Number.isSafeInteger(item.centavos) || item.centavos <= 0) {
      throw new Error("Quantidade ou preco invalido.");
    }
    total += item.centavos * item.quantidade;
    if (!Number.isSafeInteger(total)) throw new Error("Total acima do limite suportado.");
  }
  return total;
}

export function reais(centavos) {
  return new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" }).format(centavos / 100);
}

export function gerarCsv(capturas) {
  const decimal = (centavos) => (centavos / 100).toFixed(2).replace(".", ",");
  const linhas = [["Nome da carta", "Pre\u00e7o m\u00ednimo", "Tipos", "Categorias (principal)"]];
  for (const item of capturas) {
    if (item.centavos !== null && (!Number.isSafeInteger(item.centavos) || item.centavos <= 0)) {
      throw new Error("Preco invalido.");
    }
    const principal = item.categoriasTagger?.length ? item.categoriasTagger.join(" | ") : "Land";
    linhas.push([item.nome, item.centavos === null ? "" : decimal(item.centavos), (item.tiposCarta || []).join(" | "), principal]);
  }
  return codificarCsv(linhas);
}

function codificarCsv(linhas) {
  const celula = (valor) => {
    let texto = String(valor);
    if (/^[\t\r\n]|^\s*[=+@-]/.test(texto)) texto = "'" + texto;
    return /[;"\r\n]/.test(texto) ? '"' + texto.replaceAll('"', '""') + '"' : texto;
  };
  return "\uFEFF" + linhas.map((linha) => linha.map(celula).join(";")).join("\r\n") + "\r\n";
}

export function gerarCsvLote(lote) {
  return gerarCsv(lote.itens.map((item) => item.status === "ok" && item.cotacao
    ? { ...item.cotacao, nome: item.nome } : { nome: item.nome, centavos: null }));
}

export function resumirColecao(lote) {
  const validas = lote.itens.filter((item) => item.status === "ok" && item.cotacao);
  const unidades = lote.itens.reduce((total, item) => total + item.quantidade, 0);
  const avaliadas = validas.reduce((total, item) => total + item.quantidade, 0);
  return {
    distintas: lote.itens.length, unidades, avaliadas, pendentes: unidades - avaliadas,
    centavos: totalCentavos(validas.map((item) => ({ ...item.cotacao, quantidade: item.quantidade }))),
  };
}

export function gerarCsvColecao(lote) {
  const decimal = (centavos) => (centavos / 100).toFixed(2).replace(".", ",");
  if (lote.colecao?.criterios === "variantes") {
    const linhas = [["Nome da carta", "Quantidade", "Edicao CSV", "Codigo edicao", "Numero", "Acabamento CSV", "Idioma CSV", "Preco medio unitario (BRL)", "Subtotal estimado (BRL)", "Fonte do preco", "Status", "Data da consulta", "Observacoes", "Base da estimativa", "Edicao usada", "Sigla usada", "Numero usado", "Referencia substituida", "Motivo da referencia", "URL da referencia"]];
    for (const item of lote.itens) {
      const cotacao = item.status === "ok" ? item.cotacao : null;
      const subtotal = cotacao ? totalCentavos([{ ...cotacao, quantidade: item.quantidade }]) : null;
      linhas.push([item.nome, item.quantidade, item.edicao, item.codigoEdicao, item.numero, item.acabamento, item.idioma,
        cotacao ? decimal(cotacao.centavos) : "", subtotal === null ? "" : decimal(subtotal), cotacao?.tipoPreco || "",
        cotacao ? "Avaliada" : "Pendente", cotacao?.capturadoEm || "", item.erro || "",
        cotacao?.tipoPreco === "media_ofertas" ? "Media aritmetica das ofertas da variante exata; sem ajuste por conservacao; sem frete"
          : "Preco medio amarelo da LigaMagic; terreno basico comum usa referencia recente; alternativas identificadas no relatorio; sem filtro por idioma; sem ajuste por conservacao; sem frete",
        cotacao?.edicao || "", cotacao?.codigoEdicao || "", cotacao?.numero || "", cotacao ? (cotacao.referenciaSubstituida ? "Sim" : "Nao") : "",
        cotacao?.motivoReferencia || "", cotacao?.url || ""]);
    }
    return codificarCsv(linhas);
  }
  const linhas = [["Nome da carta", "Quantidade", "Preco minimo unitario (BRL)", "Subtotal estimado (BRL)", "Status", "Edicao de referencia", "Acabamento de referencia", "Data da consulta", "Observacoes", "Base da estimativa"]];
  for (const item of lote.itens) {
    const cotacao = item.status === "ok" ? item.cotacao : null;
    const subtotal = cotacao ? totalCentavos([{ ...cotacao, quantidade: item.quantidade }]) : null;
    const status = cotacao ? "Avaliada" : item.status === "pulado" ? "Ignorada" : item.status === "erro" ? "Erro" : "Pendente";
    linhas.push([item.nome, item.quantidade, cotacao ? decimal(cotacao.centavos) : "", subtotal === null ? "" : decimal(subtotal),
      status, cotacao?.edicao || "", cotacao?.variante || "", cotacao?.capturadoEm || "", item.erro || cotacao?.erroTagger || "",
      "Menor preco por nome; sem ajuste por edicao, acabamento, idioma ou estado; sem frete"]);
  }
  return codificarCsv(linhas);
}