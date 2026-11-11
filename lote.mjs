import { capturarPagina, conferirNome, extrairCotacao, extrairCotacaoColecao, prepararColecao, prepararListaDeck, nomesEquivalentes, normalizarNome, totalCentavos, terrenoBasico, consultarReferenciasColecao, ordenarEdicoesColecao } from "./core.mjs";
import { enriquecerCotacao, ORIGENS_TAGGER } from "./tagger.mjs";
import { cacheLocal, criarCachePersistente, criarFila, VALIDADE, contar, medir, obterMetricas } from "./cache.mjs";
import { criarResolvedorImagens, ORIGENS_IMAGENS } from "./imagens.mjs";

export const ORIGENS = ["https://www.ligamagic.com.br/*", "https://ligamagic.com.br/*"];
export const ORIGENS_CONSULTA = [...ORIGENS, ...ORIGENS_TAGGER, "https://cards.scryfall.io/*"];
export const ALARME_LOTE = "ligamagic-lista";
export const loteAberto = (lote) => Boolean(lote && !["concluido", "cancelado"].includes(lote.fase));

export function criarControlador(api, agora = Date.now, enriquecer = enriquecerCotacao, resolverTerreno, opcoes = {}) {
  let fila = Promise.resolve();
  const cacheTerrenos = new Map();
  const resolverReferencia = resolverTerreno || ((item) => consultarReferenciasColecao(item, fetch, cacheTerrenos));
  const cache = opcoes.cache || (typeof indexedDB !== "undefined" ? cacheLocal : criarCachePersistente({ banco: null, agora }));
  const paginasRecentes = new Map();
  const filaAntecipacao = criarFila(4);
  const antecipar = opcoes.anteciparImagem || (typeof indexedDB !== "undefined" ? criarResolvedorImagens() : null);
  const serializar = (acao) => {
    const resposta = fila.then(acao);
    fila = resposta.catch(() => {});
    return resposta;
  };
  const carregar = async () => (await api.storage.local.get("lote")).lote;
  const salvar = (lote) => api.storage.local.set({ lote });
  const agendar = (quando) => api.alarms.create(ALARME_LOTE, { when: quando });

  const chavePagina = (item) => "pagina:" + normalizarNome(item.nome);
  const chaveCotacao = (lote, item) => "cotacao:" + JSON.stringify([normalizarNome(item.nome),
    lote.colecao?.criterios === "variantes" ? item.idColecao : Boolean(lote.somenteNormal)]);
  const complementar = async (lote, preco) => lote.colecao?.criterios === "variantes" ? preco
    : await api.permissions.contains({ origins: ORIGENS_TAGGER }) ? medir("enriquecimento", () => enriquecer(preco))
      : { ...preco, categoriasTagger: [], statusTagger: "erro", erroTagger: "Acesso ao Scryfall Tagger nao autorizado. Retome a lista para autorizar." };

  async function confirmarColecao(item) {
    const verificacao = JSON.stringify([3, normalizarNome(item.nome), item.codigoEdicao, item.numero, item.acabamento]);
    if (item.verificacaoScryfall === verificacao) return;
    Object.assign(item, await resolverReferencia(item));
    item.verificacaoScryfall = verificacao;
  }

  async function navegar(lote, url) {
    if (agora() < (lote.proximaConsultaEm || 0)) {
      lote.itens[lote.indice].proximaBusca = url;
      lote.fase = "aguardando";
      lote.proxima = lote.proximaConsultaEm;
      lote.mensagem = "Aguardando completar " + String(lote.intervalo ?? 4).replace(".", ",") + " segundos antes da proxima consulta.";
      await salvar(lote);
      await agendar(lote.proxima);
      return;
    }
    lote.proximaConsultaEm = agora() + (lote.intervalo ?? 4) * 1000;
    lote.metricas = lote.metricas || {};
    lote.metricas.navegacoes = (lote.metricas.navegacoes || 0) + 1;
    lote.inicioNavegacao = Date.now();
    await salvar(lote);
    await api.tabs.update(lote.abaId, { url });
  }

  async function reutilizar(lote) {
    const item = lote.itens[lote.indice];
    const colecao = lote.colecao?.criterios === "variantes";
    if (colecao) {
      try { await confirmarColecao(item); } catch { return null; }
    }
    if (!lote.atualizarPrecos) {
      const salva = await cache.get(chaveCotacao(lote, item));
      if (salva && (!colecao || salva.verificacaoScryfall === item.verificacaoScryfall)) {
        try { conferirNome(salva, item.nome); return salva; } catch {}
      }
    }
    if (terrenoBasico(item.nome)) return null;
    const pagina = paginasRecentes.get(chavePagina(item)) || (!lote.atualizarPrecos && await cache.get(chavePagina(item)));
    if (!pagina) return null;
    if (pagina.cachePrecoEm !== undefined && agora() - pagina.cachePrecoEm >= VALIDADE.precos) return null;
    try {
      conferirNome(pagina, item.nome);
      const preco = colecao ? extrairCotacaoColecao(pagina, item) : extrairCotacao(pagina, lote.somenteNormal);
      return await complementar(lote, { ...preco, cachePrecoEm: pagina.cachePrecoEm ?? agora() });
    } catch { return null; }
  }

  async function concluirItem(lote, cotacao, deCache = false) {
    const item = lote.itens[lote.indice];
    cotacao = { ...cotacao, cachePrecoEm: cotacao.cachePrecoEm ?? agora(),
      ...(lote.colecao?.criterios === "variantes" ? { verificacaoScryfall: item.verificacaoScryfall } : {}) };
    item.cotacao = cotacao;
    item.status = "ok";
    item.erro = "";
    lote.metricas = lote.metricas || {};
    if (deCache) { lote.metricas.cartasCache = (lote.metricas.cartasCache || 0) + 1; contar("precosCache"); }
    lote.metricas.servicos = obterMetricas(lote.metricasIniciais);
    if (lote.colecao?.criterios === "variantes" || cotacao.statusTagger === "ok") {
      await cache.set(chaveCotacao(lote, item), cotacao, Math.max(0, VALIDADE.precos - (agora() - cotacao.cachePrecoEm)));
    }
    if (antecipar) void api.permissions.contains({ origins: ORIGENS_IMAGENS }).then((permitido) => permitido
      ? filaAntecipacao(() => medir("antecipacaoImagens", () => antecipar({ ...item, ...cotacao, nome: item.nome }))) : null).catch(() => contar("falhasAntecipacao"));
    if (lote.colecao?.criterios === "variantes") {
      avancar(lote);
      await salvar(lote);
      return;
    }
    const { capturas = [] } = await api.storage.local.get("capturas");
    const quantidade = lote.itens.filter((entrada) => entrada.status === "ok" && entrada.cotacao.id === cotacao.id)
      .reduce((soma, entrada) => soma + entrada.quantidade, 0);
    const nova = { ...cotacao, quantidade };
    const proximas = capturas.some((entrada) => entrada.id === cotacao.id)
      ? capturas.map((entrada) => entrada.id === cotacao.id ? nova : entrada) : [...capturas, nova];
    totalCentavos(proximas);
    avancar(lote);
    await api.storage.local.set({ lote, capturas: proximas });
  }

  function chaveBusca(endereco) {
    const url = new URL(endereco);
    url.hostname = "www.ligamagic.com.br";
    url.hash = "";
    url.searchParams.sort();
    return url.href;
  }

  function paginaConfere(aba, lote) {
    try {
      const url = new URL(aba.url);
      return ORIGENS.some((origem) => url.origin + "/*" === origem)
        && url.searchParams.get("view") === "cards/card"
        && nomesEquivalentes(url.searchParams.get("card") || "", lote.itens[lote.indice].nome);
    } catch {
      return false;
    }
  }

  async function pausar(lote, motivo, erro = false) {
    lote.fase = "pausado";
    lote.mensagem = motivo;
    if (erro && lote.itens[lote.indice]) {
      lote.itens[lote.indice].status = "erro";
      lote.itens[lote.indice].erro = motivo;
    }
    await salvar(lote);
    await api.alarms.clear(ALARME_LOTE);
  }

  function avancar(lote) {
    lote.indice++;
    while (lote.itens[lote.indice]?.status === "ok") lote.indice++;
    lote.fase = lote.indice >= lote.itens.length ? "concluido" : "aguardando";
    lote.proxima = agora();
    lote.mensagem = lote.fase === "concluido" ? "Lista concluida." : "Iniciando a proxima carta.";
  }

  async function pendenciaColecao(lote, motivo) {
    const item = lote.itens[lote.indice];
    item.status = "erro";
    item.erro = motivo;
    avancar(lote);
    await salvar(lote);
    if (lote.fase === "concluido") await api.alarms.clear(ALARME_LOTE);
    else await agendar(lote.proxima);
  }

  async function abrirReferenciaTerreno(lote, referencia) {
    const item = lote.itens[lote.indice];
    lote.fase = "navegando";
    lote.prazo = agora() + 45000;
    lote.mensagem = "Abrindo " + item.nome + " / " + referencia.code.toUpperCase() + " #" + referencia.num + ".";
    await salvar(lote);
    await agendar(lote.prazo);
    await navegar(lote, referencia.url);
  }

  async function tentarOutraBusca(lote, aba, motivo) {
    const item = lote.itens[lote.indice];
    const faces = item.nome.split(/\s*\/\/\s*/).map((nome) => nome.trim()).filter(Boolean);
    if (faces.length !== 2) return false;
    const nomes = [...new Set([item.nome, ...faces])];
    const alternativas = [true, false].flatMap((comTipo) => nomes.map((nome) => {
      const url = new URL("https://www.ligamagic.com.br/");
      url.searchParams.set("view", "cards/card");
      url.searchParams.set("card", nome);
      if (comTipo) url.searchParams.set("tipo", "1");
      return url.href;
    }));
    const tentadas = [...new Set([...(item.buscasTentadas || []), ...(item.paginasResultados || []), aba.url])];
    const chaves = new Set(tentadas.map(chaveBusca));
    const destino = alternativas.find((url) => !chaves.has(chaveBusca(url)));
    if (!destino) throw new Error("Todas as " + alternativas.length + " buscas foram esgotadas para " + item.nome + ". " + motivo);
    item.buscasTentadas = tentadas;
    item.paginasResultados = [];
    item.proximaBusca = destino;
    lote.fase = "aguardando";
    lote.proxima = agora() + 5000;
    lote.mensagem = "Nova tentativa para " + item.nome + ": " + new URL(destino).searchParams.get("card")
      + (new URL(destino).searchParams.has("tipo") ? " (busca padrao)." : " (sem filtro de tipo).");
    await salvar(lote);
    await agendar(lote.proxima);
    return true;
  }

  async function processarInterno(abaEvento) {
    const lote = await carregar();
    if (!loteAberto(lote) || lote.fase === "pausado" || (abaEvento !== undefined && abaEvento !== lote.abaId)) return;
    try {
      if (!await api.permissions.contains({ origins: ORIGENS })) {
        await pausar(lote, "Permissao da LigaMagic removida. Autorize novamente ao retomar.", true);
        return;
      }
      if (lote.fase === "aguardando") {
        while (lote.fase === "aguardando") {
          const salva = await reutilizar(lote);
          if (!salva) break;
          await concluirItem(lote, salva, true);
        }
        if (lote.fase === "concluido") { await api.alarms.clear(ALARME_LOTE); return; }
        if (agora() < lote.proxima) {
          await agendar(lote.proxima);
          return;
        }
        if (lote.abaId === null) {
          const aba = await api.tabs.create({ url: "about:blank", active: true });
          lote.abaId = aba.id;
        } else await api.tabs.get(lote.abaId);
        const item = lote.itens[lote.indice];
        const url = new URL(item.proximaBusca || "https://www.ligamagic.com.br/");
        if (!item.proximaBusca) url.search = new URLSearchParams({ view: "cards/card", card: item.nome, tipo: "1" }).toString();
        delete item.proximaBusca;
        item.buscasTentadas = [...new Set([...(item.buscasTentadas || []), url.href])];
        lote.fase = "navegando";
        lote.prazo = agora() + 45000;
        lote.mensagem = "Abrindo " + lote.itens[lote.indice].nome + ".";
        await salvar(lote);
        await agendar(lote.prazo);
        await navegar(lote, url.href);
        return;
      }
      const aba = await api.tabs.get(lote.abaId);
      if (aba.status !== "complete" || !paginaConfere(aba, lote)) {
        if (agora() >= lote.prazo) {
          if (paginaConfere(aba, lote) && await tentarOutraBusca(lote, aba, "A pagina nao carregou como esperado.")) return;
          await pausar(lote, "A pagina nao carregou como esperado. Confira a aba da lista e retome, ou pule esta carta.", true);
        }
        else await agendar(lote.prazo);
        return;
      }
      lote.fase = "lendo";
      lote.metricas = lote.metricas || {};
      if (lote.inicioNavegacao !== undefined) {
        lote.metricas.navegacaoMs = (lote.metricas.navegacaoMs || 0) + Math.max(0, Date.now() - lote.inicioNavegacao);
        delete lote.inicioNavegacao;
      }
      await salvar(lote);
      const item = lote.itens[lote.indice];
      const basico = terrenoBasico(item.nome);
      const colecao = lote.colecao?.criterios === "variantes";
      const basicoColecao = colecao && basico;
      if (colecao || basico && !Array.isArray(item.referenciasTerreno)) {
        try {
          if (colecao) await confirmarColecao(item);
          else Object.assign(item, await resolverReferencia(item));
          await salvar(lote);
        } catch (erro) {
          const motivo = "Nao foi possivel confirmar a arte e a impressao da carta: " + erro.message;
          if (colecao) {
            await pendenciaColecao(lote, motivo);
            return;
          }
          throw new Error(motivo);
        }
      }
      const criterioTerreno = colecao ? item : { ...item, codigoEdicao: "", numero: "", acabamento: "normal" };
      let pagina;
      try {
        const resultados = await medir("leituraPagina", () => api.scripting.executeScript({ target: { tabId: lote.abaId }, func: capturarPagina,
          args: lote.colecao?.criterios === "variantes" ? [true, true] : [true] }));
        pagina = resultados[0]?.result;
        if (!pagina) throw new Error("Dados da carta indisponiveis.");
      } catch (erro) {
        if (await tentarOutraBusca(lote, aba, erro.message || "Falha ao ler a carta.")) return;
        throw erro;
      }
      if (pagina.erroCaptura) throw new Error(pagina.erroCaptura);
      if (Array.isArray(pagina.resultados)) {
        if ((basico || colecao) && Array.isArray(pagina.edicoesResultados)) {
          const visitadas = new Set([aba.url, ...(item.paginasResultados || []), ...(item.buscasTentadas || [])].map(chaveBusca));
          const disponiveis = pagina.edicoesResultados.filter((edicao) => {
            try {
              const url = new URL(edicao.url);
              return ORIGENS.some((origem) => url.origin + "/*" === origem) && url.searchParams.get("view") === "cards/card"
                && nomesEquivalentes(url.searchParams.get("card") || "", item.nome) && !visitadas.has(chaveBusca(url.href));
            } catch { return false; }
          });
          const [referencia, ...alternativas] = ordenarEdicoesColecao(disponiveis, criterioTerreno,
            basico && criterioTerreno.acabamento === "normal" ? 0 : null, colecao ? "m" : "p");
          if (!referencia) {
            const motivo = basico ? "Nenhuma versao de terreno compativel com " + (colecao ? "preco medio" : "preco minimo") + " disponivel."
              : "Impressao " + item.codigoEdicao.toUpperCase() + " #" + item.numero + " nao encontrada na galeria da LigaMagic.";
            if (colecao) {
              await pendenciaColecao(lote, motivo);
              return;
            }
            throw new Error(motivo);
          }
          item.paginasResultados = [...(item.paginasResultados || []), aba.url];
          item.alternativasTerreno = alternativas.slice(0, 20).map(({ code, num, url }) => ({ code, num, url }));
          await abrirReferenciaTerreno(lote, referencia);
          return;
        }
        const candidatos = new Map();
        const chaveNome = (nome) => normalizarNome(nome).replace(/\s*\/\/\s*/g, " // ");
        for (const resultado of pagina.resultados) {
          try {
            const url = new URL(resultado.url);
            const nome = url.searchParams.get("card") || "";
            if (!ORIGENS.some((origem) => url.origin + "/*" === origem) || url.searchParams.get("view") !== "cards/card"
                || !nomesEquivalentes(nome, item.nome)) continue;
            url.hash = "";
            candidatos.set(chaveNome(nome), url.href);
          } catch {}
        }
        const exato = candidatos.get(chaveNome(item.nome));
        if (!exato && candidatos.size > 1) {
          if (await tentarOutraBusca(lote, aba, "Mais de um resultado corresponde ao nome. Abra a carta correta e retome.")) return;
          throw new Error("Mais de um resultado corresponde a " + item.nome + ". Abra a carta correta na aba da lista e retome.");
        }
        const destino = exato || [...candidatos.values()][0];
        if (!destino) {
          if (await tentarOutraBusca(lote, aba, "Nenhum resultado exato. Confira o nome ou abra a carta correta e retome.")) return;
          throw new Error("Nenhum resultado exato para " + item.nome + ". Confira o nome ou abra a carta correta na aba da lista e retome.");
        }
        const visitadas = item.paginasResultados || [];
        const repetida = [aba.url, ...visitadas, ...(item.buscasTentadas || [])]
          .some((endereco) => chaveBusca(endereco) === chaveBusca(destino));
        if (repetida || visitadas.length >= 3) {
          if (await tentarOutraBusca(lote, aba, "A LigaMagic continua exibindo resultados. Abra a pagina de precos e retome.")) return;
          throw new Error("A LigaMagic continua exibindo resultados para " + item.nome + ". Abra a pagina de precos da carta na aba da lista e retome.");
        }
        item.paginasResultados = [...visitadas, aba.url];
        lote.fase = "navegando";
        lote.prazo = agora() + 45000;
        lote.mensagem = "Abrindo o resultado exato de " + item.nome + ".";
        await salvar(lote);
        await agendar(lote.prazo);
        await navegar(lote, destino);
        return;
      }
      let preco;
      try {
        conferirNome(pagina, item.nome);
        if (!basico && Array.isArray(pagina.edicoes) && pagina.edicoes.length) {
          const snapshot = { ...pagina, cachePrecoEm: agora() };
          paginasRecentes.set(chavePagina(item), snapshot);
          await cache.set(chavePagina(item), snapshot, VALIDADE.precos);
        }
        if (basico && !basicoColecao) {
          const [edicao] = ordenarEdicoesColecao(pagina.edicoes || [], criterioTerreno, 0, "p");
          if (!edicao) throw new Error("Nenhum preco verde positivo em uma versao comum compativel do terreno.");
          preco = { ...extrairCotacao({ ...pagina, edicoes: [edicao] }, true), codigoEdicao: edicao.code, numero: String(edicao.num),
            motivoReferencia: edicao.motivoReferencia };
        } else preco = lote.colecao?.criterios === "variantes" ? extrairCotacaoColecao(pagina, item) : extrairCotacao(pagina, lote.somenteNormal);
      } catch (erro) {
        if (basico && item.alternativasTerreno?.length) {
          item.paginasResultados = [...(item.paginasResultados || []), aba.url];
          const visitadas = new Set(item.paginasResultados.map(chaveBusca));
          while (item.alternativasTerreno.length) {
            const referencia = item.alternativasTerreno.shift();
            if (!visitadas.has(chaveBusca(referencia.url))) {
              await abrirReferenciaTerreno(lote, referencia);
              return;
            }
          }
        }
        if (lote.colecao?.criterios === "variantes") {
          await pendenciaColecao(lote, erro.message);
          return;
        }
        if (await tentarOutraBusca(lote, aba, erro.message)) return;
        throw erro;
      }
      const cotacao = await complementar(lote, preco);
      await concluirItem(lote, cotacao);
      if (lote.fase === "concluido") await api.alarms.clear(ALARME_LOTE);
      else await processarInterno();
    } catch (erro) {
      await pausar(await carregar(), erro.message || "Falha ao consultar a carta. Confira a aba da lista e retome.", true);
    }
  }

  async function comandoInterno(tipo, dados) {
    let lote = await carregar();
    if (tipo === "limpar") {
      if (loteAberto(lote) && lote.fase !== "pausado") throw new Error("Pause a consulta antes de limpar a lista.");
      await api.alarms.clear(ALARME_LOTE);
      paginasRecentes.clear();
      await api.storage.local.set({ lote: null, capturas: [], rascunhoLista: "", rascunhoColecao: null, rascunhoComandante: "" });
      return;
    }
    if (tipo === "iniciar") {
      if (loteAberto(lote)) throw new Error("Conclua ou cancele a lista atual antes de iniciar outra.");
      paginasRecentes.clear();
      const porVariante = dados.colecao?.criterios === "variantes";
      const deck = porVariante ? { itens: prepararColecao(dados.colecao.variantes), comandantes: [] }
        : prepararListaDeck(dados.texto, dados.colecao ? "" : dados.comandante, Boolean(dados.colecao || dados.importacaoCsv));
      const { itens, comandantes } = deck;
      const intervalo = Number(dados.intervalo ?? 4);
      if (!Number.isFinite(intervalo) || intervalo < 3 || intervalo > 300) throw new Error("Intervalo deve estar entre 3 e 300 segundos.");
      if (!await api.permissions.contains({ origins: ORIGENS })) throw new Error("Autorize o acesso a LigaMagic para consultar a lista.");
      lote = { itens, comandantes, indice: 0, fase: "aguardando", abaId: null, proxima: agora(), intervalo,
        somenteNormal: !porVariante && Boolean(dados.somenteNormal), mensagem: "Iniciando lista.", texto: dados.texto,
        atualizarPrecos: Boolean(dados.atualizarPrecos), metricas: { navegacoes: 0, cartasCache: 0 }, metricasIniciais: obterMetricas(),
        colecao: dados.colecao ? { arquivo: String(dados.colecao.arquivo || "Colecao CSV"),
          ...(porVariante ? { criterios: "variantes" } : {}),
          proxiesExcluidos: Number.isSafeInteger(dados.colecao.proxiesExcluidos) && dados.colecao.proxiesExcluidos > 0 ? dados.colecao.proxiesExcluidos : 0 } : null };
          await api.storage.local.set({ lote, capturas: [] });
      await processarInterno();
      return;
    }
    if (tipo === "retentar") {
      if (!lote || !["concluido", "cancelado", "pausado"].includes(lote.fase)) throw new Error("Pause ou conclua a lista antes de retentar as pendencias.");
      const indice = lote.itens.findIndex((item) => item.status !== "ok");
      if (indice === -1) throw new Error("Nenhuma pendencia para retentar.");
      if (!await api.permissions.contains({ origins: ORIGENS })) throw new Error("Autorize o acesso a LigaMagic para retentar as pendencias.");
      paginasRecentes.clear();
      for (const item of lote.itens.filter((entrada) => entrada.status !== "ok")) {
        item.status = "pendente";
        item.erro = "";
        delete item.buscasTentadas;
        delete item.paginasResultados;
        delete item.proximaBusca;
        delete item.alternativasTerreno;
      }
      if (lote.abaId !== null && !await api.tabs.get(lote.abaId).catch(() => null)) lote.abaId = null;
      lote.indice = indice;
      lote.fase = "aguardando";
      lote.proxima = agora();
      lote.mensagem = "Retentando apenas as pendencias.";
      await salvar(lote);
      await processarInterno();
      return;
    }
    if (!loteAberto(lote)) throw new Error("Nenhuma lista em andamento.");
    if (tipo === "pausar") {
      await pausar(lote, "Pausada por voce.");
    } else if (tipo === "cancelar") {
      lote.fase = "cancelado";
      lote.mensagem = "Lista cancelada. Resultados anteriores preservados.";
      await salvar(lote);
      await api.alarms.clear(ALARME_LOTE);
    } else if (tipo === "retomar") {
      if (lote.fase !== "pausado") throw new Error("A lista nao esta pausada.");
      const aba = lote.abaId === null ? null : await api.tabs.get(lote.abaId).catch(() => null);
      lote.itens[lote.indice].status = "pendente";
      lote.itens[lote.indice].erro = "";
      delete lote.itens[lote.indice].buscasTentadas;
      delete lote.itens[lote.indice].paginasResultados;
      delete lote.itens[lote.indice].proximaBusca;
      lote.fase = aba && paginaConfere(aba, lote) ? "navegando" : "aguardando";
      if (!aba) lote.abaId = null;
      lote.proxima = agora();
      lote.prazo = agora() + 45000;
      await salvar(lote);
      await processarInterno();
    } else if (tipo === "pular") {
      if (lote.fase !== "pausado") throw new Error("Pause a lista antes de pular uma carta.");
      const item = lote.itens[lote.indice];
      item.status = "pulado";
      item.erro = item.erro || "Pulado por voce.";
      avancar(lote);
      await salvar(lote);
      if (lote.fase === "concluido") await api.alarms.clear(ALARME_LOTE);
      else await agendar(lote.proxima);
    } else throw new Error("Comando de lista desconhecido.");
  }

  return {
    comando: (tipo, dados = {}) => serializar(() => comandoInterno(tipo, dados)),
    processar: (abaId) => serializar(() => processarInterno(abaId)),
    abaRemovida: (abaId) => serializar(async () => {
      const lote = await carregar();
      if (loteAberto(lote) && lote.abaId === abaId) {
        lote.abaId = null;
        await pausar(lote, "A aba da lista foi fechada. Retome para abrir uma nova aba.");
      }
    }),
    reiniciar: () => serializar(async () => {
      const lote = await carregar();
      if (loteAberto(lote) && lote.fase !== "pausado") await pausar(lote, "Navegador ou extensao reiniciado. Retome quando desejar.");
    }),
  };
}