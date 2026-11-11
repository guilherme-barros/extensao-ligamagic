const TRINTA_DIAS = 30 * 86400000;
export const VALIDADE = { precos: TRINTA_DIAS, metadados: TRINTA_DIAS, categorias: TRINTA_DIAS, imagens: TRINTA_DIAS };

export function criarFila(limite = 4) {
  let ativas = 0;
  const pendentes = [];
  const executar = () => {
    while (ativas < limite && pendentes.length) {
      const { acao, resolver, rejeitar } = pendentes.shift();
      ativas++;
      Promise.resolve().then(acao).then(resolver, rejeitar).finally(() => { ativas--; executar(); });
    }
  };
  return (acao) => new Promise((resolver, rejeitar) => { pendentes.push({ acao, resolver, rejeitar }); executar(); });
}

export function criarCachePersistente({ banco = globalThis.indexedDB, agora = Date.now, limiteBytes = 1024 * 1024 * 1024 } = {}) {
  const memoria = new Map();
  const gravar = criarFila(1);
  let abertura;
  const abrir = () => {
    if (!abertura) abertura = new Promise((resolver) => {
      if (!banco) return resolver(null);
      const pedido = banco.open("ligamagic-cache", 1);
      pedido.onupgradeneeded = () => {
        for (const nome of ["dados", "indice"]) if (!pedido.result.objectStoreNames.contains(nome)) pedido.result.createObjectStore(nome, { keyPath: "chave" });
      };
      pedido.onsuccess = () => resolver(pedido.result);
      pedido.onerror = pedido.onblocked = () => resolver(null);
    }).catch(() => null);
    return abertura;
  };
  const lerBanco = (db, nome, chave) => new Promise((resolver, rejeitar) => {
    const pedido = db.transaction(nome).objectStore(nome).get(chave);
    pedido.onsuccess = () => resolver(pedido.result);
    pedido.onerror = () => rejeitar(pedido.error);
  });
  const tamanho = (valor) => valor?.byteLength ?? (typeof Blob !== "undefined" && valor instanceof Blob ? valor.size
    : new TextEncoder().encode(JSON.stringify(valor)).length);
  return {
    async get(chave) {
      try {
        const db = await abrir();
        if (!db) {
          const entrada = memoria.get(chave);
          if (!entrada || entrada.expiraEm <= agora()) { memoria.delete(chave); return null; }
          entrada.usadoEm = agora();
          return structuredClone(entrada.valor);
        }
        const meta = await lerBanco(db, "indice", chave);
        if (!meta || meta.expiraEm <= agora()) return null;
        const entrada = await lerBanco(db, "dados", chave);
        if (!entrada) return null;
        const transacao = db.transaction("indice", "readwrite");
        transacao.objectStore("indice").put({ ...meta, usadoEm: agora() });
        return entrada.valor;
      } catch { return null; }
    },
    set(chave, valor, validade) {
      return gravar(async () => {
        try {
          const bytes = tamanho(valor);
          if (bytes > limiteBytes) return false;
          const meta = { chave, bytes, expiraEm: agora() + validade, usadoEm: agora() };
          const db = await abrir();
          if (!db) {
            memoria.set(chave, { ...meta, valor: structuredClone(valor) });
            let total = [...memoria.values()].reduce((soma, entrada) => soma + entrada.bytes, 0);
            for (const entrada of [...memoria.values()].sort((primeira, segunda) => primeira.usadoEm - segunda.usadoEm)) {
              if (entrada.expiraEm <= agora() || total > limiteBytes) { memoria.delete(entrada.chave); total -= entrada.bytes; }
            }
            return true;
          }
          return await new Promise((resolver) => {
            const transacao = db.transaction(["dados", "indice"], "readwrite");
            const indice = transacao.objectStore("indice");
            const dados = transacao.objectStore("dados");
            const pedido = indice.getAll();
            pedido.onsuccess = () => {
              const entradas = pedido.result.filter((entrada) => entrada.chave !== chave);
              let total = bytes + entradas.reduce((soma, entrada) => soma + entrada.bytes, 0);
              for (const entrada of entradas.sort((primeira, segunda) => primeira.usadoEm - segunda.usadoEm)) {
                if (entrada.expiraEm <= agora() || total > limiteBytes) {
                  indice.delete(entrada.chave); dados.delete(entrada.chave); total -= entrada.bytes;
                }
              }
              dados.put({ chave, valor }); indice.put(meta);
            };
            transacao.oncomplete = () => resolver(true);
            transacao.onerror = transacao.onabort = () => resolver(false);
          });
        } catch { return false; }
      });
    },
  };
}

export const cacheLocal = criarCachePersistente();
const metricas = {};
export function contar(nome, quantidade = 1) { metricas[nome] = (metricas[nome] || 0) + quantidade; }
export function obterMetricas(desde = {}) {
  return Object.fromEntries(Object.entries(metricas).map(([nome, valor]) => [nome, valor - (desde[nome] || 0)]));
}
export async function medir(nome, acao) {
  const inicio = performance.now();
  try { return await acao(); }
  finally { contar(nome + "Ms", Math.round(performance.now() - inicio)); contar(nome + "Chamadas"); }
}

const filaApi = criarFila(1);
export async function consultarJsonScryfall(url, requisitar = fetch, aguardar = (ms) => new Promise((resolver) => setTimeout(resolver, ms))) {
  const cache = requisitar === globalThis.fetch ? cacheLocal : null;
  const chave = "scryfall:" + url;
  const salvo = await cache?.get(chave);
  if (salvo) { contar("metadadosCache"); return salvo; }
  return filaApi(async () => {
    const repetido = await cache?.get(chave);
    if (repetido) { contar("metadadosCache"); return repetido; }
    for (let tentativa = 0; tentativa < 3; tentativa++) {
      await aguardar(110);
      const resposta = await medir("scryfall", () => requisitar(url, { signal: AbortSignal.timeout(15000), headers: { Accept: "application/json" } }));
      if ([429, 503].includes(resposta.status) && tentativa < 2) {
        const segundos = Number(resposta.headers?.get("Retry-After"));
        await aguardar(Number.isFinite(segundos) && segundos > 0 ? Math.min(segundos * 1000, 60000) : 500 * 2 ** tentativa);
        continue;
      }
      if (!resposta.ok) throw new Error("Scryfall indisponivel (HTTP " + resposta.status + ").");
      const dados = await resposta.json();
      await cache?.set(chave, dados, VALIDADE.metadados);
      if (dados.name && dados.set && dados.collector_number) {
        await cache?.set("scryfall:https://api.scryfall.com/cards/" + dados.set + "/" + dados.collector_number, dados, VALIDADE.metadados);
        if (dados.id) await cache?.set("scryfall:https://api.scryfall.com/cards/" + dados.id, dados, VALIDADE.metadados);
      }
      return dados;
    }
  });
}