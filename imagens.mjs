import { nomesEquivalentes } from "./core.mjs";
import { cacheLocal, VALIDADE, criarFila, contar, medir, consultarJsonScryfall } from "./cache.mjs";

export const ORIGENS_IMAGENS = ["https://api.scryfall.com/*", "https://cards.scryfall.io/*"];

const filaImagens = criarFila(4);

export function criarResolvedorImagens(requisitar = fetch, aguardar = (ms = 110) => new Promise((resolver) => setTimeout(resolver, ms)), opcoes = {}) {
  const cache = new Map();
  const persistente = opcoes.cache ?? (requisitar === globalThis.fetch ? cacheLocal : null);
  const numero = (valor) => String(valor).toLowerCase().replace(/^0+(?=\d)/, "");
  const obter = async (url, imagem = false) => {
    if (!cache.has(url)) cache.set(url, (async () => {
      if (!imagem) return consultarJsonScryfall(url, requisitar, aguardar);
      const salva = await persistente?.get("imagem:" + url);
      if (salva) contar("imagensCache");
      const buffer = salva || await filaImagens(() => medir("downloadImagem", async () => {
        for (let tentativa = 0; tentativa < 3; tentativa++) {
          const resposta = await requisitar(url, { signal: AbortSignal.timeout(15000), headers: { Accept: "image/jpeg,image/png" } });
          if ([429, 503].includes(resposta.status) && tentativa < 2) {
            const segundos = Number(resposta.headers?.get("Retry-After"));
            await aguardar(Number.isFinite(segundos) && segundos > 0 ? Math.min(segundos * 1000, 60000) : 500 * 2 ** tentativa);
            continue;
          }
          if (!resposta.ok) throw new Error("Scryfall: consulta de imagem indisponivel (HTTP " + resposta.status + ").");
          if (Number(resposta.headers?.get("content-length")) > 10000000) throw new Error("Imagem acima do limite de tamanho.");
          return new Uint8Array(await resposta.arrayBuffer());
        }
      }));
      if (buffer.length > 10000000) throw new Error("Imagem acima do limite de tamanho.");
      const extension = buffer[0] === 255 && buffer[1] === 216 && buffer[2] === 255 ? "jpeg"
        : buffer[0] === 137 && buffer[1] === 80 && buffer[2] === 78 && buffer[3] === 71 ? "png" : null;
      if (!extension) throw new Error("O arquivo recebido nao e uma imagem JPEG ou PNG.");
      if (!salva) await persistente?.set("imagem:" + url, buffer, VALIDADE.imagens);
      return { buffer, extension };
    })());
    try { return await cache.get(url); }
    catch (erro) { cache.delete(url); throw erro; }
  };
  const resolver = async (item, alternativa) => {
    const exata = Boolean(item.codigoEdicao && item.numero);
    const caminho = exata ? "/cards/" + encodeURIComponent(item.codigoEdicao.toLowerCase()) + "/" + encodeURIComponent(numero(item.numero))
      : /^[\da-f-]{36}$/i.test(item.scryfallId || "") ? "/cards/" + item.scryfallId
        : "/cards/named?" + new URLSearchParams({ exact: item.nome });
    const carta = alternativa || await obter("https://api.scryfall.com" + caminho);
    if (![carta.name, carta.printed_name].filter(Boolean).some((nome) => nomesEquivalentes(nome, item.nome))) {
      throw new Error("A imagem retornada nao corresponde ao nome da carta.");
    }
    if (exata && (carta.set?.toLowerCase() !== item.codigoEdicao.toLowerCase() || numero(carta.collector_number) !== numero(item.numero))) {
      throw new Error("A imagem retornada nao corresponde a edicao e ao numero da carta.");
    }
    const faces = [];
    for (const face of carta.image_uris ? [carta] : carta.card_faces || []) {
      try {
        const url = new URL(face.image_uris?.normal || face.image_uris?.large || face.image_uris?.png);
        if (url.origin !== "https://cards.scryfall.io") throw new Error("Origem da imagem nao permitida.");
        faces.push({ nome: face.name || carta.name, url: url.href, ...await obter(url.href, true) });
      } catch (erro) {
        faces.push({ nome: face.name || carta.name, erro: erro.message });
      }
    }
    if (!faces.length) throw new Error("Imagem indisponivel para esta carta.");
    return { referencia: (carta.set || "").toUpperCase() + " #" + (carta.collector_number || ""),
      porNome: !exata && !item.scryfallId, faces };
  };
  return async (item) => {
    const ilustrar = opcoes.permitirIlustracao && item.codigoEdicao && item.numero;
    try {
      const resultado = await resolver(item);
      if (ilustrar && !resultado.faces.some((face) => face.buffer)) throw new Error("Imagem da impressao exata indisponivel.");
      return resultado;
    } catch (erro) {
      if (!ilustrar) throw erro;
      const consulta = new URLSearchParams({ q: '!"' + item.nome + '" game:paper -is:fullart -is:showcase -is:borderless -is:extendedart',
        unique: "prints", order: "released", dir: "desc" });
      const catalogo = await obter("https://api.scryfall.com/cards/search?" + consulta);
      const carta = catalogo.data?.find((valor) => nomesEquivalentes(valor.name || "", item.nome) && !valor.digital
        && valor.finishes?.includes("nonfoil") && !valor.full_art && valor.border_color !== "borderless"
        && !valor.frame_effects?.some((efeito) => ["showcase", "extendedart"].includes(efeito)));
      if (!carta) throw erro;
      const resultado = await resolver({ nome: item.nome }, carta);
      if (!resultado.faces.some((face) => face.buffer)) throw erro;
      return { ...resultado, porNome: false, ilustrativa: true,
        referenciaOriginal: (item.edicao ? item.edicao + " / " : "") + item.codigoEdicao.toUpperCase() + " #" + item.numero
          + (item.acabamento ? " / " + item.acabamento : "") };
    }
  };
}