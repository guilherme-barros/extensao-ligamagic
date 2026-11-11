import { nomesEquivalentes } from "./core.mjs";
import { cacheLocal, VALIDADE, contar, medir, consultarJsonScryfall } from "./cache.mjs";

export const ORIGENS_TAGGER = ["https://tagger.scryfall.com/*", "https://api.scryfall.com/*"];

const categorias = [
  ["Removal", "removal"],
  ["Permanent Removal", "removal-permanent"],
  ["Planeswalker Removal", "removal-planeswalker"],
  ["Land Removal", "removal-land"],
  ["Creature Removal", "removal-creature"],
  ["Artifact Removal", "removal-artifact"],
  ["Enchantment Removal", "removal-enchantment"],
  ["Destroy Removal", "removal-destroy"],
  ["Exile Removal", "removal-exile"],
  ["Destroy Permanent", "removal-permanent-destroy"],
  ["Exile Permanent", "removal-permanent-exile"],
  ["Destroy Planeswalker", "removal-planeswalker-destroy"],
  ["Exile Planeswalker", "removal-planeswalker-exile"],
  ["Destroy Land", "removal-land-destroy"],
  ["Exile Land", "removal-land-exile"],
  ["Destroy Creature", "removal-creature-destroy", "destroy-creature"],
  ["Exile Creature", "removal-creature-exile", "exile-creature"],
  ["Board Wipe", "board-wipe", "mass-removal", "wrath", "sweeper"],
  ["Bounce to Hand", "bounce", "bounce-to-hand", "removal-bounce"],
  ["Tuck Removal", "removal-tuck"],
  ["Counter Spell", "counterspell"],
  ["-X/-X Debuff", "debuff", "debuff-creature", "toughness-reduction"],
  ["Damage Removal", "removal-damage", "removal-creature-damage"],
  ["Force Sacrifice", "edict", "forced-sacrifice", "force-sacrifice", "removal-sacrifice"],
  ["Destroy Artifact", "removal-artifact-destroy", "destroy-artifact"],
  ["Exile Artifact", "removal-artifact-exile"],
  ["Destroy Enchantment", "removal-enchantment-destroy", "destroy-enchantment"],
  ["Exile Enchantment", "removal-enchantment-exile"],
  ["Destroy Artifact or Enchantment", "disenchant", "naturalize", "disenchant/naturalize", "disenchant-naturalize"],
  ["Draw Cards", "draw", "card-draw"],
  ["Card Advantage", "card-advantage"],
  ["Hand Size Increase", "hand-size-increase"],
  ["Scry", "scry"],
  ["Surveil", "surveil"],
  ["Tutor / Search Library", "tutor", "search-library"],
  ["Impulse Draw", "impulse-draw"],
  ["Mill", "mill"],
  ["Mill Opponent", "mill-opponent", "opponent-mill"],
  ["Investigate", "investigate"],
  ["Cascade", "cascade"],
  ["Ramp", "ramp"],
  ["Adds Mana", "mana-producer", "adds-mana"],
  ["Ritual", "ritual"],
  ["Mana Multiplier", "mana-multiplier", "mana-doubler", "mana-tripler"],
  ["Untap Lands", "untap-land", "untap-lands"],
  ["Search for Lands", "tutor-land", "search-for-lands"],
  ["Mana Dork", "mana-dork"],
  ["Mana Rock", "mana-rock"],
  ["Cost Reduction", "cost-reduction", "cost-reducer"],
  ["Treasure Tokens", "token-treasure", "treasure-token", "creates-treasure", "repeatable-treasures"],
  ["Extra Land Drops", "extra-land", "extra-land-drop", "extra-land-drops", "additional-land-play"],
  ["Put Land onto Battlefield", "tutor-land-to-battlefield", "land-to-battlefield", "land-ramp"],
  ["Reanimate", "reanimate", "reanimation", "recursion-to-battlefield"],
  ["Self-Reanimate", "self-reanimate", "reanimate-self"],
  ["Return to Hand", "regrowth", "recursion-to-hand", "graveyard-to-hand"],
  ["Recursion", "recursion", "recursion-self", "recursion-any", "recursion-land", "restock", "restock-self"],
  ["Flashback", "flashback"],
  ["Unearth", "unearth"],
  ["Escape", "escape"],
  ["Delve", "delve"],
  ["Dredge", "dredge"],
  ["Self-Mill", "self-mill", "mill-self"],
  ["Graveyard Hate", "graveyard-hate"],
  ["Create Tokens", "repeatable-token-generator"],
  ["Create Creature Tokens", "token-creature", "creature-token", "creates-creature-tokens", "repeatable-creature-tokens", "repeatable-creature-token"],
  ["Create Noncreature Tokens", "repeatable-noncreature-tokens"],
  ["Create Artifact Tokens", "repeatable-artifact-tokens"],
  ["Zombies", "token-zombie", "zombie-token"],
  ["Soldiers", "token-soldier", "soldier-token"],
  ["Spirits", "token-spirit", "spirit-token"],
  ["Goblins", "token-goblin", "goblin-token"],
  ["Saprolings", "token-saproling", "saproling-token"],
  ["Angels", "token-angel", "angel-token"],
  ["Rats", "token-rat", "rat-token"],
  ["Elves", "token-elf", "elf-token"],
  ["Vampires", "token-vampire", "vampire-token"],
  ["Insects", "token-insect", "insect-token"],
  ["Birds / Drakes", "token-bird", "bird-token", "token-drake", "drake-token"],
  ["Beasts", "token-beast", "beast-token"],
  ["Cats", "token-cat", "cat-token"],
  ["Dogs", "token-dog", "dog-token"],
  ["Humans", "token-human", "human-token"],
  ["Knights", "token-knight", "knight-token"],
  ["Warriors", "token-warrior", "warrior-token"],
  ["Squirrels", "token-squirrel", "squirrel-token"],
  ["Thopters", "token-thopter", "thopter-token"],
  ["Elementals", "token-elemental", "elemental-token"],
  ["Faeries", "token-faerie", "faerie-token"],
  ["Demons", "token-demon", "demon-token"],
  ["Dragons", "token-dragon", "dragon-token"],
  ["Wolves", "token-wolf", "wolf-token"],
  ["Food Tokens", "token-food", "food-token", "creates-food"],
  ["Clue Tokens", "token-clue", "clue-token", "creates-clue"],
  ["Blood Tokens", "token-blood", "blood-token", "creates-blood"],
  ["Map Tokens", "token-map", "map-token", "creates-map"],
  ["+1/+1 Counters", "+1/+1-counters", "+1/+1-counter", "counter-+1/+1", "plus-one-plus-one-counters"],
  ["-1/-1 Counters", "-1/-1-counters", "-1/-1-counter", "counter--1/-1", "minus-one-minus-one-counters"],
  ["Proliferate", "proliferate"],
  ["Poison / Toxic", "poison", "poison-counter", "toxic", "infect"],
  ["Landfall", "landfall"],
  ["ETB Effects", "etb", "etb-trigger", "enter-the-battlefield", "enters-the-battlefield"],
  ["Death Triggers", "death-trigger", "dies-trigger"],
  ["Self Death Triggers", "death-trigger-self", "life-and-death-trigger-self"],
  ["Attack Triggers", "attack-trigger"],
  ["Upkeep Triggers", "upkeep-trigger"],
  ["End Step Triggers", "end-step-trigger"],
  ["Damage Triggers", "damage-trigger", "combat-damage-trigger"],
  ["Cast Triggers", "cast-trigger"],
  ["Leave Battlefield", "ltb", "leaves-the-battlefield", "leave-the-battlefield"],
  ["Trigger Doubler", "trigger-doubler"],
  ["Sacrifice Outlet", "sacrifice-outlet"],
  ["Blink / Flicker", "blink", "flicker"],
  ["Copy Effects", "copy", "copy-effect", "clone"],
  ["Steal / Gain Control", "theft", "steal", "gain-control"],
  ["Life Gain", "lifegain", "life-gain"],
  ["Life Drain", "drain", "drain-life", "life-drain", "blood-artist-ability"],
  ["Group Slug / Direct Damage", "group-slug", "opponent-loses-life", "burn", "direct-damage"],
  ["Discard", "discard"],
  ["Equipment", "equipment"],
  ["Auras", "aura"],
  ["Vehicles", "vehicle"],
  ["Unblockable", "unblockable"],
  ["Evasion", "evasion"],
  ["Tribal Lord / Anthem", "lord", "tribal-lord", "anthem"],
  ["Untap", "untap", "untapper"],
  ["Phase Manipulation", "phase-manipulation"],
  ["Extra Turns", "extra-turn", "extra-turns"],
  ["Extra Combat", "extra-combat", "additional-combat"],
  ["Alternate Win Condition", "alternate-win-condition"],
  ["Populate", "populate"],
  ["Amass", "amass"],
  ["Gives Haste", "gives-haste"],
  ["Gives Vigilance", "gives-vigilance"],
  ["Gives Flying", "gives-flying"],
  ["Gives Trample", "gives-trample"],
  ["Gives Deathtouch", "gives-deathtouch"],
  ["Gives Lifelink", "gives-lifelink"],
  ["Protection", "gives-shroud", "gives-protection", "gives-indestructible", "gives-hexproof"],
  ["Pump / Buff (+X/+X)", "pump", "buff", "pump-creature"],
];

export const CATEGORIAS_PERMITIDAS = categorias.map(([nome]) => nome);

const categoriasTokens = new Set(categorias.filter(([, ...aliases]) => aliases.some((alias) => alias.startsWith("token-")
  || alias === "repeatable-token-generator" || alias.startsWith("repeatable-") && alias.endsWith("-tokens"))).map(([nome]) => nome));
const tokensNaoCriatura = ["Treasure Tokens", "Food Tokens", "Clue Tokens", "Blood Tokens", "Map Tokens"];
const tokensCriatura = [...categoriasTokens].filter((nome) => !tokensNaoCriatura.includes(nome)
  && !["Create Tokens", "Create Creature Tokens", "Create Noncreature Tokens", "Create Artifact Tokens"].includes(nome));
const categoriasGerais = new Map([
  ["Removal", ["Permanent Removal", "Planeswalker Removal", "Land Removal", "Destroy Creature", "Exile Creature", "Board Wipe",
    "Bounce to Hand", "Tuck Removal", "-X/-X Debuff", "Damage Removal", "Force Sacrifice", "Destroy Artifact", "Destroy Enchantment", "Destroy Artifact or Enchantment",
    "Creature Removal", "Artifact Removal", "Enchantment Removal", "Destroy Removal", "Exile Removal", "Destroy Permanent", "Exile Permanent",
    "Destroy Planeswalker", "Exile Planeswalker", "Destroy Land", "Exile Land", "Exile Artifact", "Exile Enchantment"]],
  ["Permanent Removal", ["Planeswalker Removal", "Land Removal", "Creature Removal", "Artifact Removal", "Enchantment Removal", "Destroy Permanent", "Exile Permanent",
    "Destroy Planeswalker", "Exile Planeswalker", "Destroy Land", "Exile Land", "Destroy Creature", "Exile Creature", "Destroy Artifact", "Exile Artifact",
    "Destroy Enchantment", "Exile Enchantment", "Destroy Artifact or Enchantment"]],
  ["Creature Removal", ["Destroy Creature", "Exile Creature", "Damage Removal", "-X/-X Debuff"]],
  ["Artifact Removal", ["Destroy Artifact", "Exile Artifact", "Destroy Artifact or Enchantment"]],
  ["Enchantment Removal", ["Destroy Enchantment", "Exile Enchantment", "Destroy Artifact or Enchantment"]],
  ["Planeswalker Removal", ["Destroy Planeswalker", "Exile Planeswalker"]],
  ["Land Removal", ["Destroy Land", "Exile Land"]],
  ["Destroy Removal", ["Destroy Permanent", "Destroy Planeswalker", "Destroy Land", "Destroy Creature", "Destroy Artifact", "Destroy Enchantment", "Destroy Artifact or Enchantment"]],
  ["Exile Removal", ["Exile Permanent", "Exile Planeswalker", "Exile Land", "Exile Creature", "Exile Artifact", "Exile Enchantment"]],
  ["Card Advantage", ["Draw Cards", "Impulse Draw", "Tutor / Search Library", "Return to Hand"]],
  ["Mill", ["Self-Mill", "Mill Opponent"]],
  ["Recursion", ["Reanimate", "Self-Reanimate", "Return to Hand", "Flashback", "Unearth", "Escape", "Dredge"]],
  ["Reanimate", ["Self-Reanimate"]],
  ["Life Gain", ["Life Drain"]],
  ["Death Triggers", ["Self Death Triggers"]],
  ["Evasion", ["Unblockable"]],
  ["Untap", ["Untap Lands"]],
  ["Ramp", ["Adds Mana", "Ritual", "Mana Multiplier", "Untap Lands", "Search for Lands", "Mana Dork", "Mana Rock",
    "Cost Reduction", "Treasure Tokens", "Extra Land Drops", "Put Land onto Battlefield"]],
  ["Adds Mana", ["Mana Dork", "Mana Rock", "Ritual", "Mana Multiplier"]],
  ["Search for Lands", ["Put Land onto Battlefield"]],
  ["Create Tokens", [...categoriasTokens].filter((nome) => nome !== "Create Tokens")],
  ["Create Creature Tokens", tokensCriatura],
  ["Create Noncreature Tokens", tokensNaoCriatura],
  ["Create Artifact Tokens", ["Thopters", ...tokensNaoCriatura]],
]);

function chaveTag(nome) {
  return nome.trim().toLowerCase().replace(/[\s_]+/g, "-");
}

export function permiteSacrificioBarato(texto = "", nomeCarta = "") {
  return texto.split("\n").some((linha) => {
    const separador = linha.indexOf(":");
    if (separador < 0 || /\b(?:only (?:once|twice|\d+ times?)|no more than)\b/i.test(linha)) return false;
    let custo = linha.slice(0, separador).trim();
    const sacrificio = /\bsacrifice\s+([^,]+)/i.exec(custo);
    if (!sacrificio || !/\b(?:creature|artifact|enchantment|permanent|land|token|goblin|elf|zombie|food|treasure|clue|blood)\b/i.test(sacrificio[1])) return false;
    if (/\b(?:this|itself|two|three|four|five|six|seven|eight|nine|ten|[2-9])\b/i.test(sacrificio[1])) return false;
    if (nomeCarta && sacrificio[1].toLowerCase().includes(nomeCarta.toLowerCase())) return false;
    custo = custo.replace(sacrificio[0], "");
    let total = 0;
    let valido = true;
    custo = custo.replace(/\bpay (\d+) life\b/gi, (_, vida) => {
      total += Number(vida);
      return "";
    }).replace(/\{([^}]+)\}/g, (_, simbolo) => {
      if (/^\d+$/.test(simbolo)) total += Number(simbolo);
      else if (/^[WUBRGC](?:\/[WUBRGC])?$/.test(simbolo)) total++;
      else valido = false;
      return "";
    });
    return valido && total <= 1 && !custo.replace(/[,\s]/g, "");
  });
}

export function filtrarCategorias(tags, relacoes = [], contexto = {}) {
  const todosNomes = new Set(tags.map((tag) => chaveTag(tag)));
  const categoriasCarta = [...categorias];
  const remocoesDinamicas = new Set();
  const aliasesConhecidos = new Set(categorias.flatMap((categoria) => categoria.map(chaveTag)));
  for (const tag of todosNomes) {
    const categoriaContagem = /^(?:.+-)?count-matters(?:-.+)?$/.test(tag);
    if ((!/^(?:gives|removal)-.+/.test(tag) && !categoriaContagem) || aliasesConhecidos.has(tag)) continue;
    const nome = tag.split("-").map((palavra) => palavra.charAt(0).toUpperCase() + palavra.slice(1)).join(" ");
    categoriasCarta.push([nome, tag]);
    if (tag.startsWith("removal-")) remocoesDinamicas.add(nome);
  }
  const nomes = new Set(todosNomes);
  const secundarios = new Set();
  for (const { tag, ancestrais } of relacoes) {
    if (categoriasCarta.some(([nome, ...aliases]) => [nome, ...aliases].some((alias) => chaveTag(alias) === chaveTag(tag)))) {
      for (const ancestral of ancestrais) {
        if (chaveTag(ancestral) !== chaveTag(tag)) secundarios.add(chaveTag(ancestral));
      }
    }
  }
  for (const secundario of secundarios) nomes.delete(secundario);
  const combinacoes = new Map([
    ["Destroy Creature", ["removal-creature", "removal-destroy"]],
    ["Exile Creature", ["removal-creature", "removal-exile"]],
    ["Destroy Artifact", ["removal-artifact", "removal-destroy"]],
    ["Destroy Enchantment", ["removal-enchantment", "removal-destroy"]],
    ["Exile Artifact", ["removal-artifact", "removal-exile"]],
    ["Exile Enchantment", ["removal-enchantment", "removal-exile"]],
    ["Destroy Permanent", ["removal-permanent", "removal-destroy"]],
    ["Exile Permanent", ["removal-permanent", "removal-exile"]],
    ["Destroy Planeswalker", ["removal-planeswalker", "removal-destroy"]],
    ["Exile Planeswalker", ["removal-planeswalker", "removal-exile"]],
    ["Destroy Land", ["removal-land", "removal-destroy"]],
    ["Exile Land", ["removal-land", "removal-exile"]],
  ]);
  const selecionadas = categoriasCarta.filter(([nome, ...aliases]) => {
    if (nome === "Life Gain" && todosNomes.has("opponent-lifegain")) return false;
    if (nome === "Sacrifice Outlet" && !permiteSacrificioBarato(contexto.textoOracle, contexto.nomeCarta)) return false;
    const candidatos = nome === "Force Sacrifice" ? todosNomes : nomes;
    return [nome, ...aliases].some((alias) => candidatos.has(chaveTag(alias)))
      || combinacoes.has(nome) && combinacoes.get(nome).every((tag) => nomes.has(tag));
  })
    .map(([nome]) => nome);
  const usosTags = new Map(Object.entries(contexto.usosTags || {}).filter(([, usos]) => Number.isSafeInteger(usos) && usos >= 0)
    .map(([tag, usos]) => [chaveTag(tag), usos]));
  const usosCategorias = new Map(categoriasCarta.map(([nome, ...aliases]) => {
    const usos = [nome, ...aliases].filter((alias) => todosNomes.has(chaveTag(alias))).map((alias) => usosTags.get(chaveTag(alias)) || 0);
    const componentes = combinacoes.get(nome);
    if (componentes?.every((tag) => nomes.has(tag))) usos.push(Math.min(...componentes.map((tag) => usosTags.get(tag) || 0)));
    return [nome, Math.max(0, ...usos)];
  }));
  return selecionadas.filter((nome) => contexto.preservarRedundantes || (!(categoriasGerais.get(nome) || []).some((especifica) => selecionadas.includes(especifica))
    && (nome !== "Removal" || !selecionadas.some((especifica) => remocoesDinamicas.has(especifica)))))
    .sort((primeiro, segundo) => usosCategorias.get(segundo) - usosCategorias.get(primeiro)
      || Number(categoriasTokens.has(segundo)) - Number(categoriasTokens.has(primeiro)));
}

const consulta = `query CategoriasCarta($set: String!, $number: String!, $twoSided: Boolean!) {
  front: cardBySet(set: $set, number: $number) {
    name oracleId taggings { status tag { name slug namespace status taggingCount ancestorTags {
      name slug namespace status taggingCount ancestorTags { name slug namespace status taggingCount }
    } } }
  }
  back: cardBySet(set: $set, number: $number, back: true) @include(if: $twoSided) {
    name oracleId taggings { status tag { name slug namespace status taggingCount ancestorTags {
      name slug namespace status taggingCount ancestorTags { name slug namespace status taggingCount }
    } } }
  }
}`;

async function obterResposta(url, opcoes, requisitar) {
  let resposta;
  try {
    resposta = await requisitar(url, { signal: AbortSignal.timeout(15000), ...opcoes });
  } catch (erro) {
    const origem = new URL(url).hostname;
    if (erro.name === "TimeoutError" || erro.name === "AbortError") {
      throw new Error(origem + " demorou a responder. Tente novamente mais tarde.");
    }
    throw new Error("Falha de conexao com " + origem + ". Verifique a rede, os bloqueadores e o certificado HTTPS do site. Detalhe: " + (erro.message || "erro de rede"));
  }
  if (!resposta.ok) {
    const origem = new URL(url).hostname;
    throw new Error(origem + " retornou HTTP " + resposta.status + (resposta.status === 429 ? ". Limite de consultas atingido; tente mais tarde." : "."));
  }
  return resposta;
}

export async function lerTagsNaPagina(query, variables) {
  try {
    const token = document.querySelector('meta[name="csrf-token"]')?.content;
    if (!token) throw new Error("Sessao do Scryfall Tagger indisponivel. Abra o site e conclua a verificacao, se solicitada.");
    let resposta;
    try {
      resposta = await fetch("/graphql", { method: "POST", credentials: "same-origin", signal: AbortSignal.timeout(15000),
        headers: { Accept: "application/json", "Content-Type": "application/json", "X-CSRF-Token": token },
        body: JSON.stringify({ query, variables }) });
    } catch (erro) {
      if (erro.name === "TimeoutError" || erro.name === "AbortError") {
        throw new Error("tagger.scryfall.com demorou a responder. Tente novamente mais tarde.");
      }
      throw new Error("Falha de conexao com tagger.scryfall.com. Abra o site e verifique a rede, os bloqueadores e o certificado HTTPS. Detalhe: " + (erro.message || "erro de rede"));
    }
    if (!resposta.ok) throw new Error("Scryfall Tagger retornou HTTP " + resposta.status + ".");
    return await resposta.json();
  } catch (erro) {
    return { erroTagger: erro.message || "Falha ao consultar tagger.scryfall.com." };
  }
}

export async function consultarTagsNavegador(url, variables, api = chrome) {
  const aba = await api.tabs.create({ url, active: false });
  try {
    await new Promise((resolver, rejeitar) => {
      let concluido = false;
      const finalizar = (erro) => {
        if (concluido) return;
        concluido = true;
        clearTimeout(prazo);
        api.tabs.onUpdated.removeListener(atualizar);
        api.tabs.onRemoved.removeListener(remover);
        if (erro) rejeitar(erro);
        else resolver();
      };
      const atualizar = (id, mudanca) => { if (id === aba.id && mudanca.status === "complete") finalizar(); };
      const remover = (id) => { if (id === aba.id) finalizar(new Error("A aba auxiliar do Scryfall Tagger foi fechada.")); };
      const prazo = setTimeout(() => finalizar(new Error("O Scryfall Tagger demorou a carregar. Abra o site e conclua a verificacao, se solicitada.")), 30000);
      api.tabs.onUpdated.addListener(atualizar);
      api.tabs.onRemoved.addListener(remover);
      api.tabs.get(aba.id).then((atual) => { if (atual.status === "complete") finalizar(); }, finalizar);
    });
    const resultados = await api.scripting.executeScript({ target: { tabId: aba.id }, world: "MAIN", func: lerTagsNaPagina, args: [consulta, variables] });
    const resultado = resultados[0]?.result;
    if (resultado?.erroTagger) throw new Error(resultado.erroTagger);
    if (!resultado) throw new Error("Nao foi possivel ler as tags do Scryfall Tagger. Abra o site e conclua a verificacao, se solicitada.");
    return resultado;
  } finally {
    await api.tabs.remove(aba.id).catch(() => {});
  }
}

export async function consultarTagger(nome, requisitar = fetch, consultarTags = consultarTagsNavegador, opcoes = {}) {
  const urlScryfall = new URL("https://api.scryfall.com/cards/named");
  urlScryfall.searchParams.set("exact", nome.split(/\s*\/\/\s*/)[0].trim());
  const cache = opcoes.cache ?? (requisitar === globalThis.fetch && consultarTags === consultarTagsNavegador ? cacheLocal : null);
  const carta = requisitar === globalThis.fetch ? await consultarJsonScryfall(urlScryfall.href, requisitar)
    : await (await obterResposta(urlScryfall.href, { credentials: "omit",
      headers: { Accept: "application/json", "User-Agent": "LigaMagicCapturaLocal/1.1" } }, requisitar)).json();
  if (typeof carta.name !== "string" || !nomesEquivalentes(carta.name, nome)
      || !/^[a-z0-9]+$/.test(carta.set || "") || typeof carta.collector_number !== "string" || !carta.oracle_id) {
    throw new Error("Carta exata nao encontrada no Scryfall: " + nome + ".");
  }
  const chaveCache = "categorias:1:" + carta.oracle_id;
  const salvo = await cache?.get(chaveCache);
  if (salvo) { contar("categoriasCache"); return salvo; }
  const urlTagger = "https://tagger.scryfall.com/card/" + carta.set + "/" + encodeURIComponent(carta.collector_number);
  const duasFaces = ["transform", "modal_dfc", "double_faced_token", "reversible_card"].includes(carta.layout);
  const resultado = await medir("tagger", () => consultarTags(urlTagger, { set: carta.set, number: carta.collector_number,
    twoSided: duasFaces }));
  if (resultado.errors?.length) throw new Error("Scryfall Tagger: " + (resultado.errors[0].message || "Erro na consulta de tags."));
  const frente = resultado.data?.front;
  if (frente?.oracleId !== carta.oracle_id || !Array.isArray(frente?.taggings)) {
    throw new Error("As tags retornadas pelo Scryfall Tagger nao correspondem a carta solicitada.");
  }
  if (duasFaces && (resultado.data.back?.oracleId !== carta.oracle_id || !Array.isArray(resultado.data.back?.taggings))) {
    throw new Error("Nao foi possivel validar as tags do verso da carta no Scryfall Tagger.");
  }
  const tags = new Set();
  const diretas = new Set();
  const secundarias = new Set();
  const relacoes = new Map();
  const usosTags = {};
  function registrarTag(tag) {
    const nome = tag.slug || tag.name;
    tags.add(nome);
    if (Number.isSafeInteger(tag.taggingCount) && tag.taggingCount >= 0) {
      const chave = chaveTag(nome);
      usosTags[chave] = Math.max(usosTags[chave] || 0, tag.taggingCount);
    }
    const ancestrais = (tag.ancestorTags || []).filter((ancestral) => ancestral.namespace === "card" && ancestral.status === "GOOD_STANDING");
    relacoes.set(nome, [...new Set([...(relacoes.get(nome) || []), ...ancestrais.map((ancestral) => ancestral.slug || ancestral.name)])]);
    for (const ancestral of ancestrais) {
      secundarias.add(ancestral.slug || ancestral.name);
      registrarTag(ancestral);
    }
  }
  for (const face of [frente, resultado.data.back].filter(Boolean)) {
    if (!Array.isArray(face.taggings)) throw new Error("Resposta de tags invalida do Scryfall Tagger.");
    for (const marcacao of face.taggings) {
      const tag = marcacao.tag;
      if (marcacao.status !== "GOOD_STANDING" || tag?.status !== "GOOD_STANDING" || tag.namespace !== "card") continue;
      diretas.add(tag.slug || tag.name);
      registrarTag(tag);
    }
  }
  const faces = Array.isArray(carta.card_faces) ? carta.card_faces : [];
  const textoOracle = carta.oracle_text || faces.map((face) => face.oracle_text || "").join("\n");
  const contexto = { textoOracle, nomeCarta: carta.name, usosTags };
  const categoriasTagger = filtrarCategorias([...tags], [...relacoes].map(([tag, ancestrais]) => ({ tag, ancestrais })), contexto);
  const categoriasTaggerSecundarias = filtrarCategorias([...secundarias], [], { ...contexto, preservarRedundantes: true }).filter((categoria) => !categoriasTagger.includes(categoria));
  const tipos = [carta.type_line, ...faces.map((face) => face.type_line)].filter((linha) => typeof linha === "string")
    .flatMap((linha) => linha.split(/\s*\/\/\s*/).map((face) => face.split(/\s*\u2014\s*/)[0]))
    .flatMap((linha) => linha.split(/\s+/)).filter((tipo) => !["Basic", "Legendary", "Snow", "World", "Ongoing", "Elite", "Token"].includes(tipo));
  const dados = { tiposCarta: [...new Set(tipos)], categoriasTagger, categoriasTaggerSecundarias,
    usosTagsTagger: usosTags,
    tagsTagger: [...tags], tagsTaggerDiretas: [...diretas], tagsTaggerSecundarias: [...secundarias],
    urlTagger, consultadoTaggerEm: new Date().toISOString(), statusTagger: "ok", erroTagger: "" };
  await cache?.set(chaveCache, dados, VALIDADE.categorias);
  return dados;
}

export async function enriquecerCotacao(cotacao, requisitar = fetch, consultarTags = consultarTagsNavegador) {
  try {
    return { ...cotacao, ...await consultarTagger(cotacao.nome, requisitar, consultarTags) };
  } catch (erro) {
    return { ...cotacao, categoriasTagger: [], categoriasTaggerSecundarias: [], tagsTagger: [], tagsTaggerDiretas: [],
      tagsTaggerSecundarias: [], usosTagsTagger: {}, statusTagger: "erro", erroTagger: erro.message || "Falha ao consultar Scryfall Tagger." };
  }
}