import { criarControlador, ALARME_LOTE } from "./lote.mjs";
import { verificarMemoria, ALARME_MEMORIA } from "./memoria.mjs";

const controlador = criarControlador(chrome);

async function atualizarBadge() {
  const { avisos, critico } = await verificarMemoria({ desempenho: null });
  const { capturas = [] } = avisos.length ? {} : await chrome.storage.local.get("capturas");
  await Promise.all([
    chrome.action.setBadgeText({ text: avisos.length ? "!" : capturas.length ? String(capturas.length) : "" }),
    chrome.action.setBadgeBackgroundColor({ color: avisos.length ? (critico ? "#b42318" : "#996300") : "#17644d" }),
    chrome.action.setTitle({ title: "LigaMagic - Captura local" + (avisos.length ? "\n" + avisos.join("\n") : "") }),
  ]);
}

async function iniciarMonitoramentoMemoria() {
  if (!await chrome.alarms.get(ALARME_MEMORIA)) await chrome.alarms.create(ALARME_MEMORIA, { periodInMinutes: 0.5 });
  await atualizarBadge();
}

chrome.runtime.onInstalled.addListener(() => {
  iniciarMonitoramentoMemoria().catch(console.error);
  controlador.reiniciar().catch(console.error);
});

chrome.runtime.onStartup.addListener(() => {
  iniciarMonitoramentoMemoria().catch(console.error);
  controlador.reiniciar().catch(console.error);
});

chrome.storage.onChanged.addListener((alteracoes, area) => {
  if (area === "local") atualizarBadge().catch(console.error);
});

chrome.alarms.onAlarm.addListener((alarme) => {
  if (alarme.name === ALARME_LOTE) controlador.processar().catch(console.error);
  if (alarme.name === ALARME_MEMORIA) atualizarBadge().catch(console.error);
});

chrome.tabs.onUpdated.addListener((abaId, mudanca) => {
  if (mudanca.status === "complete") controlador.processar(abaId).catch(console.error);
});

chrome.tabs.onRemoved.addListener((abaId) => {
  controlador.abaRemovida(abaId).catch(console.error);
});

chrome.runtime.onMessage.addListener((mensagem, remetente, responder) => {
  if (remetente.id !== chrome.runtime.id || mensagem?.tipo !== "lote") return false;
  controlador.comando(mensagem.comando, mensagem.dados).then(
    () => responder({ ok: true }),
    (erro) => responder({ ok: false, erro: erro.message }),
  );
  return true;
});

iniciarMonitoramentoMemoria().catch(console.error);