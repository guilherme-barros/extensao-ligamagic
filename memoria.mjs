export const ALARME_MEMORIA = "verificar-memoria";
const MB = 1024 * 1024;

function formatarBytes(bytes) {
  const unidade = bytes >= 1024 * MB ? "GB" : "MB";
  return (bytes / (unidade === "GB" ? 1024 * MB : MB)).toLocaleString("pt-BR", { maximumFractionDigits: 1 }) + " " + unidade;
}

export async function verificarMemoria({ sistema = globalThis.chrome?.system?.memory,
  armazenamento = globalThis.chrome?.storage?.local, desempenho = globalThis.performance } = {}) {
  const [ram, bytesSalvos] = await Promise.all([
    Promise.resolve().then(() => sistema?.getInfo()).catch(() => null),
    Promise.resolve().then(() => armazenamento?.getBytesInUse(null)).catch(() => null),
  ]);
  const avisos = [];
  let critico = false;
  if (Number.isFinite(ram?.capacity) && ram.capacity > 0 && Number.isFinite(ram.availableCapacity)
    && ram.availableCapacity >= 0 && ram.availableCapacity <= ram.capacity) {
    const proporcaoLivre = ram.availableCapacity / ram.capacity;
    if (proporcaoLivre <= 0.1 || ram.availableCapacity <= 512 * MB) {
      avisos.push("Pouca memoria RAM: " + formatarBytes(ram.availableCapacity) + " livres ("
        + Math.round(proporcaoLivre * 100) + "%). Feche abas ou outros programas para liberar memoria.");
      critico ||= proporcaoLivre <= 0.05 || ram.availableCapacity <= 256 * MB;
    }
  }
  const heap = desempenho?.memory;
  if (Number.isFinite(heap?.jsHeapSizeLimit) && heap.jsHeapSizeLimit > 0
    && Number.isFinite(heap.usedJSHeapSize) && heap.usedJSHeapSize >= 0) {
    const proporcao = heap.usedJSHeapSize / heap.jsHeapSizeLimit;
    if (proporcao >= 0.85) {
      avisos.push("Memoria da extensao quase cheia: " + Math.round(proporcao * 100)
        + "% do limite utilizado. Evite exportacoes grandes com imagens.");
      critico ||= proporcao >= 0.95;
    }
  }
  const limite = armazenamento?.QUOTA_BYTES;
  if (Number.isFinite(bytesSalvos) && bytesSalvos >= 0 && Number.isFinite(limite) && limite > 0) {
    const proporcao = bytesSalvos / limite;
    if (proporcao >= 0.85) {
      avisos.push("Armazenamento das capturas quase cheio: " + Math.round(proporcao * 100)
        + "% utilizado (" + formatarBytes(bytesSalvos) + " de " + formatarBytes(limite)
        + "). Exporte seus dados antes de limpar a lista.");
      critico ||= proporcao >= 0.95;
    }
  }
  return { avisos, critico };
}