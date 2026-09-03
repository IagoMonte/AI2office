/* Fachada do PowerPoint: dispatcher de tools (uso futuro pelo agent.js)
 * + re-exports dos 4 módulos para a tela provisória de debug.
 */

import { getCache, setCache, clearCache } from "./cache.js";
import {
  getAllContent,
  isPowerPointApi18Supported,
  getLastReadErrors,
  getLastReadSkipped,
} from "./reader.js";
import {
  buscarCombinado,
  buscarPorSimilaridade,
  buscarPorTabela,
  buscarPorIndice,
} from "./search.js";
import { substituirValor } from "./writer.js";
import {
  medirTempoLeituraInicial,
  revalidarCacheIncremental,
  iniciarPolling,
  pararPolling,
  getPollingState,
} from "./poller.js";

async function executeTool(toolName, args) {
  try {
    let result;
    switch (toolName) {
      case "ppt_list_content":
        result = await getAllContent();
        break;
      case "ppt_search": {
        const cache = getCache();
        const ref = args.referenceId ? cache.find((c) => c.id === args.referenceId) || null : null;
        const encontrados = buscarCombinado(args.query, cache, ref, args.weights);
        result = {
          query: args.query,
          total: encontrados.length,
          top: encontrados.slice(0, args.top || 5).map((r) => ({
            id: r.item.id,
            slideIndex: r.item.slideIndex,
            shapeName: r.item.shapeName,
            tipo: r.item.tipo,
            texto: r.item.texto,
            scoreSimilaridade: r.scoreSimilaridade,
            scoreLocalidade: r.scoreLocalidade,
            scoreFinal: r.scoreFinal,
          })),
        };
        break;
      }
      case "ppt_search_table": {
        const cache = getCache();
        const celulas = buscarPorTabela(args.tableId, cache);
        const primeira = celulas[0] || null;
        result = {
          tableId: args.tableId,
          slideIndex: primeira ? primeira.slideIndex : null,
          shapeId: primeira ? primeira.shapeId : null,
          total: celulas.length,
          celulas: celulas.map((item) => ({
            id: item.id,
            index: cache.indexOf(item),
            row: item.row,
            column: item.column,
            parentGroupId: item.parentGroupId,
            texto: item.texto,
          })),
        };
        break;
      }
      case "ppt_search_index": {
        const cache = getCache();
        const encontrados = buscarPorIndice(args.index, cache, args.radius);
        result = {
          index: Math.floor(Number(args.index)),
          radius: args.radius === undefined ? 3 : args.radius,
          totalNoCache: cache.length,
          itens: encontrados.map((r) => ({
            index: r.index,
            id: r.item.id,
            slideIndex: r.item.slideIndex,
            shapeName: r.item.shapeName,
            tipo: r.item.tipo,
            texto: r.item.texto,
          })),
        };
        break;
      }
      case "ppt_replace": {
        const cache = getCache();
        const alvo = cache.find((c) => c.id === args.id);
        if (!alvo)
          throw new Error("Item " + args.id + " não encontrado no cache. Rode ppt_list_content.");
        result = await substituirValor(alvo, args.text);
        break;
      }
      case "ppt_revalidate":
        result = { totalItens: (await revalidarCacheIncremental()).length };
        break;
      case "ppt_api_support":
        result = { api18: isPowerPointApi18Supported() };
        break;
      default:
        throw new Error("Ferramenta desconhecida: " + toolName);
    }
    return { success: true, data: result };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

async function describeSelection() {
  try {
    const cache = getCache();
    if (!cache.length) return "apresentação ainda não lida (rode Módulo 1)";
    const porSlide = {};
    cache.forEach((c) => {
      porSlide[c.slideIndex] = (porSlide[c.slideIndex] || 0) + 1;
    });
    const slides = Object.keys(porSlide).length;
    return `${cache.length} texto(s) em cache em ${slides} slide(s)`;
  } catch {
    return null;
  }
}

export {
  executeTool,
  describeSelection,
  getCache,
  setCache,
  clearCache,
  getAllContent,
  isPowerPointApi18Supported,
  getLastReadErrors,
  getLastReadSkipped,
  buscarCombinado,
  buscarPorSimilaridade,
  buscarPorTabela,
  buscarPorIndice,
  substituirValor,
  medirTempoLeituraInicial,
  revalidarCacheIncremental,
  iniciarPolling,
  pararPolling,
  getPollingState,
};
