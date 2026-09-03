/* Fachada do PowerPoint: dispatcher de tools (uso futuro pelo agent.js)
 * + re-exports dos módulos para a tela provisória de debug.
 */

import { getCache, setCache, clearCache } from "./cache.js";
import {
  getAllContent,
  isPowerPointApi18Supported,
  getLastReadErrors,
  getLastReadSkipped,
  lerGeometria,
} from "./reader.js";
import {
  buscarCombinado,
  buscarPorSimilaridade,
  buscarPorTabela,
  buscarPorIndice,
} from "./search.js";
import { substituirValor } from "./writer.js";
import {
  inserirTextBox,
  inserirForma,
  inserirLinha,
  moverRedimensionar,
  rotacionar,
  reordenar,
  duplicarShape,
  deletarShape,
  deletarVarias,
  ORDENS_Z,
} from "./inserter.js";
import {
  isPowerPointApi19Supported,
  criarTabela,
  adicionarLinhas,
  adicionarColunas,
  removerLinha,
  removerColuna,
  removerVariasLinhas,
  removerVariasColunas,
  relerTabelaEAtualizaCache,
} from "./tables.js";
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
            left: r.item.left ?? null,
            top: r.item.top ?? null,
            width: r.item.width ?? null,
            height: r.item.height ?? null,
            rotation: r.item.rotation ?? null,
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
          // Bounding box da tabela (todas as células compartilham).
          left: primeira ? (primeira.left ?? null) : null,
          top: primeira ? (primeira.top ?? null) : null,
          width: primeira ? (primeira.width ?? null) : null,
          height: primeira ? (primeira.height ?? null) : null,
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
            left: r.item.left ?? null,
            top: r.item.top ?? null,
            width: r.item.width ?? null,
            height: r.item.height ?? null,
            rotation: r.item.rotation ?? null,
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
        result = { api18: isPowerPointApi18Supported(), api19: isPowerPointApi19Supported() };
        break;
      case "ppt_insert": {
        const pos = {
          left: args.left,
          top: args.top,
          width: args.width,
          height: args.height,
        };
        const kind = String(args.kind || "textbox").toLowerCase();
        if (kind === "textbox") {
          result = await inserirTextBox(args.slideIndex, args.text, pos);
        } else if (kind === "shape") {
          result = await inserirForma(args.slideIndex, args.shapeType, pos, args.text);
        } else if (kind === "line") {
          result = await inserirLinha(args.slideIndex, args.connectorType, pos);
        } else {
          throw new Error(
            `Tipo de inserção desconhecido: ${args.kind}. Use textbox, shape ou line.`
          );
        }
        break;
      }
      case "ppt_move_resize":
        result = await moverRedimensionar(args.shapeId, args.slideIndex, {
          left: args.left,
          top: args.top,
          width: args.width,
          height: args.height,
        });
        break;
      case "ppt_rotate":
        result = await rotacionar(args.shapeId, args.slideIndex, args.degrees);
        break;
      case "ppt_reorder":
        result = await reordenar(args.shapeId, args.slideIndex, args.order);
        break;
      case "ppt_duplicate":
        result = await duplicarShape(args.shapeId, args.slideIndex, {
          left: args.offsetLeft,
          top: args.offsetTop,
        });
        break;
      case "ppt_delete": {
        const itens = Array.isArray(args.items)
          ? args.items
          : [{ slideIndex: args.slideIndex, shapeId: args.shapeId }];
        result =
          itens.length === 1
            ? await deletarShape(itens[0].shapeId, itens[0].slideIndex)
            : await deletarVarias(itens);
        break;
      }
      case "ppt_create_table":
        result = await criarTabela(args.slideIndex, args.rows, args.columns, {
          left: args.left,
          top: args.top,
          values: args.values,
        });
        break;
      case "ppt_table_structure": {
        const dim = String(args.dimension || "rows").toLowerCase();
        const action = String(args.action || "add").toLowerCase();
        if (dim !== "rows" && dim !== "columns") {
          throw new Error(`Dimensão desconhecida: ${args.dimension}. Use rows ou columns.`);
        }
        if (action === "add") {
          result =
            dim === "rows"
              ? await adicionarLinhas(args.shapeId, args.slideIndex, args.count, args.index)
              : await adicionarColunas(args.shapeId, args.slideIndex, args.count, args.index);
        } else if (action === "remove") {
          result =
            dim === "rows"
              ? await removerLinha(args.shapeId, args.slideIndex, args.index)
              : await removerColuna(args.shapeId, args.slideIndex, args.index);
        } else if (action === "removeMany") {
          result =
            dim === "rows"
              ? await removerVariasLinhas(args.shapeId, args.slideIndex, args.indices)
              : await removerVariasColunas(args.shapeId, args.slideIndex, args.indices);
        } else {
          throw new Error(`Ação desconhecida: ${args.action}. Use add, remove ou removeMany.`);
        }
        break;
      }
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
  lerGeometria,
  substituirValor,
  medirTempoLeituraInicial,
  revalidarCacheIncremental,
  iniciarPolling,
  pararPolling,
  getPollingState,
  inserirTextBox,
  inserirForma,
  inserirLinha,
  moverRedimensionar,
  rotacionar,
  reordenar,
  ORDENS_Z,
  duplicarShape,
  deletarShape,
  deletarVarias,
  isPowerPointApi19Supported,
  criarTabela,
  adicionarLinhas,
  adicionarColunas,
  removerLinha,
  removerColuna,
  removerVariasLinhas,
  removerVariasColunas,
  relerTabelaEAtualizaCache,
};
