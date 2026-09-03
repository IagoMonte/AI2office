/* Módulo 4 — Atualização via polling (adaptativo + incremental). */

import { getCache, setCache } from "./cache.js";
import { lerConteudoDeTabela, lerGeometria, temTextoSuportado } from "./reader.js";

let intervalId = null;
let tempoUltimaLeituraMs = null;
let ultimoIntervaloMs = null;

function debounce(fn, ms) {
  let timeoutId;
  return (...args) => {
    clearTimeout(timeoutId);
    timeoutId = setTimeout(() => fn(...args), ms);
  };
}

async function medirTempoLeituraInicial(getAllContentFn) {
  const inicio = performance.now();
  const conteudo = await getAllContentFn();
  tempoUltimaLeituraMs = performance.now() - inicio;
  return { tempoMs: tempoUltimaLeituraMs, totalItens: conteudo.length };
}

function calcularIntervaloPolling(tempoLeituraMs) {
  // Regra: 20x o tempo de uma leitura completa, piso 5s, teto 60s.
  const intervaloSugerido = tempoLeituraMs * 20;
  return Math.min(Math.max(intervaloSugerido, 5000), 60000);
}

async function revalidarCacheIncremental() {
  let cacheConteudo = getCache();

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    for (let i = 0; i < slides.items.length; i++) {
      const shapes = slides.items[i].shapes;
      shapes.load("items/id, items/type, items/name");
      await context.sync();

      const idsAtuais = new Set(shapes.items.map((s) => s.id));
      const itensCacheDoSlide = cacheConteudo.filter((c) => c.slideIndex === i);
      const idsCache = new Set(itensCacheDoSlide.map((c) => c.shapeId));

      const idsRemovidos = [...idsCache].filter((id) => !idsAtuais.has(id));
      const shapesNovas = shapes.items.filter((s) => !idsCache.has(s.id));

      if (idsRemovidos.length > 0) {
        cacheConteudo = cacheConteudo.filter(
          (c) => !(c.slideIndex === i && idsRemovidos.includes(c.shapeId))
        );
      }

      for (const shapeNova of shapesNovas) {
        if (shapeNova.type === "Table") {
          const novos = await lerConteudoDeTabela(context, shapeNova, {
            slideIndex: i,
            slideId: slides.items[i].id,
            shapeId: shapeNova.id,
            shapeName: shapeNova.name,
            tipo: "Table",
          });
          cacheConteudo.push(...novos);
        } else if (shapeNova.type !== "Group" && temTextoSuportado(shapeNova.type)) {
          shapeNova.textFrame.load("hasText");
          await context.sync();
          if (shapeNova.textFrame.hasText) {
            shapeNova.textFrame.textRange.load("text");
            await context.sync();
            const geo = await lerGeometria(context, shapeNova);
            cacheConteudo.push({
              id: `s${i}_${shapeNova.id}`,
              slideIndex: i,
              slideId: slides.items[i].id,
              shapeId: shapeNova.id,
              shapeName: shapeNova.name,
              tipo: shapeNova.type,
              row: null,
              column: null,
              parentGroupId: null,
              texto: shapeNova.textFrame.textRange.text || "",
              ...geo,
            });
          }
        }
      }

      // Shapes que continuam existindo: revalida texto e geometria
      // (o usuário pode ter arrastado a shape no slide).
      for (const itemCacheado of itensCacheDoSlide) {
        if (
          idsAtuais.has(itemCacheado.shapeId) &&
          itemCacheado.tipo !== "Table" &&
          itemCacheado.tipo !== "GroupChild"
        ) {
          const shape = shapes.items.find((s) => s.id === itemCacheado.shapeId);
          if (shape && shape.type !== "Group" && temTextoSuportado(shape.type)) {
            shape.textFrame.load("hasText");
            await context.sync();
            if (shape.textFrame.hasText) {
              shape.textFrame.textRange.load("text");
              shape.load("left, top, width, height, rotation");
              await context.sync();
              const idxCache = cacheConteudo.findIndex((c) => c.id === itemCacheado.id);
              if (idxCache !== -1) {
                cacheConteudo[idxCache].texto = shape.textFrame.textRange.text || "";
                cacheConteudo[idxCache].left = shape.left ?? null;
                cacheConteudo[idxCache].top = shape.top ?? null;
                cacheConteudo[idxCache].width = shape.width ?? null;
                cacheConteudo[idxCache].height = shape.height ?? null;
                cacheConteudo[idxCache].rotation = shape.rotation ?? null;
              }
            }
          }
        }
      }
    }
  });

  setCache(cacheConteudo);
  return cacheConteudo;
}

async function iniciarPolling(getAllContentFn, onTick) {
  if (tempoUltimaLeituraMs === null) {
    if (typeof getAllContentFn === "function") {
      await medirTempoLeituraInicial(getAllContentFn);
    } else {
      tempoUltimaLeituraMs = 500;
    }
  }
  const intervalo = calcularIntervaloPolling(tempoUltimaLeituraMs);
  pararPolling();
  ultimoIntervaloMs = intervalo;

  intervalId = setInterval(async () => {
    try {
      const cache = await revalidarCacheIncremental();
      if (typeof onTick === "function") onTick(null, cache);
    } catch (err) {
      if (typeof onTick === "function") onTick(err);
    }
  }, intervalo);

  // Complementa com o evento nativo quando disponível (Windows/web).
  try {
    if (
      Office.context &&
      Office.context.document &&
      typeof Office.context.document.addHandlerAsync === "function"
    ) {
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        debounce(async () => {
          try {
            const cache = await revalidarCacheIncremental();
            if (typeof onTick === "function") onTick(null, cache);
          } catch (err) {
            if (typeof onTick === "function") onTick(err);
          }
        }, 400)
      );
    }
  } catch {
    // Mac ou hosts sem o evento: polling por tempo cobre o buraco.
  }

  return intervalo;
}

function pararPolling() {
  if (intervalId !== null) {
    clearInterval(intervalId);
    intervalId = null;
  }
}

function getPollingState() {
  return {
    ativo: intervalId !== null,
    tempoUltimaLeituraMs,
    intervaloMs: ultimoIntervaloMs,
    totalItens: getCache().length,
  };
}

export {
  medirTempoLeituraInicial,
  calcularIntervaloPolling,
  revalidarCacheIncremental,
  iniciarPolling,
  pararPolling,
  getPollingState,
  debounce,
};
