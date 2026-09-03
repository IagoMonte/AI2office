/* Módulo 5 — Inserção de conteúdo e manipulação de shapes.
 * Inserção (addTextBox/addGeometricShape/addLine, API 1.4),
 * mover/redimensionar/rotacionar (1.4), reordenar/z-order (1.8),
 * duplicar (recriação com offset) e deletar (1.4, idempotente).
 * Toda operação bem-sucedida atualiza o cache na hora,
 * sem depender do polling do Módulo 4.
 */

import { getCache, setCache } from "./cache.js";
import { isPowerPointApi18Supported, temTextoSuportado } from "./reader.js";

function numOu(valor, padrao) {
  const n = Number(valor);
  return Number.isFinite(n) ? n : padrao;
}

function validarSlideIndex(slideIndex) {
  const idx = Math.floor(Number(slideIndex));
  if (!Number.isFinite(idx) || idx < 0) {
    throw new Error(
      `Slide inválido: ${slideIndex}. Use o índice 0-based (ex: 0 = primeiro slide).`
    );
  }
  return idx;
}

function pushCacheItem(item) {
  const cache = getCache();
  cache.push(item);
  return item;
}

function removerDoCachePorShape(slideIndex, shapeId) {
  const antes = getCache().length;
  setCache(
    getCache().filter(
      (c) => !(c.slideIndex === slideIndex && String(c.shapeId) === String(shapeId))
    )
  );
  return antes - getCache().length;
}

function getSlidePorIndice(context) {
  const slides = context.presentation.slides;
  slides.load("items/id");
  return slides;
}

async function resolverShape(context, slideIndex, shapeId) {
  const slides = getSlidePorIndice(context);
  await context.sync();
  const slide = slides.items[slideIndex];
  if (!slide) {
    throw new Error(`Slide ${slideIndex} não existe mais. Revalide o cache (Módulo 1).`);
  }
  const shape = slide.shapes.getItemOrNullObject(shapeId);
  await context.sync();
  if (shape.isNullObject) {
    throw new Error(
      `Shape ${shapeId} não existe mais no slide ${slideIndex}. Revalide o cache (Módulo 1).`
    );
  }
  return { slide, shape };
}

// --- 5.1 Inserção ---

async function inserirTextBox(slideIndex, texto, posicao = {}) {
  const idx = validarSlideIndex(slideIndex);
  const textoFinal = String(texto === undefined || texto === null ? "" : texto);
  const left = numOu(posicao.left, 50);
  const top = numOu(posicao.top, 50);
  const width = numOu(posicao.width, 300);
  const height = numOu(posicao.height, 100);

  let shapeId = null;
  let shapeName = null;
  let slideId = null;
  await PowerPoint.run(async (context) => {
    const slides = getSlidePorIndice(context);
    await context.sync();
    const slide = slides.items[idx];
    if (!slide) throw new Error(`Slide ${idx} não existe.`);
    slideId = slide.id;

    const shape = slide.shapes.addTextBox(textoFinal, { left, top, width, height });
    shape.load("id, name");
    await context.sync();
    shapeId = shape.id;
    shapeName = shape.name;
  });

  pushCacheItem({
    id: `s${idx}_${shapeId}`,
    slideIndex: idx,
    slideId,
    shapeId,
    shapeName,
    tipo: "TextBox",
    row: null,
    column: null,
    parentGroupId: null,
    texto: textoFinal,
  });
  return { shapeId, shapeName, slideIndex: idx, id: `s${idx}_${shapeId}` };
}

async function inserirForma(slideIndex, tipoForma, posicao = {}, texto = "") {
  const idx = validarSlideIndex(slideIndex);
  const nomeTipo = String(tipoForma || "").trim();
  if (!nomeTipo) throw new Error("Tipo de forma vazio. Ex: Rectangle, Oval, RightArrow.");
  const enumForma = PowerPoint.GeometricShapeType[nomeTipo];
  if (enumForma === undefined) {
    throw new Error(
      `Tipo de forma desconhecido: "${nomeTipo}". Exemplos: Rectangle, Oval, RightArrow, Star5Point.`
    );
  }

  let shapeId = null;
  let shapeName = null;
  let slideId = null;
  await PowerPoint.run(async (context) => {
    const slides = getSlidePorIndice(context);
    await context.sync();
    const slide = slides.items[idx];
    if (!slide) throw new Error(`Slide ${idx} não existe.`);
    slideId = slide.id;

    const shape = slide.shapes.addGeometricShape(enumForma, {
      left: numOu(posicao.left, 50),
      top: numOu(posicao.top, 50),
      width: numOu(posicao.width, 150),
      height: numOu(posicao.height, 100),
    });
    const textoFinal = String(texto || "");
    if (textoFinal) shape.textFrame.textRange.text = textoFinal;
    shape.load("id, name");
    await context.sync();
    shapeId = shape.id;
    shapeName = shape.name;
  });

  pushCacheItem({
    id: `s${idx}_${shapeId}`,
    slideIndex: idx,
    slideId,
    shapeId,
    shapeName,
    tipo: "GeometricShape",
    row: null,
    column: null,
    parentGroupId: null,
    texto: String(texto || ""),
  });
  return { shapeId, shapeName, slideIndex: idx, id: `s${idx}_${shapeId}` };
}

async function inserirLinha(slideIndex, tipoConector = "Straight", posicao = {}) {
  const idx = validarSlideIndex(slideIndex);
  const nomeConector = String(tipoConector || "Straight").trim() || "Straight";
  const enumConector = PowerPoint.ConnectorType[nomeConector];
  if (enumConector === undefined) {
    throw new Error(`Conector desconhecido: "${nomeConector}". Use Straight, Elbow ou Curve.`);
  }

  let shapeId = null;
  let shapeName = null;
  await PowerPoint.run(async (context) => {
    const slides = getSlidePorIndice(context);
    await context.sync();
    const slide = slides.items[idx];
    if (!slide) throw new Error(`Slide ${idx} não existe.`);

    // Em linhas, width/height são o deslocamento do ponto final
    // relativo ao inicial (ex: width 200 + height 0 = linha deitada de 200pt).
    const shape = slide.shapes.addLine(enumConector, {
      left: numOu(posicao.left, 50),
      top: numOu(posicao.top, 50),
      width: numOu(posicao.width, 200),
      height: numOu(posicao.height, 0),
    });
    shape.load("id, name");
    await context.sync();
    shapeId = shape.id;
    shapeName = shape.name;
  });

  // Linha não tem texto: não entra no cache de conteúdo, mas o id
  // retornado permite mover/deletar depois.
  return { shapeId, shapeName, slideIndex: idx };
}

// --- 5.2 Manipulação ---

async function moverRedimensionar(shapeId, slideIndex, novaPosicao = {}) {
  const idx = validarSlideIndex(slideIndex);
  const { left, top, width, height } = novaPosicao || {};
  if (left === undefined && top === undefined && width === undefined && height === undefined) {
    throw new Error("Nada para alterar: informe ao menos left, top, width ou height.");
  }

  let final = null;
  await PowerPoint.run(async (context) => {
    const { shape } = await resolverShape(context, idx, shapeId);
    if (left !== undefined) shape.left = Number(left);
    if (top !== undefined) shape.top = Number(top);
    if (width !== undefined) shape.width = Number(width);
    if (height !== undefined) shape.height = Number(height);
    shape.load("left, top, width, height");
    await context.sync();
    final = { left: shape.left, top: shape.top, width: shape.width, height: shape.height };
  });
  return { shapeId, slideIndex: idx, ...final };
}

async function rotacionar(shapeId, slideIndex, graus) {
  const idx = validarSlideIndex(slideIndex);
  const g = Number(graus);
  if (!Number.isFinite(g))
    throw new Error(`Rotação inválida: ${graus}. Use graus (positivo = horário).`);

  await PowerPoint.run(async (context) => {
    const { shape } = await resolverShape(context, idx, shapeId);
    shape.rotation = g;
    await context.sync();
  });
  return { shapeId, slideIndex: idx, rotation: g };
}

const ORDENS_Z = ["bringToFront", "sendToBack", "bringForward", "sendBackward"];

async function reordenar(shapeId, slideIndex, ordem) {
  const idx = validarSlideIndex(slideIndex);
  const nomeOrdem = String(ordem || "bringToFront").trim() || "bringToFront";
  if (ORDENS_Z.indexOf(nomeOrdem) === -1) {
    throw new Error(`Ordem desconhecida: "${nomeOrdem}". Use: ${ORDENS_Z.join(", ")}.`);
  }
  if (!isPowerPointApi18Supported()) {
    throw new Error("Reordenar (setZOrder) exige PowerPoint API 1.8+.");
  }

  await PowerPoint.run(async (context) => {
    const { shape } = await resolverShape(context, idx, shapeId);
    shape.setZOrder(PowerPoint.ShapeZOrder[nomeOrdem]);
    await context.sync();
  });
  return { shapeId, slideIndex: idx, ordem: nomeOrdem };
}

async function duplicarShape(shapeId, slideIndex, offset = {}) {
  const idx = validarSlideIndex(slideIndex);
  const offLeft = numOu(offset.left, 20);
  const offTop = numOu(offset.top, 20);

  let nova = null;
  await PowerPoint.run(async (context) => {
    const { slide, shape: original } = await resolverShape(context, idx, shapeId);
    original.load("left, top, width, height, type, rotation");
    await context.sync();

    let textoOriginal = "";
    if (
      temTextoSuportado(original.type) &&
      original.type !== "Group" &&
      original.type !== "Table"
    ) {
      try {
        original.textFrame.load("hasText");
        await context.sync();
        if (original.textFrame.hasText) {
          original.textFrame.textRange.load("text");
          await context.sync();
          textoOriginal = original.textFrame.textRange.text || "";
        }
      } catch {
        textoOriginal = "";
      }
    }

    // Limitação conhecida: a duplicação recria como TextBox, preservando
    // posição (+offset), tamanho, rotação e texto — mas não preenchimento,
    // contorno nem o tipo geométrico original.
    const copia = slide.shapes.addTextBox(textoOriginal, {
      left: original.left + offLeft,
      top: original.top + offTop,
      width: original.width,
      height: original.height,
    });
    copia.rotation = original.rotation;
    copia.load("id, name");
    await context.sync();
    nova = { shapeId: copia.id, shapeName: copia.name, texto: textoOriginal };
  });

  pushCacheItem({
    id: `s${idx}_${nova.shapeId}`,
    slideIndex: idx,
    slideId: null,
    shapeId: nova.shapeId,
    shapeName: nova.shapeName,
    tipo: "TextBox",
    row: null,
    column: null,
    parentGroupId: null,
    texto: nova.texto,
  });
  return { ...nova, slideIndex: idx, id: `s${idx}_${nova.shapeId}` };
}

async function deletarShape(shapeId, slideIndex) {
  const idx = validarSlideIndex(slideIndex);
  let removida = false;
  await PowerPoint.run(async (context) => {
    const slides = getSlidePorIndice(context);
    await context.sync();
    const slide = slides.items[idx];
    if (!slide) throw new Error(`Slide ${idx} não existe mais.`);
    const shape = slide.shapes.getItemOrNullObject(shapeId);
    await context.sync();
    if (shape.isNullObject) return; // idempotente: já foi deletada
    shape.delete();
    await context.sync();
    removida = true;
  });

  const removidosDoCache = removerDoCachePorShape(idx, shapeId);
  return { shapeId, slideIndex: idx, removida, removidosDoCache };
}

async function deletarVarias(itens) {
  const lista = Array.isArray(itens) ? itens : [];
  if (!lista.length) throw new Error("Lista vazia: informe ao menos um {slideIndex, shapeId}.");
  const normalizados = lista.map((i) => ({
    slideIndex: validarSlideIndex(i.slideIndex),
    shapeId: String(i.shapeId),
  }));

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    // Agrupa por slide; um único sync para todas as exclusões.
    const porSlide = {};
    normalizados.forEach(({ slideIndex, shapeId }) => {
      porSlide[slideIndex] = porSlide[slideIndex] || [];
      porSlide[slideIndex].push(shapeId);
    });
    for (const key of Object.keys(porSlide)) {
      const slide = slides.items[Number(key)];
      if (!slide) continue;
      porSlide[key].forEach((shapeId) => {
        slide.shapes.getItemOrNullObject(shapeId).delete();
      });
    }
    await context.sync();
  });

  let removidosDoCache = 0;
  normalizados.forEach(({ slideIndex, shapeId }) => {
    removidosDoCache += removerDoCachePorShape(slideIndex, shapeId);
  });
  return { solicitadas: normalizados.length, removidosDoCache };
}

export {
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
};
