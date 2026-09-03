/* Módulo 6 — Manipulação de tabelas.
 * Criar tabela nova (addTable, API 1.8), adicionar/remover
 * linhas e colunas (API 1.9). Após qualquer mudança estrutural,
 * relê só aquela tabela para manter o cache consistente.
 */

import { getCache, setCache, resolverRefShape } from "./cache.js";
import { lerConteudoDeTabela, isPowerPointApi18Supported } from "./reader.js";

function isPowerPointApi19Supported() {
  try {
    return Office.context.requirements.isSetSupported("PowerPointApi", "1.9");
  } catch {
    return false;
  }
}

function validarSlideIndex(slideIndex) {
  const idx = Math.floor(Number(slideIndex));
  if (!Number.isFinite(idx) || idx < 0) {
    throw new Error(`Slide inválido: ${slideIndex}. Use o índice 0-based.`);
  }
  return idx;
}

function validarInteiroPositivo(valor, nome) {
  const n = Math.floor(Number(valor));
  if (!Number.isFinite(n) || n < 1) {
    throw new Error(`${nome} inválido: ${valor}. Use um inteiro >= 1.`);
  }
  return n;
}

async function resolverTabela(context, slideIndex, shapeId) {
  const slides = context.presentation.slides;
  slides.load("items/id");
  await context.sync();
  const slide = slides.items[slideIndex];
  if (!slide) {
    throw new Error(`Slide ${slideIndex} não existe mais. Revalide o cache (Módulo 1).`);
  }
  const shape = slide.shapes.getItemOrNullObject(shapeId);
  shape.load("id, name");
  await context.sync();
  if (shape.isNullObject) {
    throw new Error(
      `Shape ${shapeId} não existe mais no slide ${slideIndex}. Revalide o cache (Módulo 1).`
    );
  }
  return { slide, shape };
}

async function relerTabelaEAtualizaCache(shapeId, slideIndex) {
  const alvo = resolverRefShape(shapeId, slideIndex);
  const idx = alvo.slideIndex;
  const sid = alvo.shapeId;
  setCache(getCache().filter((c) => !(c.slideIndex === idx && String(c.shapeId) === sid)));

  let novosItens = [];
  await PowerPoint.run(async (context) => {
    const { shape } = await resolverTabela(context, idx, sid);
    novosItens = await lerConteudoDeTabela(context, shape, {
      slideIndex: idx,
      slideId: null,
      shapeId: shape.id,
      shapeName: shape.name,
      tipo: "Table",
    });
  });

  getCache().push(...novosItens);
  return novosItens;
}

// --- 6.1 Criar tabela nova ---

async function criarTabela(slideIndex, numLinhas, numColunas, posicao = {}) {
  if (!isPowerPointApi18Supported()) {
    throw new Error("Criar tabela (addTable) exige PowerPoint API 1.8+.");
  }
  const idx = validarSlideIndex(slideIndex);
  const linhas = validarInteiroPositivo(numLinhas, "Número de linhas");
  const colunas = validarInteiroPositivo(numColunas, "Número de colunas");

  let shapeId = null;
  let shapeName = null;
  let novosItens = [];
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();
    const slide = slides.items[idx];
    if (!slide) throw new Error(`Slide ${idx} não existe.`);

    const opcoes = {
      left: Number(posicao.left ?? 50),
      top: Number(posicao.top ?? 50),
    };
    if (posicao.values !== undefined) opcoes.values = posicao.values;
    const shape = slide.shapes.addTable(linhas, colunas, opcoes);
    shape.load("id, name");
    await context.sync();
    shapeId = shape.id;
    shapeName = shape.name;

    // Já popula o cache com as células (vazias ou do values), sem polling.
    novosItens = await lerConteudoDeTabela(context, shape, {
      slideIndex: idx,
      slideId: slide.id,
      shapeId,
      shapeName,
      tipo: "Table",
    });
  });

  getCache().push(...novosItens);
  return {
    shapeId,
    shapeName,
    slideIndex: idx,
    tableId: `s${idx}_${shapeId}`,
    linhas,
    colunas,
    celulas: novosItens.length,
  };
}

// --- 6.2 Adicionar linhas/colunas ---

async function adicionarLinhas(shapeId, slideIndex, quantidade = 1, indice = null) {
  if (!isPowerPointApi19Supported()) {
    throw new Error("Adicionar linhas exige PowerPoint API 1.9+.");
  }
  const alvo = resolverRefShape(shapeId, slideIndex);
  const idx = alvo.slideIndex;
  const sid = alvo.shapeId;
  const qtd = validarInteiroPositivo(quantidade, "Quantidade");

  await PowerPoint.run(async (context) => {
    const { shape } = await resolverTabela(context, idx, sid);
    const table = shape.getTable();
    table.rows.load("count");
    await context.sync();
    // Sem índice = adiciona no fim (posição = contagem atual).
    const pos =
      indice === null || indice === undefined ? table.rows.count : Math.floor(Number(indice));
    table.rows.add(pos, qtd);
    await context.sync();
  });

  const novosItens = await relerTabelaEAtualizaCache(sid, idx);
  return {
    shapeId: sid,
    slideIndex: idx,
    adicionadas: qtd,
    celulas: novosItens.length,
  };
}

async function adicionarColunas(shapeId, slideIndex, quantidade = 1, indice = null) {
  if (!isPowerPointApi19Supported()) {
    throw new Error("Adicionar colunas exige PowerPoint API 1.9+.");
  }
  const alvo = resolverRefShape(shapeId, slideIndex);
  const idx = alvo.slideIndex;
  const sid = alvo.shapeId;
  const qtd = validarInteiroPositivo(quantidade, "Quantidade");

  await PowerPoint.run(async (context) => {
    const { shape } = await resolverTabela(context, idx, sid);
    const table = shape.getTable();
    table.columns.load("count");
    await context.sync();
    const pos =
      indice === null || indice === undefined ? table.columns.count : Math.floor(Number(indice));
    table.columns.add(pos, qtd);
    await context.sync();
  });

  const novosItens = await relerTabelaEAtualizaCache(sid, idx);
  return {
    shapeId: sid,
    slideIndex: idx,
    adicionadas: qtd,
    celulas: novosItens.length,
  };
}

// --- 6.3 Remover linhas/colunas ---

async function removerLinha(shapeId, slideIndex, indiceLinha) {
  if (!isPowerPointApi19Supported()) {
    throw new Error("Remover linha exige PowerPoint API 1.9+.");
  }
  const alvo = resolverRefShape(shapeId, slideIndex);
  const idx = alvo.slideIndex;
  const sid = alvo.shapeId;
  const linha = Math.floor(Number(indiceLinha));
  if (!Number.isFinite(linha) || linha < 0) throw new Error(`Linha inválida: ${indiceLinha}.`);

  await PowerPoint.run(async (context) => {
    const { shape } = await resolverTabela(context, idx, sid);
    const table = shape.getTable();
    table.rows.getItemAt(linha).delete();
    await context.sync();
  });

  const novosItens = await relerTabelaEAtualizaCache(sid, idx);
  return { shapeId: sid, slideIndex: idx, removida: linha, celulas: novosItens.length };
}

async function removerColuna(shapeId, slideIndex, indiceColuna) {
  if (!isPowerPointApi19Supported()) {
    throw new Error("Remover coluna exige PowerPoint API 1.9+.");
  }
  const alvo = resolverRefShape(shapeId, slideIndex);
  const idx = alvo.slideIndex;
  const sid = alvo.shapeId;
  const coluna = Math.floor(Number(indiceColuna));
  if (!Number.isFinite(coluna) || coluna < 0) throw new Error(`Coluna inválida: ${indiceColuna}.`);

  await PowerPoint.run(async (context) => {
    const { shape } = await resolverTabela(context, idx, sid);
    const table = shape.getTable();
    table.columns.getItemAt(coluna).delete();
    await context.sync();
  });

  const novosItens = await relerTabelaEAtualizaCache(sid, idx);
  return {
    shapeId: sid,
    slideIndex: idx,
    removida: coluna,
    celulas: novosItens.length,
  };
}

async function removerVariasLinhas(shapeId, slideIndex, indices) {
  if (!isPowerPointApi19Supported()) {
    throw new Error("Remover linhas em lote exige PowerPoint API 1.9+.");
  }
  const alvo = resolverRefShape(shapeId, slideIndex);
  const idx = alvo.slideIndex;
  const sid = alvo.shapeId;
  const lista = Array.isArray(indices) ? indices.map((i) => Math.floor(Number(i))) : [];
  if (!lista.length) throw new Error("Lista de linhas vazia.");
  if (lista.some((i) => !Number.isFinite(i) || i < 0)) {
    throw new Error(`Índices inválidos: ${JSON.stringify(indices)}.`);
  }

  await PowerPoint.run(async (context) => {
    const { shape } = await resolverTabela(context, idx, sid);
    const table = shape.getTable();
    table.rows.load("items");
    await context.sync();
    // Opera sobre referências reais (não posições), então remover em
    // lote não desloca os índices das demais.
    const alvos = lista.map((i) => table.rows.items[i]).filter(Boolean);
    if (!alvos.length) throw new Error("Nenhuma das linhas existe na tabela.");
    table.rows.deleteRows(alvos);
    await context.sync();
  });

  const novosItens = await relerTabelaEAtualizaCache(sid, idx);
  return {
    shapeId: sid,
    slideIndex: idx,
    removidas: lista,
    celulas: novosItens.length,
  };
}

async function removerVariasColunas(shapeId, slideIndex, indices) {
  if (!isPowerPointApi19Supported()) {
    throw new Error("Remover colunas em lote exige PowerPoint API 1.9+.");
  }
  const alvo = resolverRefShape(shapeId, slideIndex);
  const idx = alvo.slideIndex;
  const sid = alvo.shapeId;
  const lista = Array.isArray(indices) ? indices.map((i) => Math.floor(Number(i))) : [];
  if (!lista.length) throw new Error("Lista de colunas vazia.");
  if (lista.some((i) => !Number.isFinite(i) || i < 0)) {
    throw new Error(`Índices inválidos: ${JSON.stringify(indices)}.`);
  }

  await PowerPoint.run(async (context) => {
    const { shape } = await resolverTabela(context, idx, sid);
    const table = shape.getTable();
    table.columns.load("items");
    await context.sync();
    const alvos = lista.map((i) => table.columns.items[i]).filter(Boolean);
    if (!alvos.length) throw new Error("Nenhuma das colunas existe na tabela.");
    table.columns.deleteColumns(alvos);
    await context.sync();
  });

  const novosItens = await relerTabelaEAtualizaCache(sid, idx);
  return {
    shapeId: sid,
    slideIndex: idx,
    removidas: lista,
    celulas: novosItens.length,
  };
}

export {
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
