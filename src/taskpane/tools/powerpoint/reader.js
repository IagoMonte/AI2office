/* Módulo 1 — Leitura e disponibilização de conteúdo (PowerPoint).
 * Lê shapes normais, tabelas e shapes dentro de grupos.
 * Segunda fase de carga: geometria (left, top, width, height, rotation)
 * para noção de posicionamento — células de tabela carregam a
 * bounding box da tabela + row/column.
 * Requer PowerPoint API 1.8+ para ungroup()/addGroup().
 */

import { setCache } from "./cache.js";

let lastReadErrors = [];
let lastReadSkipped = [];

// Tipos de shape que nunca têm textFrame (confirmado via debug:
// Image e Line lançam InvalidArgument ao acessar textFrame).
// Chart/Diagram/Media entram por segurança — se o nome do enum não
// bater, o try/catch por shape continua capturando.
const SHAPES_SEM_TEXTO = ["Image", "Line", "Chart", "Diagram", "Media", "Ink", "OleObject"];

function getLastReadErrors() {
  return lastReadErrors;
}

function getLastReadSkipped() {
  return lastReadSkipped;
}

function temTextoSuportado(tipo) {
  return SHAPES_SEM_TEXTO.indexOf(tipo) === -1;
}

function erroMsg(e) {
  return (e && e.message ? e.message : String(e)) + (e && e.code ? ` [${e.code}]` : "");
}

// Segunda fase de carga do Módulo 1: geometria da shape (pontos,
// origem no canto superior esquerdo do slide). Nunca pode quebrar
// a leitura do texto — em caso de falha retorna tudo null.
async function lerGeometria(context, shape) {
  const vazia = { left: null, top: null, width: null, height: null, rotation: null };
  try {
    shape.load("left, top, width, height, rotation");
    await context.sync();
    return {
      left: shape.left ?? null,
      top: shape.top ?? null,
      width: shape.width ?? null,
      height: shape.height ?? null,
      rotation: shape.rotation ?? null,
    };
  } catch {
    return vazia;
  }
}

function isPowerPointApi18Supported() {
  try {
    // O nome correto do requirement set é "PowerPointApi" ("PowerPoint"
    // sozinho não é um set válido e isSetSupported retorna false).
    return Office.context.requirements.isSetSupported("PowerPointApi", "1.8");
  } catch {
    return false;
  }
}

async function lerConteudoDeTabela(context, shape, base) {
  const itens = [];
  const table = shape.getTable();
  table.load("rowCount, columnCount");
  await context.sync();

  // Geometria da tabela toda (cada célula carrega a bounding box da
  // tabela + sua posição row/column — suficiente para posicionamento).
  const geoTabela = await lerGeometria(context, shape);

  for (let r = 0; r < table.rowCount; r++) {
    for (let c = 0; c < table.columnCount; c++) {
      try {
        const cell = table.getCellOrNullObject(r, c);
        cell.load("text");
        await context.sync();
        // Células mescladas/encobertas retornam NullObject: pular sem erro.
        if (cell.isNullObject) {
          continue;
        }
        itens.push({
          id: `s${base.slideIndex}_${shape.id}_r${r}c${c}`,
          ...base,
          row: r,
          column: c,
          parentGroupId: null,
          texto: cell.text || "",
          ...geoTabela,
        });
      } catch (erroCell) {
        lastReadErrors.push({
          slideIndex: base.slideIndex,
          shapeId: shape.id,
          shapeName: shape.name,
          tipo: "TableCell",
          row: r,
          column: c,
          erro: erroMsg(erroCell),
        });
      }
    }
  }
  return itens;
}

async function lerConteudoDeGrupo(context, slide, grupoShape, slideIdx) {
  const idGrupoOriginal = grupoShape.id;
  const nomeGrupoOriginal = grupoShape.name;

  // Requer PowerPoint API 1.8+.
  grupoShape.group.ungroup();
  await context.sync();

  slide.shapes.load("items/id, items/type, items/name");
  await context.sync();

  const itens = [];
  // Heurística: tudo que não tem o nome do grupo original são (prováveis) filhas.
  // A lista de IDs é persistida implicitamente via parentGroupId + shapeId.
  const shapesFilhas = slide.shapes.items.filter((s) => s.name !== nomeGrupoOriginal);

  // O finally garante o reagrupamento mesmo se a leitura de uma filha falhar,
  // para não deixar a apresentação desfeita no meio do debug.
  let reagrupado = false;
  try {
    for (const filha of shapesFilhas) {
      if (!temTextoSuportado(filha.type) && filha.type !== "Table") {
        lastReadSkipped.push({
          slideIndex: slideIdx,
          shapeId: filha.id,
          shapeName: filha.name,
          tipo: filha.type,
          motivo: "tipo sem textFrame (pulado antes da leitura)",
          parentGroupId: idGrupoOriginal,
        });
        continue;
      }
      if (filha.type === "Table") {
        const t = await lerConteudoDeTabela(context, filha, {
          slideIndex: slideIdx,
          slideId: slide.id,
          shapeId: filha.id,
          shapeName: filha.name,
          tipo: "Table",
        });
        t.forEach((item) => (item.parentGroupId = idGrupoOriginal));
        itens.push(...t);
      } else {
        filha.textFrame.load("hasText");
        await context.sync();
        if (filha.textFrame.hasText) {
          filha.textFrame.textRange.load("text");
          await context.sync();
          const geo = await lerGeometria(context, filha);
          itens.push({
            id: `s${slideIdx}_${filha.id}`,
            slideIndex: slideIdx,
            slideId: slide.id,
            shapeId: filha.id,
            shapeName: filha.name,
            tipo: "GroupChild",
            row: null,
            column: null,
            parentGroupId: idGrupoOriginal,
            texto: filha.textFrame.textRange.text || "",
            ...geo,
          });
        }
      }
    }
  } finally {
    // Reagrupa para preservar a estrutura visual original.
    if (!reagrupado) {
      try {
        const idsParaReagrupar = shapesFilhas.map((s) => s.id);
        if (idsParaReagrupar.length > 0) {
          slide.shapes.addGroup(idsParaReagrupar);
          await context.sync();
        }
        reagrupado = true;
      } catch {
        // Mantém o erro original da leitura; o reagrupamento falhou junto.
      }
    }
  }

  return itens;
}

async function getAllContent(onProgress) {
  const resultado = [];
  lastReadErrors = [];
  lastReadSkipped = [];

  const avisar = (info) => {
    if (typeof onProgress === "function") {
      try {
        onProgress(info);
      } catch {
        // Progresso é só log de debug; nunca pode quebrar a leitura.
      }
    }
  };

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    avisar({ fase: "slides", totalSlides: slides.items.length });

    for (let slideIdx = 0; slideIdx < slides.items.length; slideIdx++) {
      const slide = slides.items[slideIdx];
      const shapes = slide.shapes;
      shapes.load("items/id, items/type, items/name");
      await context.sync();

      avisar({ fase: "slide", slideIndex: slideIdx, totalSlides: slides.items.length });

      for (const shape of shapes.items) {
        const base = {
          slideIndex: slideIdx,
          slideId: slide.id,
          shapeId: shape.id,
          shapeName: shape.name,
          tipo: shape.type,
        };

        // Uma shape problemática (gráfico, SmartArt, mídia...) não pode
        // derrubar a leitura da apresentação inteira: registra e segue.
        // Tipos sabidamente sem textFrame nem passam pelo load.
        try {
          if (!temTextoSuportado(shape.type) && shape.type !== "Table" && shape.type !== "Group") {
            lastReadSkipped.push({
              slideIndex: slideIdx,
              shapeId: shape.id,
              shapeName: shape.name,
              tipo: shape.type,
              motivo: "tipo sem textFrame (pulado antes da leitura)",
            });
            continue;
          }
          if (shape.type === "Group") {
            if (!isPowerPointApi18Supported()) {
              // Sem API 1.8 não há como ler dentro do grupo: registra placeholder para debug.
              const geoGrupo = await lerGeometria(context, shape);
              resultado.push({
                id: `s${slideIdx}_${shape.id}_grupo-nao-suportado`,
                ...base,
                row: null,
                column: null,
                parentGroupId: null,
                texto: "[Grupo: leitura exige PowerPoint API 1.8+]",
                ...geoGrupo,
              });
              continue;
            }
            const itensDoGrupo = await lerConteudoDeGrupo(context, slide, shape, slideIdx);
            resultado.push(...itensDoGrupo);
          } else if (shape.type === "Table") {
            const itensDaTabela = await lerConteudoDeTabela(context, shape, base);
            resultado.push(...itensDaTabela);
          } else {
            shape.textFrame.load("hasText");
            await context.sync();
            if (shape.textFrame.hasText) {
              shape.textFrame.textRange.load("text");
              await context.sync();
              // Segunda fase de carga: geometria para noção de posicionamento.
              const geo = await lerGeometria(context, shape);
              resultado.push({
                id: `s${slideIdx}_${shape.id}`,
                ...base,
                row: null,
                column: null,
                parentGroupId: null,
                texto: shape.textFrame.textRange.text || "",
                ...geo,
              });
            }
          }
        } catch (erroShape) {
          lastReadErrors.push({
            slideIndex: slideIdx,
            shapeId: shape.id,
            shapeName: shape.name,
            tipo: shape.type,
            erro: erroMsg(erroShape),
          });
        }
      }
    }
  });

  setCache(resultado);
  return resultado;
}

export {
  getAllContent,
  lerConteudoDeTabela,
  lerConteudoDeGrupo,
  lerGeometria,
  isPowerPointApi18Supported,
  getLastReadErrors,
  getLastReadSkipped,
  temTextoSuportado,
  SHAPES_SEM_TEXTO,
};
