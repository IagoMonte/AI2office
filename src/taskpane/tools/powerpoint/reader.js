
import { setCache } from "./cache.js";

let lastReadErrors = [];
let lastReadSkipped = [];

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

  const geoTabela = await lerGeometria(context, shape);

  for (let r = 0; r < table.rowCount; r++) {
    for (let c = 0; c < table.columnCount; c++) {
      try {
        const cell = table.getCellOrNullObject(r, c);
        cell.load("text");
        await context.sync();
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

  grupoShape.group.ungroup();
  await context.sync();

  slide.shapes.load("items/id, items/type, items/name");
  await context.sync();

  const itens = [];
  const shapesFilhas = slide.shapes.items.filter((s) => s.name !== nomeGrupoOriginal);

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
    if (!reagrupado) {
      try {
        const idsParaReagrupar = shapesFilhas.map((s) => s.id);
        if (idsParaReagrupar.length > 0) {
          slide.shapes.addGroup(idsParaReagrupar);
          await context.sync();
        }
        reagrupado = true;
      } catch {
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
