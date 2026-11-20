
import { updateCacheItem, findCacheItem } from "./cache.js";
import { lerConteudoDeTabela } from "./reader.js";

async function substituirValor(itemCache, novoTexto) {
  if (!itemCache || !itemCache.id) {
    throw new Error("Item de cache inválido: informe o id do item.");
  }
  const alvo = findCacheItem(itemCache) || itemCache;
  let textoFinal = String(novoTexto === undefined || novoTexto === null ? "" : novoTexto);

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    const slide = slides.items[alvo.slideIndex];
    if (!slide) {
      throw new Error(`Slide ${alvo.slideIndex} não existe mais. Revalide o cache.`);
    }

    if (alvo.tipo === "Table") {
      const shape = slide.shapes.getItemOrNullObject(alvo.shapeId);
      await context.sync();

      if (shape.isNullObject) {
        throw new Error(
          `Shape ${alvo.shapeId} não existe mais no slide ${alvo.slideIndex}. Revalide o cache.`
        );
      }

      const table = shape.getTable();
      const cell = table.getCellOrNullObject(alvo.row, alvo.column);
      await context.sync();

      if (cell.isNullObject) {
        throw new Error(`Célula [${alvo.row},${alvo.column}] não existe mais.`);
      }

      cell.text = textoFinal;
    } else if (alvo.tipo === "GroupChild") {
      await substituirDentroDeGrupo(context, slide, alvo, textoFinal);
      return;
    } else {
      const shape = slide.shapes.getItemOrNullObject(alvo.shapeId);
      await context.sync();

      if (shape.isNullObject) {
        throw new Error(
          `Shape ${alvo.shapeId} não existe mais no slide ${alvo.slideIndex}. Revalide o cache.`
        );
      }

      shape.textFrame.textRange.text = textoFinal;
    }

    await context.sync();
  });

  updateCacheItem(alvo.id, textoFinal);
  return { id: alvo.id, texto: textoFinal };
}

async function substituirDentroDeGrupo(context, slide, itemCache, novoTexto) {
  const grupoShape = slide.shapes.getItemOrNullObject(itemCache.parentGroupId);
  await context.sync();

  if (grupoShape.isNullObject) {
    throw new Error("Grupo pai não existe mais. Revalide o cache.");
  }

  grupoShape.group.ungroup();
  await context.sync();

  slide.shapes.load("items/id, items/name");
  await context.sync();

  const filha = slide.shapes.items.find((s) => s.id === itemCache.shapeId);
  if (!filha) throw new Error("Shape filha não encontrada após ungroup.");

  filha.textFrame.textRange.text = novoTexto;
  await context.sync();

  const idsParaReagrupar = slide.shapes.items.map((s) => s.id);
  slide.shapes.addGroup(idsParaReagrupar);
  await context.sync();
}

export { substituirValor, substituirDentroDeGrupo, findCacheItem };
export { lerConteudoDeTabela };
