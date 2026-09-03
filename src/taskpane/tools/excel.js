const SELECTION_VALUES_CELL_LIMIT = 200;

async function executeTool(toolName, args) {
  try {
    let result;
    switch (toolName) {
      case "read_cells":
        result = await readCells(args.range, args.sheet);
        break;
      case "write_cells":
        result = await writeCells(args.range, args.values, args.sheet);
        break;
      case "format_cells":
        result = await formatCells(args.range, args, args.sheet);
        break;
      case "clear_cells":
        result = await clearCells(args.range, args.sheet);
        break;
      case "resize_cells":
        result = await resizeCells(args.range, args.sheet, args);
        break;
      case "get_sheet_info":
        result = await getSheetInfo();
        break;
      case "get_selection":
        result = await getSelection();
        break;
      default:
        throw new Error("Ferramenta desconhecida: " + toolName);
    }
    return { success: true, data: result };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function resolveSheet(context, sheetName) {
  if (sheetName) {
    return context.workbook.worksheets.getItem(sheetName);
  }
  return context.workbook.worksheets.getActiveWorksheet();
}

function rangeLabel(sheetName, range) {
  return (sheetName ? sheetName + "!" : "") + range;
}

function readCells(range, sheetName) {
  return new Promise((resolve, reject) => {
    Excel.run(async (context) => {
      const rangeObj = resolveSheet(context, sheetName).getRange(range);
      rangeObj.load("values");
      await context.sync();
      resolve(rangeObj.values);
    }).catch(reject);
  });
}

function writeCells(range, values, sheetName) {
  return new Promise((resolve, reject) => {
    Excel.run(async (context) => {
      const rangeObj = resolveSheet(context, sheetName).getRange(range);
      rangeObj.values = values;
      await context.sync();
      resolve("Valores escritos com sucesso em " + rangeLabel(sheetName, range));
    }).catch(reject);
  });
}

function formatCells(range, format, sheetName) {
  return new Promise((resolve, reject) => {
    Excel.run(async (context) => {
      const rangeObj = resolveSheet(context, sheetName).getRange(range);

      if (format.bold !== undefined) rangeObj.format.font.bold = format.bold;
      if (format.italic !== undefined) rangeObj.format.font.italic = format.italic;
      if (format.fontColor) rangeObj.format.font.color = format.fontColor;
      if (format.backgroundColor) rangeObj.format.fill.color = format.backgroundColor;
      if (format.numberFormat) rangeObj.numberFormat = format.numberFormat;

      await context.sync();
      resolve("Formatação aplicada em " + rangeLabel(sheetName, range));
    }).catch(reject);
  });
}

function clearCells(range, sheetName) {
  return new Promise((resolve, reject) => {
    Excel.run(async (context) => {
      const rangeObj = resolveSheet(context, sheetName).getRange(range);
      rangeObj.clear("Contents");
      await context.sync();
      resolve("Células limpas: " + rangeLabel(sheetName, range));
    }).catch(reject);
  });
}

function resizeCells(range, sheetName, options) {
  return new Promise((resolve, reject) => {
    Excel.run(async (context) => {
      const opts = options || {};
      const hasWidth = opts.columnWidth !== undefined && opts.columnWidth !== null;
      const hasHeight = opts.rowHeight !== undefined && opts.rowHeight !== null;
      const autofit = opts.autofit;

      if (!hasWidth && !hasHeight && !autofit) {
        throw new Error(
          'Nada para redimensionar: informe "columnWidth", "rowHeight" ou "autofit".'
        );
      }

      let columnWidth;
      if (hasWidth) {
        columnWidth = Number(opts.columnWidth);
        if (!Number.isFinite(columnWidth) || columnWidth <= 0) {
          throw new Error('Parâmetro "columnWidth" inválido: informe número > 0.');
        }
      }

      let rowHeight;
      if (hasHeight) {
        rowHeight = Number(opts.rowHeight);
        if (!Number.isFinite(rowHeight) || rowHeight <= 0) {
          throw new Error('Parâmetro "rowHeight" inválido: informe número > 0.');
        }
      }

      if (
        autofit !== undefined &&
        autofit !== null &&
        !["columns", "rows", "both"].includes(autofit)
      ) {
        throw new Error('Parâmetro "autofit" inválido: use "columns", "rows" ou "both".');
      }

      const rangeObj = resolveSheet(context, sheetName).getRange(range);
      const applied = [];

      if (hasWidth) {
        rangeObj.format.columnWidth = columnWidth;
        applied.push("largura=" + columnWidth);
      }
      if (hasHeight) {
        rangeObj.format.rowHeight = rowHeight;
        applied.push("altura=" + rowHeight);
      }
      if (autofit === "columns" || autofit === "both") {
        rangeObj.format.autofitColumns();
        applied.push("autofit-colunas");
      }
      if (autofit === "rows" || autofit === "both") {
        rangeObj.format.autofitRows();
        applied.push("autofit-linhas");
      }

      await context.sync();
      resolve(
        "Células redimensionadas em " +
          rangeLabel(sheetName, range) +
          ": " +
          applied.join(", ") +
          "."
      );
    }).catch(reject);
  });
}

function getSheetInfo() {
  return new Promise((resolve, reject) => {
    Excel.run(async (context) => {
      const worksheets = context.workbook.worksheets;
      worksheets.load("items/name");
      const active = worksheets.getActiveWorksheet();
      active.load("name");
      await context.sync();

      const usedRanges = worksheets.items.map((sheet) => sheet.getUsedRangeOrNullObject(true));
      usedRanges.forEach((usedRange) => usedRange.load("address"));
      await context.sync();

      const sheets = worksheets.items.map((sheet, index) => ({
        name: sheet.name,
        active: sheet.name === active.name,
        usedRange: usedRanges[index].isNullObject ? null : usedRanges[index].address,
      }));

      resolve({ activeSheet: active.name, sheets: sheets });
    }).catch(reject);
  });
}

function getSelection() {
  return new Promise((resolve, reject) => {
    Excel.run(async (context) => {
      const rangeObj = context.workbook.getSelectedRange();
      rangeObj.load("address,rowCount,columnCount");
      await context.sync();

      const info = {
        address: rangeObj.address,
        rowCount: rangeObj.rowCount,
        columnCount: rangeObj.columnCount,
      };

      if (rangeObj.rowCount * rangeObj.columnCount <= SELECTION_VALUES_CELL_LIMIT) {
        rangeObj.load("values");
        await context.sync();
        info.values = rangeObj.values;
      }

      resolve(info);
    }).catch(reject);
  });
}

async function describeSelection() {
  try {
    const info = await getSelection();
    let text =
      info.address + " (" + info.rowCount + " linha(s) x " + info.columnCount + " coluna(s))";
    if (info.values) {
      text += " valores: " + JSON.stringify(info.values);
    }
    return text;
  } catch {
    return null;
  }
}

export { executeTool, describeSelection };
