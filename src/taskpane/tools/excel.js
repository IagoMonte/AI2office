const SELECTION_VALUES_CELL_LIMIT = 200;
const VERIFY_CELL_LIMIT = 10000;

function excelApiLevel() {
  try {
    if (typeof Office === "undefined" || !Office.context || !Office.context.requirements) {
      return null;
    }
    let level = null;
    for (let v = 1; v <= 17; v++) {
      if (Office.context.requirements.isSetSupported("ExcelApi", "1." + v)) {
        level = "1." + v;
      }
    }
    return level;
  } catch {
    return null;
  }
}

function normalizeCell(value) {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value);
}

function valuesEqual(expected, actual) {
  if (!Array.isArray(expected) || !Array.isArray(actual) || expected.length !== actual.length) {
    return false;
  }
  for (let r = 0; r < expected.length; r++) {
    const expRow = expected[r] || [];
    const actRow = actual[r] || [];
    if (expRow.length !== actRow.length) {
      return false;
    }
    for (let c = 0; c < expRow.length; c++) {
      if (normalizeCell(expRow[c]) !== normalizeCell(actRow[c])) {
        return false;
      }
    }
  }
  return true;
}

function containsFormula(values) {
  return (values || []).some((line) =>
    (line || []).some((cell) => typeof cell === "string" && cell.startsWith("="))
  );
}

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
      case "get_range_info":
        result = await getRangeInfo(args.range, args.sheet);
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
      const sheet = resolveSheet(context, sheetName);
      const rangeObj = sheet.getRange(range);
      await assertWritableRange(context, sheet, rangeObj, rangeLabel(sheetName, range));
      rangeObj.values = values;
      await context.sync();

      const rows = Array.isArray(values) ? values.length : 0;
      const cols = rows && Array.isArray(values[0]) ? values[0].length : 0;
      if (rows * cols > 0 && rows * cols <= VERIFY_CELL_LIMIT && !containsFormula(values)) {
        const check = sheet.getRange(range);
        check.load("values");
        await context.sync();
        if (!valuesEqual(values, check.values)) {
          throw new Error(
            "A escrita em " +
              rangeLabel(sheetName, range) +
              " foi aceita mas NÃO persistiu ao reler (lido: " +
              JSON.stringify(check.values) +
              "). " +
              WRITE_PERSIST_HELP
          );
        }
      }

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
      const sheet = resolveSheet(context, sheetName);
      const rangeObj = sheet.getRange(range);
      await assertWritableRange(context, sheet, rangeObj, rangeLabel(sheetName, range));
      try {
        rangeObj.clear("Contents");
        await context.sync();
      } catch (error) {
        throw new Error(
          "Falha ao limpar " + rangeLabel(sheetName, range) + ": " + describeOfficeError(error)
        );
      }

      const dims = await getRangeDims(context, sheet.getRange(range));
      if (dims.rowCount * dims.columnCount <= VERIFY_CELL_LIMIT) {
        const check = sheet.getRange(range);
        check.load("values");
        await context.sync();
        const stillThere = (check.values || []).some((line) =>
          (line || []).some((cell) => normalizeCell(cell) !== "")
        );
        if (stillThere) {
          throw new Error(
            "A limpeza de " +
              rangeLabel(sheetName, range) +
              " foi aceita mas o conteúdo persiste ao reler. " +
              WRITE_PERSIST_HELP
          );
        }
      }

      resolve("Células limpas: " + rangeLabel(sheetName, range));
    }).catch(reject);
  });
}

function describeOfficeError(error) {
  const message = (error && error.message) || String(error);
  if (/merged/i.test(message)) {
    return (
      message +
      " (o intervalo toca células mescladas; escrita só fixa na célula superior-esquerda do merge)."
    );
  }
  return message;
}

async function getRangeDims(context, rangeObj) {
  rangeObj.load("address,rowCount,columnCount");
  await context.sync();
  return {
    address: rangeObj.address,
    rowCount: rangeObj.rowCount,
    columnCount: rangeObj.columnCount,
  };
}

async function readProtectionOrNull(context, sheet) {
  try {
    sheet.protection.load("protected");
    await context.sync();
    return !!sheet.protection.protected;
  } catch {
    return null;
  }
}

async function assertWritableRange(context, sheet, rangeObj, label) {
  const protection = await readProtectionOrNull(context, sheet);
  if (protection === true) {
    throw new Error(
      "A aba está protegida: desproteja (Revisar > Desproteger Planilha) ou peça ao usuário para desproteger antes de escrever em " +
        label +
        "."
    );
  }
}

const WRITE_PERSIST_HELP =
  "Causas prováveis: (1) células mescladas — o Office.js não permite detectar merge via API, confira no Excel selecionando a célula e vendo o botão Mesclar; escrita só fixa na célula superior-esquerda do merge; " +
  "(2) macro/evento VBA revertendo a edição (Alt+F11 para verificar); " +
  "(3) edição concorrente. Verifique com get_range_info.";

function getRangeInfo(range, sheetName) {
  return new Promise((resolve, reject) => {
    Excel.run(async (context) => {
      const sheet = resolveSheet(context, sheetName);
      const rangeObj = sheet.getRange(range);
      const protection = await readProtectionOrNull(context, sheet);
      const info = await getRangeDims(context, rangeObj);

      let values = null;
      let formulas = null;
      let numberFormat = null;
      if (info.rowCount * info.columnCount <= VERIFY_CELL_LIMIT) {
        const probe = sheet.getRange(range);
        probe.load("values,formulas,numberFormat");
        await context.sync();
        values = probe.values;
        formulas = probe.formulas;
        numberFormat = probe.numberFormat;
      }

      resolve({
        range: rangeLabel(sheetName, range),
        address: info.address,
        rowCount: info.rowCount,
        columnCount: info.columnCount,
        mergeState: "unknown (o Office.js não expõe leitura de mesclagem; confira no Excel)",
        sheetProtected: protection,
        excelApi: excelApiLevel(),
        values: values,
        formulas: formulas,
        numberFormat: numberFormat,
      });
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
