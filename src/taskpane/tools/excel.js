const SELECTION_VALUES_CELL_LIMIT = 200;

async function executeTool(toolName, args) {
  try {
    let result;
    switch (toolName) {
      case "read_cells":
        result = await readCells(args.range);
        break;
      case "write_cells":
        result = await writeCells(args.range, args.values);
        break;
      case "format_cells":
        result = await formatCells(args.range, args);
        break;
      case "clear_cells":
        result = await clearCells(args.range);
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

function readCells(range) {
  return new Promise((resolve, reject) => {
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const rangeObj = sheet.getRange(range);
      rangeObj.load("values");
      await context.sync();
      resolve(rangeObj.values);
    }).catch(reject);
  });
}

function writeCells(range, values) {
  return new Promise((resolve, reject) => {
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const rangeObj = sheet.getRange(range);
      rangeObj.values = values;
      await context.sync();
      resolve("Valores escritos com sucesso em " + range);
    }).catch(reject);
  });
}

function formatCells(range, format) {
  return new Promise((resolve, reject) => {
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const rangeObj = sheet.getRange(range);

      if (format.bold !== undefined) rangeObj.format.font.bold = format.bold;
      if (format.italic !== undefined) rangeObj.format.font.italic = format.italic;
      if (format.fontColor) rangeObj.format.font.color = format.fontColor;
      if (format.backgroundColor) rangeObj.format.fill.color = format.backgroundColor;
      if (format.numberFormat) rangeObj.numberFormat = format.numberFormat;

      await context.sync();
      resolve("Formatação aplicada em " + range);
    }).catch(reject);
  });
}

function clearCells(range) {
  return new Promise((resolve, reject) => {
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const rangeObj = sheet.getRange(range);
      rangeObj.clear("Contents");
      await context.sync();
      resolve("Células limpas: " + range);
    }).catch(reject);
  });
}

function getSheetInfo() {
  return new Promise((resolve, reject) => {
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      const usedRange = sheet.getUsedRange();
      usedRange.load("address");
      await context.sync();
      resolve({
        sheetName: sheet.name,
        usedRange: usedRange.address,
      });
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
