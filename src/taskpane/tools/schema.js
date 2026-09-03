import { EXCEL_SYSTEM_PROMPT, EXCEL_TOOLS } from "./schemas/excel.js";
import { WORD_SYSTEM_PROMPT, WORD_TOOLS } from "./schemas/word.js";

const SCHEMAS = {
  excel: { SYSTEM_PROMPT: EXCEL_SYSTEM_PROMPT, TOOLS: EXCEL_TOOLS },
  word: { SYSTEM_PROMPT: WORD_SYSTEM_PROMPT, TOOLS: WORD_TOOLS },
};

function getSchemaForHost(host) {
  return SCHEMAS[host] || SCHEMAS.excel;
}

export { getSchemaForHost };
