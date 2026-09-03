import { EXCEL_SYSTEM_PROMPT, EXCEL_TOOLS } from "./schemas/excel.js";
import { WORD_SYSTEM_PROMPT, WORD_TOOLS } from "./schemas/word.js";
import { POWERPOINT_SYSTEM_PROMPT, POWERPOINT_TOOLS } from "./schemas/powerpoint.js";

const SCHEMAS = {
  excel: { SYSTEM_PROMPT: EXCEL_SYSTEM_PROMPT, TOOLS: EXCEL_TOOLS },
  word: { SYSTEM_PROMPT: WORD_SYSTEM_PROMPT, TOOLS: WORD_TOOLS },
  powerpoint: { SYSTEM_PROMPT: POWERPOINT_SYSTEM_PROMPT, TOOLS: POWERPOINT_TOOLS },
};

function getSchemaForHost(host) {
  return SCHEMAS[host] || SCHEMAS.excel;
}

export { getSchemaForHost };
