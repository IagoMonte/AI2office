export const SYSTEM_PROMPT = `Você é um assistente especializado em manipular planilhas Excel/Office. Você tem acesso a ferramentas para ler, escrever e formatar células.

IMPORTANTE:
- Use as ferramentas disponíveis para interagir com a planilha
- Sempre confirme ações destrutivas (como limpar células) antes de executá-las
- Formate números, datas e moedas apropriadamente
- Use endereços de célula padrão (ex: "A1", "B2:C10")

Ferramentas disponíveis:
- read_cells: Lê valores de um intervalo de células
- write_cells: Escreve valores em células
- format_cells: Aplica formatação (negrito, itálico, cores, etc)
- clear_cells: Limpa o conteúdo de células
- get_sheet_info: Obtém informações sobre a planilha atual
- get_selection: Obtém o intervalo selecionado no momento (endereço, dimensões e valores)

Quando o usuário se referir à seleção atual ("aqui", "isto", "essas células", "minha seleção"), use get_selection para identificá-la antes de agir. O contexto do sistema pode incluir a seleção corrente; se os valores não vierem incluídos (seleções grandes), use read_cells com o endereço informado.`;

export const TOOLS = [
  {
    type: "function",
    function: {
      name: "read_cells",
      description: "Lê valores de um intervalo de células na planilha ativa",
      parameters: {
        type: "object",
        properties: {
          range: {
            type: "string",
            description: "Intervalo de células (ex: 'A1', 'A1:C10', 'A:A')",
          },
        },
        required: ["range"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "write_cells",
      description: "Escreve valores em células ou intervalo de células",
      parameters: {
        type: "object",
        properties: {
          range: {
            type: "string",
            description: "Intervalo de células (ex: 'A1', 'A1:C10')",
          },
          values: {
            type: "array",
            description: "Array 2D de valores para escrever (array de arrays)",
            items: {
              type: "array",
              items: {},
            },
          },
        },
        required: ["range", "values"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "format_cells",
      description: "Aplica formatação a células",
      parameters: {
        type: "object",
        properties: {
          range: {
            type: "string",
            description: "Intervalo de células para formatar",
          },
          bold: {
            type: "boolean",
            description: "Aplicar negrito",
          },
          italic: {
            type: "boolean",
            description: "Aplicar itálico",
          },
          fontColor: {
            type: "string",
            description: "Cor da fonte (hex, ex: '#FF0000')",
          },
          backgroundColor: {
            type: "string",
            description: "Cor de fundo (hex, ex: '#FFFF00')",
          },
          numberFormat: {
            type: "string",
            description: "Formato de número (ex: '0.00', 'yyyy-mm-dd', '$#,##0.00')",
          },
        },
        required: ["range"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "clear_cells",
      description: "Limpa o conteúdo de células",
      parameters: {
        type: "object",
        properties: {
          range: {
            type: "string",
            description: "Intervalo de células para limpar",
          },
        },
        required: ["range"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "get_sheet_info",
      description: "Obtém informações sobre a planilha ativa",
      parameters: {
        type: "object",
        properties: {},
        required: [],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "get_selection",
      description:
        "Obtém o intervalo atualmente selecionado na planilha ativa (endereço, dimensões e, se a seleção for pequena, os valores)",
      parameters: {
        type: "object",
        properties: {},
        required: [],
      },
    },
  },
];
