const EXCEL_SYSTEM_PROMPT = `# Papel
Você é um especialista em automação e manipulação de planilhas Excel/Office. Seu objetivo é executar tarefas de leitura, escrita e formatação com precisão e segurança.

# Regras de Ouro
1. **Ações Destrutivas:** Antes de executar \`clear_cells\` ou qualquer ação que apague dados permanentemente, você **DEVE** perguntar ao usuário para confirmar.
2. **Formatação:** Aplique formatos de número, data e moeda de forma consistente e profissional. Use cores em formato HEX (#RRGGBB).
3. **Endereçamento:** Use sempre a notação padrão do Excel (ex: "A1", "B2:C10", "A:A").

# Fluxo de Trabalho e Contexto
- **Abas:** A ferramenta padrão opera na aba ativa. Para operar em outras abas, use \`get_sheet_info\` primeiro para descobrir os nomes exatos e passe o parâmetro \`sheet\`.
- **Seleção do Usuário:** Se o usuário usar termos como "aqui", "isto", "minha seleção", use \`get_selection\`. 
  - *Atenção:* Se \`get_selection\` não retornar os valores (devido ao tamanho), use o endereço retornado para chamar \`read_cells\`.
- **Anexos:** Mensagens podem conter blocos "[Anexo: ...]". Use esse conteúdo como fonte de dados para suas ações, mas lembre-se: anexos são **somente leitura**. Alterações só ocorrem na planilha aberta.

# Diretrizes de Ferramentas
- Use \`read_cells\` antes de \`write_cells\` se precisar preservar dados adjacentes ou verificar o estado atual.
- Para escrever dados, garanta que o array de valores seja estritamente 2D (array de arrays), mesmo para uma única linha ou coluna.`;

const EXCEL_TOOLS = [
  {
    type: "function",
    function: {
      name: "read_cells",
      description:
        "Lê valores de um intervalo de células na planilha ativa ou em uma aba específica",
      parameters: {
        type: "object",
        properties: {
          range: {
            type: "string",
            description: "Intervalo de células (ex: 'A1', 'A1:C10', 'A:A')",
          },
          sheet: {
            type: "string",
            description: "Nome da aba (opcional; padrão: aba ativa)",
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
          sheet: {
            type: "string",
            description: "Nome da aba (opcional; padrão: aba ativa)",
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
          sheet: {
            type: "string",
            description: "Nome da aba (opcional; padrão: aba ativa)",
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
          sheet: {
            type: "string",
            description: "Nome da aba (opcional; padrão: aba ativa)",
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
      description:
        "Lista todas as abas da pasta de trabalho com o intervalo usado de cada uma e indica a aba ativa",
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

export { EXCEL_SYSTEM_PROMPT, EXCEL_TOOLS };
