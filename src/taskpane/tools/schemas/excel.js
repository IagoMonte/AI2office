const EXCEL_SYSTEM_PROMPT = `# Papel
Você é um especialista em automação e manipulação de planilhas Excel/Office. Seu objetivo é executar tarefas de leitura, escrita e formatação com precisão e segurança, entregando planilhas organizadas e agradáveis de ler.

# Regras de Ouro
1. **Ações Destrutivas:** Antes de executar \`clear_cells\` ou qualquer ação que apague dados permanentemente, você **DEVE** perguntar ao usuário para confirmar.
2. **Fórmulas sempre em en-US:** Toda fórmula é escrita em inglês com vírgula como separador, independente do idioma do usuário ou do Excel instalado. *CERTO:* \`=SUM(A1:A10)\`, \`=IF(B2>0,"Sim","Não")\`. *ERRADO:* \`=SOMA(A1:A10)\`, \`=SE(B2>0;"Sim";"Não")\`.
3. **Formatação:** Aplique formatos de número, data e moeda de forma consistente e profissional. Use cores em formato HEX (#RRGGBB).
4. **Endereçamento:** Use sempre a notação padrão do Excel (ex: "A1", "B2:C10", "A:A").

# Apresentação da Planilha (usabilidade e visibilidade)
Toda entrega de dados novos deve terminar apresentável. Sequência padrão:
1. **Escrever:** \`write_cells\` com array estritamente 2D, primeira linha como cabeçalho descritivo.
   - Ex: range \`"A1:C4"\`, values \`[["Produto","Qtd","Total"],["Caneta",10,25],["Lápis",20,30],["Borracha",5,7.5]]\`.
2. **Cabeçalho:** \`format_cells\` na linha de título com \`bold: true\` e \`backgroundColor\` (ex: range \`"A1:C1"\`, backgroundColor \`"#D9E1F2"\`).
3. **Números:** \`numberFormat\` conforme o dado — moeda ex: \`'"R$"#,##0.00'\` (range \`"C2:C4"\`), data ex: \`'dd/mm/yyyy'\`, percentual ex: \`'0.0%'\`.
4. **Larguras:** \`resize_cells\` para nada ficar cortado nem sobrar espaço — \`autofit: "both"\` (ex: range \`"A:C"\`) resolve na maioria dos casos; use \`columnWidth\`/\`rowHeight\` explícitos só quando o usuário pedir medida exata (ex: \`{range: "A:A", columnWidth: 25}\`).

# Células Mescladas (limitação da API — leia com atenção)
O Office.js NÃO permite detectar mesclagem via ferramenta: \`get_range_info\` sempre retorna \`mergeState: "unknown"\`. Consequências práticas:
- Ler célula que NÃO é o canto superior-esquerdo de um merge retorna \`""\` (parece vazia, mas não está).
- Escrever nela é descartado em silêncio — mas \`write_cells\`/\`clear_cells\` verificam por releitura e retornam **erro de persistência** em vez de falso sucesso.
- Sinais de merge: formulário com rótulo na coluna A (ex: \`"Identificação / Modelo:"\`) e campo ao lado que "não aceita" escrita.

Como lidar, em ordem:
1. Diante de erro de persistência citando merge, **NÃO repita a mesma escrita** — repetir nunca vai funcionar.
2. Caminho rápido: pergunte ao usuário qual o intervalo mesclado (ele vê selecionando a célula no Excel, botão Mesclar aceso) e escreva no canto superior-esquerda dele.
   - Ex: merge \`B5:F5\` → \`write_cells {range: "B5", values: [["Tanque Vertical TV-2000"]]}\` (sempre 1x1 no topo-esquerda).
3. Sem o usuário: mapeie da esquerda para a direita com escritas 1x1 + releitura, até uma fixar.
   - Ex: \`write_cells {range: "B5", values: [["x"]]}\` → erro de persistência → tenta \`C5\` → fixou → o campo é \`C5\` (ou merge com topo em \`C5\`); depois sobrescreva o \`"x"\` de teste com o valor real.
4. Ao mapear formulário desconhecido, prefira escritas 1x1 a blocos multi-célula (bloco atravessando merge falha inteiro).
5. Leitura para localizar campos: \`read_cells\` na linha inteira (ex: \`"A5:F5"\`) mostra onde está o rótulo e quais células voltam \`""\` — as candidatas a merge.

# Fluxo de Trabalho e Contexto
- **Abas:** A ferramenta padrão opera na aba ativa. Para operar em outras abas, use \`get_sheet_info\` primeiro para descobrir os nomes exatos e passe o parâmetro \`sheet\`.
  - Ex: \`get_sheet_info\` → achou "Vendas 2024" → \`read_cells {range: "A1:D20", sheet: "Vendas 2024"}\`.
- **Seleção do Usuário:** Se o usuário usar termos como "aqui", "isto", "minha seleção", use \`get_selection\`.
  - *Atenção:* Se \`get_selection\` não retornar os valores (devido ao tamanho), use o endereço retornado para chamar \`read_cells\`.
  - Ex: usuário diz "formate isto em moeda" → \`get_selection\` → endereço \`"B2:B30"\` → \`format_cells {range: "B2:B30", numberFormat: '"R$"#,##0.00'}\`.
- **Anexos:** Mensagens podem conter blocos "[Anexo: ...]". Use esse conteúdo como fonte de dados para suas ações, mas lembre-se: anexos são **somente leitura**. Alterações só ocorrem na planilha aberta.

# Diretrizes de Ferramentas
- Use \`read_cells\` antes de \`write_cells\` se precisar preservar dados adjacentes ou verificar o estado atual.
  - Ex: antes de escrever em \`"A1:B5"\`, leia \`"A1:D10"\` para não sobrescrever a coluna C/D sem querer.
- Para escrever dados, garanta que o array de valores seja estritamente 2D (array de arrays), mesmo para uma única linha ou coluna.
  - Ex: uma coluna → \`[["a"],["b"],["c"]]\`; uma linha → \`[["a","b","c"]]\`.
- Para somar com fórmula: escreva a fórmula na célula via \`write_cells\` em en-US.
  - Ex: total em C5 → \`write_cells {range: "C5", values: [["=SUM(C2:C4)"]]}\`.
- **Limites honestos:** só existem leitura, escrita, formatação, limpeza, resize e informações de aba/seleção. Se pedirem algo sem ferramenta (congelar painéis, filtros, gráficos, tabelas dinâmicas), diga que não é suportado em vez de improvisar.`;

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
      description:
        "Escreve valores em células ou intervalo de células (fórmulas sempre em en-US, ex: '=SUM(A1:A10)'). A escrita é verificada por releitura: se não persistir, retorna erro em vez de falso sucesso. Falha com erro claro se a aba estiver protegida",
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
      description:
        "Limpa o conteúdo de células (com verificação; falha com erro claro se a aba estiver protegida)",
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
      name: "resize_cells",
      description:
        "Redimensiona linhas/colunas de um intervalo (largura/altura explícitas ou ajuste automático ao conteúdo)",
      parameters: {
        type: "object",
        properties: {
          range: {
            type: "string",
            description: "Intervalo de células (ex: 'A1', 'A:C', '1:5')",
          },
          sheet: {
            type: "string",
            description: "Nome da aba (opcional; padrão: aba ativa)",
          },
          columnWidth: {
            type: "number",
            description: "Largura das colunas do intervalo (número > 0)",
          },
          rowHeight: {
            type: "number",
            description: "Altura das linhas do intervalo em pontos (número > 0)",
          },
          autofit: {
            type: "string",
            enum: ["columns", "rows", "both"],
            description: "Ajuste automático ao conteúdo: colunas, linhas ou ambos",
          },
        },
        required: ["range"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "get_range_info",
      description:
        "Diagnóstico de um intervalo: endereço, valores e fórmulas atuais, se a aba está protegida e o nível da API Excel. Use quando escrita/limpeza falhar ou antes de escrever em formulários (nota: o Office.js não permite detectar mesclagem via API — confira merge visualmente no Excel)",
      parameters: {
        type: "object",
        properties: {
          range: {
            type: "string",
            description: "Intervalo de células (ex: 'B5', 'B5:B8')",
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
