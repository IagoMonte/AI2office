const POWERPOINT_SYSTEM_PROMPT = `# Papel
Você é um especialista em edição de apresentações Microsoft PowerPoint.
Seu objetivo é localizar e substituir textos em slides de forma segura.

# Regras de Ouro
1. **Leitura primeiro:** antes de buscar ou substituir, use \`ppt_list_content\` para popular o cache.
2. **Um item por vez:** \`ppt_replace\` substitui um único item do cache; repita a chamada para vários itens.
3. **Grupos e tabelas:** itens do tipo GroupChild exigem desagrupar/editar/reagrupar; itens Table usam linha/coluna.
4. **Cache desatualizado:** se a escrita falhar com "não existe mais", rode \`ppt_revalidate\` e peça o id novamente.

# Fluxo recomendado
- \`ppt_list_content\` → \`ppt_search\` (termo) → mostrar top 5 com id/slide/texto → \`ppt_replace\` no id confirmado.
- Para varrer uma tabela inteira use \`ppt_search_table\` (prefixo do id, ex: "s0_85"); para ver a vizinhança de um item no cache use \`ppt_search_index\` (índice + raio).`;

const POWERPOINT_TOOLS = [
  {
    type: "function",
    function: {
      name: "ppt_list_content",
      description: "Lê todos os textos da apresentação (shapes, tabelas, grupos) e popula o cache",
      parameters: { type: "object", properties: {}, required: [] },
    },
  },
  {
    type: "function",
    function: {
      name: "ppt_search",
      description:
        "Busca no cache por similaridade textual + localidade; retorna top resultados com id",
      parameters: {
        type: "object",
        properties: {
          query: { type: "string", description: "Termo a buscar" },
          top: { type: "number", description: "Quantos resultados retornar (padrão 5)" },
          referenceId: {
            type: "string",
            description: "Id do item de referência para priorizar proximidade",
          },
        },
        required: ["query"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "ppt_search_table",
      description:
        "Retorna todas as células de uma tabela do cache pelo prefixo do id (ex: 's0_85'), ordenadas por linha/coluna",
      parameters: {
        type: "object",
        properties: {
          tableId: {
            type: "string",
            description:
              "Prefixo do id da tabela (ex: 's0_85') ou id de uma célula (puxa a tabela toda)",
          },
        },
        required: ["tableId"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "ppt_search_index",
      description:
        "Retorna a vizinhança de um item no cache pela posição (índice ± raio, padrão 3)",
      parameters: {
        type: "object",
        properties: {
          index: { type: "number", description: "Posição do item no cache (0-based)" },
          radius: { type: "number", description: "Raio da vizinhança (padrão 3)" },
        },
        required: ["index"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "ppt_replace",
      description: "Substitui o texto de um único item do cache pelo id",
      parameters: {
        type: "object",
        properties: {
          id: { type: "string", description: "Id do item (ex: s1_shape42 ou s1_shape42_r0c1)" },
          text: { type: "string", description: "Novo texto" },
        },
        required: ["id", "text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "ppt_revalidate",
      description:
        "Revalida o cache de forma incremental (remove deletadas, lê novas, atualiza textos)",
      parameters: { type: "object", properties: {}, required: [] },
    },
  },
  {
    type: "function",
    function: {
      name: "ppt_api_support",
      description: "Verifica suporte a PowerPoint API 1.8 (necessária para grupos)",
      parameters: { type: "object", properties: {}, required: [] },
    },
  },
];

export { POWERPOINT_SYSTEM_PROMPT, POWERPOINT_TOOLS };
