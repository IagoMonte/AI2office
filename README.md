# AI2office

**Assistente de IA para o Microsoft Office — manipule seus documentos conversando.**

O AI2office adiciona um painel de chat (task pane) aos aplicativos do Microsoft Office, permitindo ler, escrever e formatar conteúdo por meio de linguagem natural. O projeto começa pelo **Excel** e foi desenhado desde o início para se expandir ao **Word**, **PowerPoint** e **Outlook**.

## Visão geral

- Conecta-se a **qualquer API compatível com OpenAI** (`/chat/completions`) — OpenAI, Azure OpenAI, LM Studio, Ollama, proxies locais etc.
- Loop agêntico de *tool calling*: o modelo decide quais ferramentas chamar, o add-in executa via Office.js e devolve os resultados até a resposta final.
- Configuração flexível: endpoint, chave de API e modelo definidos na própria interface; a lista de modelos é carregada diretamente da API configurada.
- Sem back-end: tudo roda no add-in, e as credenciais ficam apenas no `localStorage` do painel.

## Status por aplicativo

| Aplicativo | Status | Ferramentas |
|---|---|---|
| Excel | ✅ Implementado | Leitura, escrita, formatação, limpeza, info da planilha, seleção atual |
| Word | ✅ Implementado | Leitura, inserção, substituição de seleção, formatação, buscar/substituir, limpar documento, estatísticas, seleção atual |
| PowerPoint | 🔜 Planejado | — |
| Outlook | 🔜 Planejado | — |

## Funcionalidades atuais (Excel)

- Chat com histórico de conversa mantido durante a sessão
- Loop de *tool calling* com múltiplas chamadas por turno
- **Visão de todas as abas**: `get_sheet_info` lista todas as worksheets do arquivo e as ferramentas de intervalo aceitam o parâmetro opcional `sheet` para operar em qualquer aba
- **Consciência de seleção**: a seleção atual do usuário é capturada a cada envio e injetada no prompt do sistema — o assistente entende "isto", "aqui", "minha seleção" sem perguntas extras
- Tool dedicada `get_selection` para buscar endereço, dimensões e valores da seleção (valores limitados a 200 células para proteger o contexto da API)
- Painel de configurações (endpoint, chave, modelo) com validação e carregamento dinâmico de modelos
- Uso de ferramentas visível em tempo real no chat: o grupo expande durante a execução e mostra qual ferramenta está rodando
- **Anexos**: botão 📎 ou **arrastar e soltar** arquivos `.xlsx`, `.xls`, `.csv`, `.txt`, `.pdf` e `.docx` direto no painel; o conteúdo é extraído no próprio add-in e injetado na conversa (limite de 12.000 caracteres por arquivo e 40.000 no total) — permite consultar outros arquivos abertos ou locais sem sair do chat

## Funcionalidades atuais (Word)

- Mesmo loop agêntico do Excel: o host ativo (`Office.context.host`) define o prompt, as ferramentas e o módulo de execução usados no turno
- Consciência de seleção textual: o trecho selecionado é injetado no contexto (limitado a 2.000 caracteres)
- `read_document` retorna até 8.000 caracteres com flag de truncamento para proteger o contexto da API
- Leitura posicional para documentos longos: `get_doc_info` traz um outline de títulos com índice de parágrafo, `read_paragraphs` lê janelas de parágrafos (estilo + truncamento de 500 caracteres cada) e `find_in_document` mapeia ocorrências com número do parágrafo e contexto

## Arquitetura

Fluxo de um turno:

1. Usuário envia mensagem → `agent.js` captura a seleção atual (`describeSelection`) e monta as mensagens (system prompt + histórico + contexto)
2. Requisição POST para `/chat/completions` com as ferramentas registradas
3. Se o modelo responde com `tool_calls`, o add-in executa via Office.js e devolve os resultados como novas mensagens
4. O ciclo se repete até a resposta final, exibida no chat

Estrutura em `src/taskpane/`:

```
src/taskpane/
├── taskpane.js        # entry point: wiring de eventos
├── config.js          # persistência de configurações (localStorage)
├── http.js            # proxyUrl + requestJson (XHR promisificado)
├── state.js           # histórico da conversa e estado de carregamento
├── chatView.js        # renderização de mensagens e indicador de loading
├── settingsView.js    # formulário de configurações
├── modelSelect.js     # fetch/render do seletor de modelos
├── attachments.js     # anexos: extração de texto (xlsx/csv/pdf/docx) e fila pendente
├── agent.js           # sendMessage + loop de tool calling + roteamento por host
└── tools/
    ├── schema.js      # SYSTEM_PROMPT + definição JSON das ferramentas por aplicativo
    ├── excel.js       # execução das ferramentas do Excel via Office.js
    └── word.js        # execução das ferramentas do Word via Office.js
```

A camada `tools/` foi projetada para crescer por aplicativo: cada novo app do Office receberá seu próprio módulo de execução (ex.: `tools/word.js`, `tools/outlook.js`) e entradas correspondentes em `schema.js`, reaproveitando o mesmo loop do `agent.js`.

## Ferramentas disponíveis (Excel)

| Ferramenta | Descrição |
|---|---|
| `read_cells` | Lê valores de um intervalo (ex.: `"B2:D10"`); aceita aba específica via `sheet` |
| `write_cells` | Escreve valores em células a partir de um range inicial; aceita aba específica |
| `format_cells` | Aplica negrito, itálico, cor de fundo, cor da fonte, tamanho; aceita aba específica |
| `clear_cells` | Limpa o conteúdo de um intervalo; aceita aba específica |
| `get_sheet_info` | Lista todas as abas com seus intervalos usados e indica a aba ativa |
| `get_selection` | Retorna endereço, dimensões e valores da seleção atual |

## Ferramentas disponíveis (Word)

| Ferramenta | Descrição |
|---|---|
| `read_document` | Lê o texto do corpo do documento (limitado a 8.000 caracteres) |
| `read_paragraphs` | Lê uma janela de parágrafos por posição (`from` 1-based, `count`; até 50) com estilo e truncamento por parágrafo |
| `find_in_document` | Localiza ocorrências de um texto e retorna o número do parágrafo e contexto de cada uma (até 10) |
| `insert_text` | Insere texto no cursor, no início ou no final do documento |
| `replace_selection` | Substitui o trecho selecionado pelo texto informado |
| `format_selection` | Aplica negrito, itálico, sublinhado, tachado, cor e tamanho na seleção |
| `find_replace` | Localiza e substitui todas as ocorrências de um texto |
| `clear_content` | Limpa todo o conteúdo do documento |
| `get_doc_info` | Retorna contagens (parágrafos, caracteres, palavras) e outline de títulos com índice de parágrafo |
| `get_selection` | Retorna o texto selecionado (limitado a 2.000 caracteres) |

Em documentos longos, o fluxo recomendado para a IA é: `get_doc_info` (mapa dos títulos) → `read_paragraphs` nas regiões relevantes → `find_in_document` para localizar termos; `read_document` fica para documentos curtos.

## Como executar

Pré-requisitos: [Node.js](https://nodejs.org) (LTS) e npm.

```bash
# 1. Instalar dependências
npm install

# 2. Confiar nos certificados de desenvolvimento (uma vez só)
npm run dev-certs

# 3. Iniciar o add-in (build + servidor dev + sideload no Excel)
npm start
```

Ao finalizar a depuração:

```bash
npm run stop
```

Outros comandos úteis:

```bash
npm run build:dev   # build de desenvolvimento
npm run build       # build de produção
npm run lint        # ESLint + Prettier
npm run lint:fix    # correções automáticas de lint
npm run validate    # valida o manifest.xml
```

## Configuração da API

1. Abra o painel do AI2office no Excel
2. Clique em ⚙️ e preencha:
   - **Endpoint**: URL base compatível com OpenAI (ex.: `https://api.openai.com/v1`, `http://localhost:1234/v1`)
   - **Chave de API**: token de acesso (pode ficar vazio em servidores locais)
   - **Modelo**: selecionado automaticamente a partir dos modelos expostos pela API
3. Salve — as credenciais ficam apenas no `localStorage` do painel

> Se você usa um servidor local (LM Studio, Ollama), pode ser necessário habilitar CORS ou usar HTTPS no endpoint, pois o add-in roda em `https://localhost`.

## Stack técnica

- [Office.js](https://learn.microsoft.com/office/dev/add-ins/) — integração com os aplicativos Office
- Webpack 5 + Babel — build e transpilação (ES5 para compatibilidade)
- ESLint (flat config) + Prettier — qualidade e formatação de código
- JavaScript puro, sem frameworks — bundle leve e controle total do DOM

## Roadmap

- [ ] Validação completa em runtime das ferramentas do Excel
- [ ] Novas ferramentas do Excel (gráficos, ordenar/filtrar, tabelas)
- [ ] Novas ferramentas do Word (títulos/estilos, listas, tabelas)
- [ ] Streaming de respostas do modelo
- [x] Módulo de ferramentas para **Word**
- [ ] Módulo de ferramentas para **PowerPoint**
- [ ] Módulo de ferramentas para **Outlook**
