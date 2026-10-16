# Arquitetura

## Stack

JavaScript puro (sem framework), Office.js, Webpack 5 + Babel, ESLint e Prettier. Bibliotecas de apoio: `xlsx`, `pdfjs-dist` e `mammoth` (extração de anexos), `marked` e `dompurify` (renderização segura de markdown nas respostas).

## Fluxo de um turno

1. O usuário envia a mensagem; `agent.js` captura a seleção atual e monta o contexto (system prompt + histórico).
2. POST para `/chat/completions` com as ferramentas do aplicativo ativo.
3. Se vier `tool_calls`, o add-in executa via Office.js e devolve os resultados como novas mensagens.
4. O ciclo repete até a resposta final, exibida no chat.

## Organização de `src/taskpane/`

| Arquivo/pasta | Responsabilidade |
|---|---|
| `taskpane.js` | entry point e eventos |
| `agent.js` | loop de tool calling e roteamento por host |
| `config.js`, `settingsView.js`, `modelSelect.js` | configuração e lista de modelos |
| `http.js` | requisições (XHR promisificado) |
| `attachments.js` | extração de texto dos anexos |
| `chatView.js`, `markdown.js` | renderização do chat |
| `tools/schemas/*.js` | definição JSON das ferramentas por aplicativo |
| `tools/excel.js`, `word.js`, `outlook.js`, `powerpoint/` | execução das ferramentas via Office.js |

## Decisões

- **Roteamento por host:** `Office.context.host` define prompt, ferramentas e módulo de execução do turno, reaproveitando o mesmo loop em todos os aplicativos.
- **Proteção de contexto:** leituras são limitadas (valores da seleção, trechos de parágrafos, anexos) com flag de truncamento, e há leitura posicional para documentos longos.
- **PowerPoint com cache e busca:** o módulo mantém um cache do conteúdo dos slides e ferramentas de busca para o modelo localizar shapes sem reler tudo.
- **Sem back-end:** tudo roda no add-in; nenhuma credencial sai do navegador do painel além da chamada à API configurada.
