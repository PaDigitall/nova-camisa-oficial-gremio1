# GitHub ↔ Google Apps Script — Loja Digital do Grêmio

Este repositório está vinculado ao projeto oficial do Google Apps Script usado pela planilha **Sistema de Encomendas — Nova Camisa Oficial — Grêmio Estudantil**.

## Identificadores oficiais

- Script ID: `1WzyK0YkQidg6A2RsYFOmS9pLl0DQ8_K4DDzfQQJJzuHNs3drJT4LGIMb`
- Web App / Deployment ID: `AKfycbyJnfhx5RS1gFnpPnFV9e2PdhQpq1gqoDXo4rOFVyHMjZ0PqisUAEWAyPUPBRNOH0pf`
- Branch de produção: `main`
- Código do Apps Script no GitHub: `apps-script/`

## Arquitetura

`GitHub main → GitHub Actions → clasp push → Google Apps Script → atualização da implantação existente → mesma URL /exec → Vercel`

O deploy atualiza a implantação existente. Portanto, a URL utilizada pela Vercel não precisa ser trocada a cada versão.

## Bootstrap inicial

Antes do primeiro deploy automático, execute o workflow **Bootstrap Apps Script source**. Ele usa `clasp pull` para importar para `apps-script/` a versão que está atualmente no Google Apps Script. Isso evita substituir o projeto por uma cópia antiga.

O bootstrap deve ser executado apenas depois de configurar o secret `CLASPRC_JSON`.

## Secret obrigatório: CLASPRC_JSON

1. Em um computador autenticado na conta Google que possui acesso ao projeto, instale/use o clasp 3.x.
2. Execute `npx @google/clasp@3 login`.
3. Conclua a autorização Google.
4. Localize o arquivo `~/.clasprc.json` criado pelo clasp.
5. No GitHub, abra **Settings → Secrets and variables → Actions → New repository secret**.
6. Nome: `CLASPRC_JSON`.
7. Valor: conteúdo completo de `~/.clasprc.json`.

Nunca faça commit de `.clasprc.json`. O `.gitignore` do repositório já bloqueia esse arquivo.

## Fluxo depois do bootstrap

1. Alterar `apps-script/Code.gs`, `apps-script/Index.html` ou outro arquivo do Apps Script no GitHub.
2. Fazer commit na branch `main`.
3. O workflow **Deploy Apps Script** executa automaticamente.
4. O código é enviado ao projeto oficial com `clasp push`.
5. A implantação existente é atualizada, mantendo a mesma URL pública.

## Regra operacional

Depois do bootstrap, o **GitHub é a fonte oficial do código**. Evite editar `Code.gs` ou `Index.html` diretamente no editor do Apps Script, porque isso cria divergência entre produção e repositório.

A planilha continua sendo usada normalmente para dados operacionais, pedidos, configurações, produtos e relatórios. O que passa a ser controlado pelo GitHub é o código do Apps Script.
