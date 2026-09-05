# Publicação na Vercel — Loja Digital do Grêmio v7.2.8

Esta versão acrescenta uma rota de segurança em `index.html` para evitar tela vazia na raiz e deixa as variáveis documentadas de forma separada.

## Diagnóstico rápido
- `404 NOT_FOUND` da própria Vercel: arquivos/rotas não foram implantados. Não é erro das chaves.
- `/api/health` abre, mas algum `...Configured` está `false`: variável de ambiente faltando ou com nome incorreto.
- `/api/health` mostra tudo `true`, mas a loja não abre: confira `APPS_SCRIPT_API_URL` e se o Web App do Apps Script está público.

Leia `CONFIGURACAO_VERCEL_SEM_CONFUSAO.txt`.
