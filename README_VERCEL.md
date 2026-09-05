# Loja Digital do Grêmio — Vercel v7.2.8

Hotfix de compatibilidade com a Vercel atual.

## O que foi corrigido

- Node fixado em `24.x` (sem o aviso `>=20`).
- Funções convertidas para o formato Web Standard/ESM atualmente documentado pela Vercel.
- A URL atual do Apps Script já existe como fallback seguro; `APPS_SCRIPT_API_URL` pode continuar configurada para facilitar futuras trocas.
- Novo endpoint `/api/diagnostico` testa, sem exibir segredo, o carregamento GET do Apps Script e a autenticação RPC da ponte.

## Environment Variables

Obrigatória para a loja funcionar por completo:

- `APPS_SCRIPT_PROXY_SECRET` = o MESMO valor da propriedade `VERCEL_PROXY_SECRET` do Apps Script.

Recomendadas:

- `APPS_SCRIPT_API_URL` = `https://script.google.com/macros/s/AKfycbyJnfhx5RS1gFnpPnFV9e2PdhQpq1gqoDXo4rOFVyHMjZ0PqisUAEWAyPUPBRNOH0pf/exec`
- `ASAAS_WEBHOOK_TOKEN` = token próprio que depois será configurado no webhook do Asaas.

Nunca coloque `ASAAS_API_KEY` na Vercel ou no GitHub.

## Teste após publicar

1. `/api/health` — deve responder JSON e mostrar Node 24.x.
2. `/api/diagnostico` — `backendGet.ok` e `backendRpc.ok` devem ser `true`.
3. `/` — deve abrir a Loja Digital do Grêmio.

Se o deployment aparecer como READY mas `/` falhar, consulte `/api/diagnostico`; ele separa erro de acesso ao Apps Script de erro do segredo da ponte.
