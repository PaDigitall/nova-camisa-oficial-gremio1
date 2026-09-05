import { APP_VERSION, appsScriptUrl, html } from './_config.js';

function safePage(value) {
  const p = String(value || 'estudantes').toLowerCase();
  return ['estudantes', 'servidores', 'modelos', 'acompanhar', 'admin'].includes(p)
    ? p
    : 'estudantes';
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, c => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  }[c]));
}

function buildAppsScriptUrl(requestUrl, base) {
  const incoming = new URL(requestUrl);
  const target = new URL(base);

  // Preserve public route and useful callback/tracking parameters.
  for (const [key, value] of incoming.searchParams.entries()) {
    if (key === '_vercel_share') continue;
    target.searchParams.set(key, value);
  }

  target.searchParams.set('p', safePage(incoming.searchParams.get('p')));
  target.searchParams.set('_host', 'vercel');
  return target.toString();
}

function shell(targetUrl) {
  const src = escapeHtml(targetUrl);
  return `<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
  <meta name="theme-color" content="#063667">
  <title>Loja Digital do Grêmio — Nova Camisa Oficial</title>
  <style>
    html,body{width:100%;height:100%;margin:0;background:#f6f8f8;overflow:hidden}
    body{font-family:Inter,Arial,sans-serif}
    #app{position:fixed;inset:0;width:100%;height:100%;border:0;background:#f6f8f8;display:block}
  </style>
</head>
<body>
  <iframe id="app" src="${src}" title="Loja Digital do Grêmio" allow="payment *; clipboard-write" referrerpolicy="strict-origin-when-cross-origin"></iframe>
</body>
</html>`;
}

export default {
  async fetch(request) {
    const base = appsScriptUrl();
    if (!base) {
      return html(`<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Loja Digital do Grêmio</title><style>body{font-family:Arial,sans-serif;background:#f6f8f8;color:#143b60;padding:28px}.box{max-width:720px;margin:8vh auto;background:#fff;border:1px solid #dde3e1;border-radius:22px;padding:28px}</style><div class="box"><h1>Loja temporariamente indisponível</h1><p>A integração com o servidor ainda não foi configurada.</p><p><small>Versão ${escapeHtml(APP_VERSION)}</small></p></div>`, 503);
    }

    try {
      const targetUrl = buildAppsScriptUrl(request.url, base);
      return html(shell(targetUrl), 200);
    } catch (err) {
      return html(`<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Loja Digital do Grêmio</title><style>body{font-family:Arial,sans-serif;background:#f6f8f8;color:#143b60;padding:28px}.box{max-width:720px;margin:8vh auto;background:#fff;border:1px solid #dde3e1;border-radius:22px;padding:28px}.err{background:#fff4eb;color:#87420a;border-radius:14px;padding:14px}</style><div class="box"><h1>Loja Digital do Grêmio</h1><p>Não foi possível iniciar a loja.</p><div class="err">${escapeHtml(err?.message || 'Falha inesperada.')}</div><p><small><a href="/api/diagnostico">Abrir diagnóstico técnico</a></small></p></div>`, 500);
    }
  }
};
