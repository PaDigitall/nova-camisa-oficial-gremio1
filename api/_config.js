export const APP_VERSION = '7.2.8';
export const DEFAULT_APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyJnfhx5RS1gFnpPnFV9e2PdhQpq1gqoDXo4rOFVyHMjZ0PqisUAEWAyPUPBRNOH0pf/exec';

export function appsScriptUrl() {
  return String(process.env.APPS_SCRIPT_API_URL || DEFAULT_APPS_SCRIPT_URL).trim();
}

export function proxySecret() {
  return String(process.env.APPS_SCRIPT_PROXY_SECRET || '').trim();
}

export function webhookToken() {
  return String(process.env.ASAAS_WEBHOOK_TOKEN || '').trim();
}

export function json(data, status = 200, extraHeaders = {}) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      'content-type': 'application/json; charset=utf-8',
      'cache-control': 'no-store, max-age=0',
      'x-content-type-options': 'nosniff',
      ...extraHeaders
    }
  });
}

export function html(body, status = 200) {
  return new Response(body, {
    status,
    headers: {
      'content-type': 'text/html; charset=utf-8',
      'cache-control': 'no-store, max-age=0',
      'x-content-type-options': 'nosniff',
      'referrer-policy': 'strict-origin-when-cross-origin'
    }
  });
}
