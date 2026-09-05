import { APP_VERSION, appsScriptUrl, proxySecret, webhookToken, json } from './_config.js';

function shortError(err) {
  if (!err) return 'erro desconhecido';
  if (err.name === 'AbortError') return 'timeout';
  return String(err.message || err).slice(0, 400);
}

async function testGet(url) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 15000);
  try {
    const r = await fetch(url + (url.includes('?') ? '&' : '?') + 'p=estudantes&_diag=1', {
      redirect: 'follow',
      headers: { accept: 'text/html', 'user-agent': `LojaDigitalGremio-Diagnostico/${APP_VERSION}` },
      signal: controller.signal
    });
    const text = await r.text();
    return {
      ok: r.ok && /<html[\s>]/i.test(text) && !/accounts\.google\.com|ServiceLogin|Sign in with Google/i.test(text),
      http: r.status,
      html: /<html[\s>]/i.test(text),
      googleLogin: /accounts\.google\.com|ServiceLogin|Sign in with Google/i.test(text),
      finalHost: (() => { try { return new URL(r.url).host; } catch { return ''; } })()
    };
  } catch (err) {
    return { ok: false, error: shortError(err) };
  } finally {
    clearTimeout(timer);
  }
}

async function testRpc(url, secret) {
  if (!secret) return { ok: false, skipped: true, reason: 'APPS_SCRIPT_PROXY_SECRET ausente' };
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 15000);
  try {
    const r = await fetch(url, {
      method: 'POST',
      redirect: 'follow',
      headers: { 'content-type': 'application/json; charset=utf-8', accept: 'application/json' },
      body: JSON.stringify({ secret, method: 'getCommerceCapabilities', args: [] }),
      signal: controller.signal
    });
    const text = await r.text();
    let data = null;
    try { data = JSON.parse(text); } catch {}
    return {
      ok: Boolean(r.ok && data && data.ok === true),
      http: r.status,
      json: Boolean(data),
      backendAcceptedSecret: Boolean(data && data.ok === true),
      backendMessage: data?.error?.message ? String(data.error.message).slice(0, 300) : ''
    };
  } catch (err) {
    return { ok: false, error: shortError(err) };
  } finally {
    clearTimeout(timer);
  }
}

export default {
  async fetch() {
    const url = appsScriptUrl();
    const secret = proxySecret();
    const [backendGet, backendRpc] = await Promise.all([
      testGet(url),
      testRpc(url, secret)
    ]);

    return json({
      ok: backendGet.ok && backendRpc.ok,
      app: 'Loja Digital do Grêmio',
      version: APP_VERSION,
      node: process.version,
      runtime: 'Vercel Functions / Web Standard ESM',
      config: {
        appsScriptUrlSource: process.env.APPS_SCRIPT_API_URL ? 'environment' : 'embedded-fallback',
        proxySecretConfigured: Boolean(secret),
        webhookTokenConfigured: Boolean(webhookToken())
      },
      backendGet,
      backendRpc
    }, 200);
  }
};
