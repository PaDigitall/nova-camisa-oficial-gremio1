import { APP_VERSION, appsScriptUrl, proxySecret, json } from './_config.js';

const ALLOWED_METHODS = new Set([
  'getPublicData',
  'getLogosPublicas',
  'getMockupsPublicos',
  'getCommerceCapabilities',
  'simularOpcoesCartao',
  'criarPedido',
  'consultarPedido',
  'adminLogin',
  'adminLogout',
  'adminGetData',
  'adminUpdateSettings',
  'adminGetOrderHistory',
  'adminUpdateOrder',
  'adminResendStatusEmail'
]);

async function parseRequestJson(request) {
  try { return await request.json(); } catch { return {}; }
}

async function callAppsScript(method, args) {
  const url = appsScriptUrl();
  const secret = proxySecret();
  if (!url || !secret) throw new Error('A ponte com o Apps Script ainda não foi configurada na Vercel.');

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 20000);
  try {
    const upstream = await fetch(url, {
      method: 'POST',
      redirect: 'follow',
      headers: {
        'content-type': 'application/json; charset=utf-8',
        'accept': 'application/json',
        'user-agent': `LojaDigitalGremio-Vercel/${APP_VERSION}`
      },
      body: JSON.stringify({ secret, method, args: Array.isArray(args) ? args : [] }),
      signal: controller.signal
    });
    const text = await upstream.text();
    let data;
    try { data = JSON.parse(text); }
    catch {
      if (/accounts\.google\.com|ServiceLogin|Sign in with Google/i.test(text)) {
        throw new Error('O Web App do Apps Script está exigindo login do Google.');
      }
      throw new Error(`O Apps Script retornou resposta inesperada (HTTP ${upstream.status}).`);
    }
    if (!data || data.ok !== true) {
      throw new Error(data?.error?.message || 'Erro no backend da loja.');
    }
    return data.result;
  } finally {
    clearTimeout(timer);
  }
}

export default {
  async fetch(request) {
    if (request.method !== 'POST') {
      return json({ ok: false, error: { message: 'Método não permitido.' } }, 405, { allow: 'POST' });
    }

    const body = await parseRequestJson(request);
    const method = String(body.method || '').trim();
    const args = Array.isArray(body.args) ? body.args : [];
    if (!ALLOWED_METHODS.has(method)) {
      return json({ ok: false, error: { message: 'Operação não permitida.' } }, 400);
    }

    try {
      const result = await callAppsScript(method, args);
      return json({ ok: true, result });
    } catch (err) {
      const message = err?.name === 'AbortError'
        ? 'O servidor da loja demorou demais para responder. Tente novamente.'
        : (err?.message || 'Falha de comunicação com o sistema.');
      return json({ ok: false, error: { message } }, 502);
    }
  }
};
