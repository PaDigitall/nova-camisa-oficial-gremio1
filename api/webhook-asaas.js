import crypto from 'node:crypto';
import { APP_VERSION, appsScriptUrl, proxySecret, webhookToken, json } from './_config.js';

function secureEqual(a, b) {
  const aa = Buffer.from(String(a || ''), 'utf8');
  const bb = Buffer.from(String(b || ''), 'utf8');
  return aa.length === bb.length && aa.length > 0 && crypto.timingSafeEqual(aa, bb);
}

async function forwardEvent(event) {
  const url = appsScriptUrl();
  const secret = proxySecret();
  if (!url || !secret) throw new Error('Ponte Apps Script não configurada.');

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 15000);
  try {
    const upstream = await fetch(url, {
      method: 'POST',
      redirect: 'follow',
      headers: {
        'content-type': 'application/json; charset=utf-8',
        'accept': 'application/json',
        'user-agent': `LojaDigitalGremio-AsaasWebhook/${APP_VERSION}`
      },
      body: JSON.stringify({ secret, method: 'processAsaasWebhookEvent', args: [event] }),
      signal: controller.signal
    });
    const text = await upstream.text();
    let data;
    try { data = JSON.parse(text); }
    catch { throw new Error(`Resposta inválida do backend (HTTP ${upstream.status}).`); }
    if (!data || data.ok !== true) throw new Error(data?.error?.message || 'Backend recusou o evento.');
    return data.result;
  } finally {
    clearTimeout(timer);
  }
}

export default {
  async fetch(request) {
    if (request.method !== 'POST') return json({ ok: false }, 405, { allow: 'POST' });

    const expected = webhookToken();
    const received = String(request.headers.get('asaas-access-token') || '').trim();
    if (!expected || !secureEqual(received, expected)) return json({ ok: false }, 401);

    let event = {};
    try { event = await request.json(); } catch {}
    if (!event.id || !event.event) return json({ ok: false, error: 'Evento inválido.' }, 400);

    try {
      const result = await forwardEvent(event);
      return json({ ok: true, result });
    } catch (err) {
      const message = err?.name === 'AbortError' ? 'Timeout do backend.' : (err?.message || 'Erro ao processar webhook.');
      return json({ ok: false, error: message }, 502);
    }
  }
};
