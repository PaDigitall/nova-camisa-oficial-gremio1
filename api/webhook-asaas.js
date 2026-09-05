const crypto = require('crypto');

function parseBody(body) {
  if (body && typeof body === 'object') return body;
  try { return JSON.parse(String(body || '{}')); } catch (_) { return {}; }
}

function secureEqual(a, b) {
  const aa = Buffer.from(String(a || ''), 'utf8');
  const bb = Buffer.from(String(b || ''), 'utf8');
  return aa.length === bb.length && aa.length > 0 && crypto.timingSafeEqual(aa, bb);
}

async function forwardEvent(event) {
  const url = String(process.env.APPS_SCRIPT_API_URL || '').trim();
  const secret = String(process.env.APPS_SCRIPT_PROXY_SECRET || '').trim();
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
        'user-agent': 'LojaDigitalGremio-AsaasWebhook/7.2.8'
      },
      body: JSON.stringify({ secret, method: 'processAsaasWebhookEvent', args: [event] }),
      signal: controller.signal
    });
    const text = await upstream.text();
    let data;
    try { data = JSON.parse(text); } catch (_) { throw new Error('Resposta inválida do backend.'); }
    if (!data || data.ok !== true) throw new Error(data && data.error && data.error.message ? data.error.message : 'Backend recusou o evento.');
    return data.result;
  } finally {
    clearTimeout(timer);
  }
}

module.exports = async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store, max-age=0');
  if (req.method !== 'POST') return res.status(405).json({ ok: false });

  const expected = String(process.env.ASAAS_WEBHOOK_TOKEN || '').trim();
  const received = String(req.headers['asaas-access-token'] || '').trim();
  if (!expected || !secureEqual(received, expected)) return res.status(401).json({ ok: false });

  const event = parseBody(req.body);
  if (!event.id || !event.event) return res.status(400).json({ ok: false, error: 'Evento inválido.' });

  try {
    const result = await forwardEvent(event);
    return res.status(200).json({ ok: true, result });
  } catch (err) {
    const message = err && err.name === 'AbortError' ? 'Timeout do backend.' : (err && err.message ? err.message : 'Erro ao processar webhook.');
    return res.status(502).json({ ok: false, error: message });
  }
};
