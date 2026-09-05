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

function parseBody(body) {
  if (!body) return {};
  if (typeof body === 'object') return body;
  try { return JSON.parse(String(body)); } catch (_) { return {}; }
}

async function callAppsScript(method, args) {
  const url = String(process.env.APPS_SCRIPT_API_URL || '').trim();
  const secret = String(process.env.APPS_SCRIPT_PROXY_SECRET || '').trim();
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
        'user-agent': 'LojaDigitalGremio-Vercel/7.2.8'
      },
      body: JSON.stringify({ secret, method, args: Array.isArray(args) ? args : [] }),
      signal: controller.signal
    });
    const text = await upstream.text();
    let data;
    try { data = JSON.parse(text); }
    catch (_) {
      if (/accounts\.google\.com|ServiceLogin|Sign in with Google/i.test(text)) {
        throw new Error('O Web App do Apps Script está exigindo login do Google.');
      }
      throw new Error('O Apps Script retornou uma resposta inesperada.');
    }
    if (!data || data.ok !== true) {
      throw new Error(data && data.error && data.error.message ? data.error.message : 'Erro no backend da loja.');
    }
    return data.result;
  } finally {
    clearTimeout(timer);
  }
}

module.exports = async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store, max-age=0');
  res.setHeader('X-Content-Type-Options', 'nosniff');
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: { message: 'Método não permitido.' } });

  const body = parseBody(req.body);
  const method = String(body.method || '').trim();
  const args = Array.isArray(body.args) ? body.args : [];
  if (!ALLOWED_METHODS.has(method)) {
    return res.status(400).json({ ok: false, error: { message: 'Operação não permitida.' } });
  }

  try {
    const result = await callAppsScript(method, args);
    return res.status(200).json({ ok: true, result });
  } catch (err) {
    const message = err && err.name === 'AbortError'
      ? 'O servidor da loja demorou demais para responder. Tente novamente.'
      : (err && err.message ? err.message : 'Falha de comunicação com o sistema.');
    return res.status(502).json({ ok: false, error: { message } });
  }
};
