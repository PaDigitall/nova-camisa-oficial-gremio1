import crypto from 'node:crypto';

function safeEqual(a, b) {
  const aBuf = Buffer.from(String(a || ''), 'utf8');
  const bBuf = Buffer.from(String(b || ''), 'utf8');
  if (aBuf.length !== bBuf.length || aBuf.length === 0) return false;
  return crypto.timingSafeEqual(aBuf, bBuf);
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return res.status(405).json({ ok: false, error: 'Method not allowed' });
  }

  const expectedToken = String(process.env.ASAAS_WEBHOOK_TOKEN || '').trim();
  const receivedToken = String(req.headers['asaas-access-token'] || '').trim();

  if (!expectedToken) {
    console.error('ASAAS_WEBHOOK_TOKEN não configurado na Vercel.');
    return res.status(500).json({ ok: false, error: 'Webhook not configured' });
  }

  if (!safeEqual(receivedToken, expectedToken)) {
    return res.status(401).json({ ok: false, error: 'Unauthorized' });
  }

  const appsScriptUrl = String(process.env.APPS_SCRIPT_API_URL || '').trim();
  const proxySecret = String(process.env.APPS_SCRIPT_PROXY_SECRET || '').trim();

  if (!appsScriptUrl || !proxySecret) {
    console.error('Ponte Apps Script não configurada na Vercel.');
    return res.status(500).json({ ok: false, error: 'Backend bridge not configured' });
  }

  let event = req.body;

  if (typeof event === 'string') {
    try {
      event = JSON.parse(event);
    } catch {
      return res.status(400).json({ ok: false, error: 'Invalid JSON' });
    }
  }

  if (!event || typeof event !== 'object' || !event.id || !event.event) {
    return res.status(400).json({ ok: false, error: 'Invalid Asaas event' });
  }

  try {
    const response = await fetch(appsScriptUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      body: JSON.stringify({
        secret: proxySecret,
        method: 'processAsaasWebhookEvent',
        args: [event]
      }),
      redirect: 'follow'
    });

    const text = await response.text();

    let data;
    try {
      data = JSON.parse(text);
    } catch {
      console.error('Resposta inválida do Apps Script:', text.slice(0, 500));
      return res.status(502).json({ ok: false, error: 'Invalid backend response' });
    }

    if (!response.ok || !data || data.ok !== true) {
      console.error('Falha ao processar webhook no Apps Script:', data);
      return res.status(502).json({
        ok: false,
        error: data?.error?.message || 'Backend processing failed'
      });
    }

    return res.status(200).json({
      ok: true,
      eventId: event.id,
      event: event.event
    });
  } catch (error) {
    console.error('Erro no webhook Asaas:', error);
    return res.status(500).json({ ok: false, error: 'Webhook processing failed' });
  }
}
