module.exports = async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store, max-age=0');
  return res.status(200).json({
    ok: true,
    app: 'Loja Digital do Grêmio',
    version: '7.2.8',
    apiUrlConfigured: Boolean(String(process.env.APPS_SCRIPT_API_URL || '').trim()),
    proxySecretConfigured: Boolean(String(process.env.APPS_SCRIPT_PROXY_SECRET || '').trim()),
    webhookTokenConfigured: Boolean(String(process.env.ASAAS_WEBHOOK_TOKEN || '').trim())
  });
};
