import { APP_VERSION, appsScriptUrl, proxySecret, webhookToken, json } from './_config.js';

export default {
  async fetch() {
    return json({
      ok: true,
      app: 'Loja Digital do Grêmio',
      version: APP_VERSION,
      node: process.version,
      apiUrlConfigured: Boolean(appsScriptUrl()),
      apiUrlSource: process.env.APPS_SCRIPT_API_URL ? 'environment' : 'embedded-fallback',
      proxySecretConfigured: Boolean(proxySecret()),
      webhookTokenConfigured: Boolean(webhookToken())
    });
  }
};
