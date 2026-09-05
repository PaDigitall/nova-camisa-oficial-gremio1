import { APP_VERSION, appsScriptUrl, html } from './_config.js';

function googleScriptRunShim() {
  return `<script id="vercel-google-script-run-shim">
  (function(){
    async function rpc(method,args){
      const response=await fetch('/api/rpc',{
        method:'POST',
        headers:{'content-type':'application/json'},
        cache:'no-store',
        body:JSON.stringify({method:method,args:Array.isArray(args)?args:[]})
      });
      let data=null;
      try{data=await response.json();}catch(_){ }
      if(!response.ok||!data||data.ok!==true){
        throw new Error(data&&data.error&&data.error.message?data.error.message:'Falha de comunicação com a loja.');
      }
      return data.result;
    }
    function makeRun(success,failure){
      return new Proxy({}, {get:function(_target,prop){
        if(prop==='withSuccessHandler') return function(fn){return makeRun(fn,failure)};
        if(prop==='withFailureHandler') return function(fn){return makeRun(success,fn)};
        return function(){
          const args=Array.prototype.slice.call(arguments);
          rpc(String(prop),args).then(function(result){
            if(typeof success==='function') success(result);
          }).catch(function(error){
            if(typeof failure==='function') failure(error);
            else console.error(error);
          });
        };
      }});
    }
    window.google=window.google||{};
    window.google.script=window.google.script||{};
    window.google.script.run=makeRun(null,null);
    window.__LOJA_VERCEL__=true;
  })();
  </script>`;
}

function safePage(value) {
  const p = String(value || 'estudantes').toLowerCase();
  return ['estudantes','servidores','modelos','acompanhar','admin'].includes(p) ? p : 'estudantes';
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, c => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  }[c]));
}

export default {
  async fetch(request) {
    const base = appsScriptUrl();
    if (!base) {
      return html('<!doctype html><meta charset="utf-8"><title>Loja Digital do Grêmio</title><h1>Loja temporariamente indisponível</h1><p>A integração com o servidor ainda não foi configurada.</p>', 503);
    }

    try {
      const incomingUrl = new URL(request.url);
      const target = new URL(base);
      target.searchParams.set('p', safePage(incomingUrl.searchParams.get('p')));
      target.searchParams.set('_host', 'vercel');

      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), 20000);
      let upstream;
      try {
        upstream = await fetch(target.toString(), {
          method: 'GET',
          redirect: 'follow',
          headers: {
            'accept': 'text/html',
            'user-agent': `LojaDigitalGremio-Vercel/${APP_VERSION}`
          },
          signal: controller.signal
        });
      } finally {
        clearTimeout(timer);
      }

      let pageHtml = await upstream.text();
      if (!upstream.ok) throw new Error('O Apps Script respondeu com HTTP ' + upstream.status + '.');
      if (/accounts\.google\.com|ServiceLogin|Sign in with Google/i.test(pageHtml)) {
        throw new Error('O Web App do Apps Script está exigindo login do Google.');
      }
      if (!/<html[\s>]/i.test(pageHtml)) throw new Error('O Apps Script não retornou a página da loja.');

      pageHtml = pageHtml.replace(/<base\s+target=["']_top["']\s*\/?\s*>/i, '<base target="_self">');
      pageHtml = pageHtml.split("BOOT.baseUrl || window.location.href.split('?')[0]").join("window.location.origin + window.location.pathname");
      const shim = googleScriptRunShim();
      if (/<\/head>/i.test(pageHtml)) pageHtml = pageHtml.replace(/<\/head>/i, shim + '</head>');
      else pageHtml = shim + pageHtml;

      return html(pageHtml, 200);
    } catch (err) {
      const msg = err?.name === 'AbortError'
        ? 'O servidor da loja demorou demais para responder.'
        : (err?.message || 'Falha ao carregar a loja.');
      return html(`<!doctype html><html lang="pt-BR"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Loja Digital do Grêmio</title><style>body{font-family:Arial,sans-serif;background:#f6f8f8;color:#143b60;margin:0;padding:28px}.box{max-width:720px;margin:8vh auto;background:#fff;border:1px solid #dde3e1;border-radius:22px;padding:28px;box-shadow:0 14px 42px rgba(6,54,103,.11)}h1{color:#063667}.err{background:#fff4eb;color:#87420a;border-radius:14px;padding:14px}</style><div class="box"><h1>Loja Digital do Grêmio</h1><p>Não foi possível carregar a loja neste momento.</p><div class="err">${escapeHtml(msg)}</div><p><small>Diagnóstico técnico: <a href="/api/diagnostico">/api/diagnostico</a></small></p></div></html>`, 502);
    }
  }
};
