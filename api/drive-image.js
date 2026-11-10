import { APP_VERSION } from './_config.js';

const GOOGLE_SOURCES = (id, width) => [
  `https://drive.google.com/thumbnail?id=${encodeURIComponent(id)}&sz=w${width}`,
  `https://lh3.googleusercontent.com/d/${encodeURIComponent(id)}=w${width}`,
  `https://drive.google.com/uc?export=view&id=${encodeURIComponent(id)}`,
  `https://drive.usercontent.google.com/download?id=${encodeURIComponent(id)}&export=download&confirm=t`
];

function validDriveId(value) {
  const id = String(value || '').trim();
  return /^[A-Za-z0-9_-]{20,}$/.test(id) ? id : '';
}

export default {
  async fetch(request) {
    const url = new URL(request.url);
    const id = validDriveId(url.searchParams.get('id'));
    const width = Math.max(64, Math.min(1200, Number(url.searchParams.get('w') || 384)));

    if (!id) {
      return new Response('ID de imagem inválido.', {
        status: 400,
        headers: { 'content-type': 'text/plain; charset=utf-8', 'cache-control': 'no-store' }
      });
    }

    const errors = [];
    for (const source of GOOGLE_SOURCES(id, width)) {
      try {
        const upstream = await fetch(source, {
          redirect: 'follow',
          headers: {
            accept: 'image/avif,image/webp,image/png,image/jpeg,image/*,*/*;q=0.8',
            'user-agent': `LojaDigitalGremio-ImageProxy/${APP_VERSION}`
          }
        });

        const type = String(upstream.headers.get('content-type') || '').toLowerCase();
        if (!upstream.ok || !type.startsWith('image/')) {
          errors.push(`${upstream.status}:${type || 'sem-content-type'}`);
          continue;
        }

        const bytes = await upstream.arrayBuffer();
        if (!bytes.byteLength) {
          errors.push('imagem-vazia');
          continue;
        }

        return new Response(bytes, {
          status: 200,
          headers: {
            'content-type': type,
            'content-length': String(bytes.byteLength),
            'cache-control': 'public, max-age=300, s-maxage=1800, stale-while-revalidate=86400',
            'x-content-type-options': 'nosniff',
            'access-control-allow-origin': '*',
            'x-logo-proxy': 'gei-drive',
            'x-logo-source': new URL(source).hostname
          }
        });
      } catch (err) {
        errors.push(String(err?.message || err).slice(0, 80));
      }
    }

    return new Response('Não foi possível carregar a imagem pública do Google Drive.', {
      status: 502,
      headers: {
        'content-type': 'text/plain; charset=utf-8',
        'cache-control': 'no-store',
        'x-logo-proxy': 'gei-drive-failed',
        'x-logo-debug': errors.join('|').slice(0, 500)
      }
    });
  }
};
