import { APP_VERSION } from './_config.js';

const GOOGLE_SOURCES = (id, width) => [
  // Prioridade máxima: bytes do arquivo original. Isso preserva o canal alpha
  // do PNG e evita o fundo preto introduzido por thumbnails do Drive.
  `https://drive.usercontent.google.com/download?id=${encodeURIComponent(id)}&export=download&confirm=t`,
  `https://drive.google.com/uc?export=download&id=${encodeURIComponent(id)}&confirm=t`,
  `https://drive.google.com/uc?export=view&id=${encodeURIComponent(id)}`,
  // Último recurso: previews/miniaturas geradas pelo Google.
  `https://drive.google.com/thumbnail?id=${encodeURIComponent(id)}&sz=w${width}`,
  `https://lh3.googleusercontent.com/d/${encodeURIComponent(id)}=w${width}`
];

function validDriveId(value) {
  const id = String(value || '').trim();
  return /^[A-Za-z0-9_-]{20,}$/.test(id) ? id : '';
}

function sniffImageMime(bytes, declaredType) {
  const b = new Uint8Array(bytes);
  if (b.length >= 8 &&
      b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47 &&
      b[4] === 0x0d && b[5] === 0x0a && b[6] === 0x1a && b[7] === 0x0a) {
    return 'image/png';
  }
  if (b.length >= 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) {
    return 'image/jpeg';
  }
  if (b.length >= 12 &&
      String.fromCharCode(...b.slice(0,4)) === 'RIFF' &&
      String.fromCharCode(...b.slice(8,12)) === 'WEBP') {
    return 'image/webp';
  }
  const type = String(declaredType || '').toLowerCase();
  return type.startsWith('image/') ? type : '';
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

        const declaredType = String(upstream.headers.get('content-type') || '').toLowerCase();
        if (!upstream.ok) {
          errors.push(`${upstream.status}:${declaredType || 'sem-content-type'}`);
          continue;
        }

        const bytes = await upstream.arrayBuffer();
        if (!bytes.byteLength) {
          errors.push('imagem-vazia');
          continue;
        }

        const type = sniffImageMime(bytes, declaredType);
        if (!type) {
          errors.push(`${upstream.status}:conteudo-nao-imagem:${declaredType || 'sem-content-type'}`);
          continue;
        }

        const sourceHost = new URL(source).hostname;
        if (url.searchParams.get('diag') === '1') {
          return new Response(JSON.stringify({
            ok: true,
            version: APP_VERSION,
            id,
            mime: type,
            bytes: bytes.byteLength,
            source: sourceHost
          }), {
            status: 200,
            headers: {
              'content-type': 'application/json; charset=utf-8',
              'cache-control': 'no-store'
            }
          });
        }

        return new Response(bytes, {
          status: 200,
          headers: {
            'content-type': type,
            'content-length': String(bytes.byteLength),
            'cache-control': 'public, max-age=60, s-maxage=300, stale-while-revalidate=3600',
            'x-content-type-options': 'nosniff',
            'access-control-allow-origin': '*',
            'x-logo-proxy': 'gei-drive',
            'x-logo-source': sourceHost
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
