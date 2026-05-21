/**
 * Proxy CORS para descargar el Excel RDV desde SharePoint.
 *
 * Cómo desplegarlo / actualizarlo:
 *  1. Entra a dash.cloudflare.com  →  Workers & Pages  →  abre "thsa-excel-proxy"
 *  2. Click en "Edit code", borra TODO y pega este archivo completo  →  Deploy
 *
 * Uso:  https://tu-worker.workers.dev/?url=<URL_DE_SHAREPOINT>
 */

// Convierte un link "Compartir" de SharePoint a su URL de descarga directa.
// Ej: https://host/:x:/g/personal/USER/TOKEN
//  →  https://host/personal/USER/_layouts/15/download.aspx?share=TOKEN
function toDownloadUrl(shareUrl) {
  try {
    const u = new URL(shareUrl);
    // OneDrive personal:  /:x:/g/personal/{user}/{token}
    const mp = u.pathname.match(/\/:[a-z]:\/[a-z]\/personal\/([^/]+)\/([^/?]+)/i);
    if (mp) {
      return u.origin + '/personal/' + mp[1] + '/_layouts/15/download.aspx?share=' + mp[2];
    }
    // Sitios de SharePoint:  /:x:/s/{site}/{token}
    const ms = u.pathname.match(/\/:[a-z]:\/[a-z]\/(sites\/[^/]+)\/([^/?]+)/i);
    if (ms) {
      return u.origin + '/' + ms[1] + '/_layouts/15/download.aspx?share=' + ms[2];
    }
  } catch (e) { /* link no parseable */ }
  return null;
}

export default {
  async fetch(request) {
    const cors = {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, OPTIONS',
      'Access-Control-Allow-Headers': '*',
    };

    // Preflight
    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: cors });
    }

    const target = new URL(request.url).searchParams.get('url');
    if (!target) {
      return new Response('Falta el parametro ?url=', { status: 400, headers: cors });
    }

    try {
      // 1) Si es un link "Compartir" de SharePoint, convertirlo a descarga directa.
      // 2) Si no, intentar con download=1 (OneDrive personal, 1drv.ms).
      let dl = toDownloadUrl(target);
      if (!dl) {
        dl = target;
        if ((dl.includes('sharepoint.com') || dl.includes('1drv.ms')) && !/[?&]download=1/.test(dl)) {
          dl += (dl.includes('?') ? '&' : '?') + 'download=1';
        }
      }

      const resp = await fetch(dl, {
        redirect: 'follow',
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
          'Accept': '*/*',
        },
      });

      if (!resp.ok) {
        const body = await resp.text().catch(() => '');
        return new Response(
          'SharePoint respondio HTTP ' + resp.status + '. ' + body.slice(0, 400),
          { status: 502, headers: cors }
        );
      }

      const ct = resp.headers.get('content-type') || '';
      const buf = await resp.arrayBuffer();

      // Si devolvio HTML en vez del archivo, casi seguro es una pagina de login:
      // el link no es publico ("Cualquier persona con el vinculo").
      if (ct.includes('text/html')) {
        return new Response(
          'El link no devolvio un archivo Excel sino una pagina web. ' +
          'Probablemente pide iniciar sesion. En SharePoint usa Compartir y elige ' +
          '"Cualquier persona con el vinculo" (sin requerir inicio de sesion).',
          { status: 502, headers: cors }
        );
      }

      return new Response(buf, {
        headers: {
          ...cors,
          'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          'Cache-Control': 'no-store',
        },
      });
    } catch (e) {
      return new Response('Error en el Worker: ' + e.message, { status: 500, headers: cors });
    }
  },
};
