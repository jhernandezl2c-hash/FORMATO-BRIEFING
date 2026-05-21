/**
 * Worker del briefing TH:
 *   1. Proxy CORS para descargar el Excel RDV desde SharePoint.
 *   2. Verificación de la contraseña de acceso (endpoint /auth).
 *
 * Cómo desplegarlo / actualizarlo:
 *   1. dash.cloudflare.com  →  Workers & Pages  →  abre "thsa-excel-proxy"
 *   2. Edit code  →  borra TODO  →  pega este archivo  →  Deploy
 *
 * IMPORTANTE — define la contraseña (una sola vez):
 *   En el Worker  →  Settings  →  Variables and Secrets  →  Add
 *   Type: Secret · Name: BRIEFING_PASSWORD · Value: la contraseña que tú elijas
 *   →  Deploy. La contraseña queda solo aquí, nunca en el index.html.
 */

// Convierte un link "Compartir" de SharePoint a su URL de descarga directa.
function toDownloadUrl(shareUrl) {
  try {
    const u = new URL(shareUrl);
    const mp = u.pathname.match(/\/:[a-z]:\/[a-z]\/personal\/([^/]+)\/([^/?]+)/i);
    if (mp) {
      return u.origin + '/personal/' + mp[1] + '/_layouts/15/download.aspx?share=' + mp[2];
    }
    const ms = u.pathname.match(/\/:[a-z]:\/[a-z]\/(sites\/[^/]+)\/([^/?]+)/i);
    if (ms) {
      return u.origin + '/' + ms[1] + '/_layouts/15/download.aspx?share=' + ms[2];
    }
  } catch (e) { /* link no parseable */ }
  return null;
}

export default {
  async fetch(request, env) {
    const cors = {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': '*',
    };

    // Preflight
    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: cors });
    }

    const reqUrl = new URL(request.url);

    // ── 1. Verificación de contraseña (login del briefing) ──
    if (reqUrl.pathname === '/auth') {
      let pass = '';
      try { pass = (await request.json()).pass || ''; } catch (e) { /* sin body */ }
      const real = env && env.BRIEFING_PASSWORD ? env.BRIEFING_PASSWORD : '';
      const ok = real !== '' && pass === real;
      return new Response(JSON.stringify({ ok: ok }), {
        headers: { ...cors, 'Content-Type': 'application/json' },
      });
    }

    // ── 2. Proxy de descarga del Excel desde SharePoint ──
    const target = reqUrl.searchParams.get('url');
    if (!target) {
      return new Response('Falta el parametro ?url=', { status: 400, headers: cors });
    }

    try {
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
