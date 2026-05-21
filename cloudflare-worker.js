/**
 * Proxy CORS para descargar el Excel RDV desde SharePoint.
 *
 * Cómo desplegarlo / actualizarlo:
 *  1. Entra a dash.cloudflare.com  →  Workers & Pages  →  abre "thsa-excel-proxy"
 *  2. Click en "Edit code", borra TODO y pega este archivo completo  →  Deploy
 *
 * Uso:  https://tu-worker.workers.dev/?url=<URL_DE_SHAREPOINT>
 */

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
      // Los links "Anyone" de SharePoint / OneDrive necesitan download=1 para dar el archivo crudo
      let dl = target;
      const esSharePoint = dl.includes('sharepoint.com') || dl.includes('1drv.ms');
      if (esSharePoint && !/[?&]download=1/.test(dl)) {
        dl += (dl.includes('?') ? '&' : '?') + 'download=1';
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
