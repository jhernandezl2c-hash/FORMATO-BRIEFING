/**
 * Proxy CORS para descargar el Excel RDV desde SharePoint.
 *
 * Cómo desplegarlo:
 *  1. Entra a dash.cloudflare.com  →  Workers & Pages  →  Create  →  Create Worker
 *  2. Ponle un nombre (ej: thsa-excel-proxy)  →  Deploy
 *  3. Click en "Edit code", borra todo y pega ESTE archivo completo  →  Deploy
 *  4. Copia la URL que te queda (ej: https://thsa-excel-proxy.TUUSUARIO.workers.dev)
 *  5. Pega esa URL en el briefing, en "Cargar desde link"  →  campo "URL del Worker"
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
      if (dl.includes('sharepoint.com') || dl.includes('1drv.ms') || dl.includes('-my.sharepoint')) {
        dl += (dl.includes('?') ? '&' : '?') + 'download=1';
      }

      const resp = await fetch(dl, { redirect: 'follow' });
      if (!resp.ok) {
        return new Response('Error al descargar: HTTP ' + resp.status, { status: 502, headers: cors });
      }

      const buf = await resp.arrayBuffer();
      return new Response(buf, {
        headers: {
          ...cors,
          'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          'Cache-Control': 'no-store',
        },
      });
    } catch (e) {
      return new Response('Error: ' + e.message, { status: 500, headers: cors });
    }
  },
};
