// ╔══════════════════════════════════════════════════════════╗
// ║  V O I D T A L K  —  Signaling Relay                    ║
// ║  Cloudflare Worker                                       ║
// ║  Deploy: push to GitHub → connect to Cloudflare Pages   ║
// ╚══════════════════════════════════════════════════════════╝

// Uses Cloudflare's built-in KV store for signals.
// Create a KV namespace called SIGNALS in your Cloudflare dashboard
// and bind it to this worker.

const TTL_SEC = 900; // 15 min

const CORS = {
  'Access-Control-Allow-Origin':  '*',
  'Access-Control-Allow-Methods': 'GET, POST, DELETE, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'Content-Type': 'application/json', ...CORS },
  });
}

export default {
  async fetch(request, env) {
    const url    = new URL(request.url);
    const method = request.method;

    // CORS preflight
    if (method === 'OPTIONS') return new Response(null, { headers: CORS });

    // GET /ping
    if (method === 'GET' && url.pathname === '/ping') {
      return json({ ok: true, result: { msg: 'pong', ts: Date.now() } });
    }

    // POST /write  { room, role, payload }
    if (method === 'POST' && url.pathname === '/write') {
      const { room, role, payload } = await request.json();
      if (!room || !role || !payload)
        return json({ ok: false, error: 'Missing room, role, or payload' }, 400);

      const key = `${sanitize(room)}__${sanitize(role)}`;
      await env.SIGNALS.put(key, payload, { expirationTtl: TTL_SEC });
      return json({ ok: true, result: { written: true } });
    }

    // GET /read?room=X&role=Y
    if (method === 'GET' && url.pathname === '/read') {
      const room = url.searchParams.get('room');
      const role = url.searchParams.get('role');
      if (!room || !role)
        return json({ ok: false, error: 'Missing room or role' }, 400);

      const key     = `${sanitize(room)}__${sanitize(role)}`;
      const payload = await env.SIGNALS.get(key);
      if (!payload)
        return json({ ok: true, result: { payload: null, reason: 'not_found' } });

      return json({ ok: true, result: { payload } });
    }

    // DELETE /clear?room=X
    if (method === 'DELETE' && url.pathname === '/clear') {
      const room = url.searchParams.get('room');
      if (!room) return json({ ok: false, error: 'Missing room' }, 400);

      const r    = sanitize(room);
      const keys = await env.SIGNALS.list({ prefix: r + '__' });
      await Promise.all(keys.keys.map(k => env.SIGNALS.delete(k.name)));
      return json({ ok: true, result: { cleared: true, deleted: keys.keys.length } });
    }

    return json({ ok: false, error: 'Not found' }, 404);
  },
};

function sanitize(str) {
  return String(str).replace(/[^a-zA-Z0-9\-_]/g, '').substring(0, 64);
}
