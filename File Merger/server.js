const http = require('http');
const fs = require('fs');
const path = require('path');
const { URL } = require('url');

const PORT = 8787;
const HOST = '127.0.0.1';
const ROOT = __dirname;
const HTML_FILE = path.join(ROOT, 'workload_merge_console_corrected.html');

// Only allow the official UVPCE host. The server is a transport layer only;
// subject verification still reads the original UVPCE pages and no local
// subject-code cache is used.
const ALLOWED_HOSTS = new Set(['uvpce.guni.ac.in']);

function send(res, status, type, body) {
  res.writeHead(status, {
    'Content-Type': type,
    'Cache-Control': 'no-store',
    'Access-Control-Allow-Origin': '*',
  });
  res.end(body);
}

async function fetchOfficial(targetUrl) {
  const u = new URL(targetUrl);

  if (u.protocol !== 'https:' || !ALLOWED_HOSTS.has(u.hostname)) {
    throw new Error('Only https://uvpce.guni.ac.in official URLs are allowed.');
  }

  const response = await fetch(u.href, {
    redirect: 'follow',
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/139 Safari/537.36',
      'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
      'Accept-Language': 'en-US,en;q=0.9',
      'Referer': 'https://uvpce.guni.ac.in/'
    }
  });

  const text = await response.text();

  if (!response.ok) {
    throw new Error(`UVPCE returned HTTP ${response.status}`);
  }

  return text;
}

const server = http.createServer(async (req, res) => {
  try {
    const requestUrl = new URL(req.url, `http://${req.headers.host || `${HOST}:${PORT}`}`);

    if (requestUrl.pathname === '/official') {
      const target = requestUrl.searchParams.get('url');
      if (!target) throw new Error('Missing url parameter.');

      const html = await fetchOfficial(target);
      send(res, 200, 'text/html; charset=utf-8', html);
      return;
    }

    if (requestUrl.pathname === '/' || requestUrl.pathname === '/index.html') {
      const html = fs.readFileSync(HTML_FILE, 'utf8');
      send(res, 200, 'text/html; charset=utf-8', html);
      return;
    }

    send(res, 404, 'text/plain; charset=utf-8', 'Not found');
  } catch (err) {
    send(res, 502, 'text/plain; charset=utf-8', err.message || String(err));
  }
});

server.listen(PORT, HOST, () => {
  console.log(`Workload Merge Console running at http://${HOST}:${PORT}`);
  console.log('B.Tech verification fetches official UVPCE pages server-side to avoid browser CORS blocking.');
  console.log('Open the URL above instead of opening the HTML with file://.');
});
