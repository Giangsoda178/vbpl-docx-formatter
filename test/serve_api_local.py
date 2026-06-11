"""Serve the Vercel function api/format.py on a local port so the e2e tests
can run against the exact handler that deploys to Vercel.

Usage:  python test/serve_api_local.py [port]      (default 3002)
Then:   API_BASE=http://127.0.0.1:3002 node test/test_api.mjs
"""

import importlib.util
import sys
from http.server import ThreadingHTTPServer
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location('vercel_format', ROOT / 'api' / 'format.py')
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)

port = int(sys.argv[1]) if len(sys.argv) > 1 else 3002
server = ThreadingHTTPServer(('127.0.0.1', port), module.handler)
print(f'Serving api/format.py handler on http://127.0.0.1:{port}', flush=True)
server.serve_forever()
