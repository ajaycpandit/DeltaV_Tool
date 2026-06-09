"""
Run this to start the local server:  python serve.py
Then open:  http://localhost:5000
"""
import os, sys

# Try Flask first (full app with Python backend)
try:
    from app import app
    port = int(os.environ.get('PORT', 5000))
    print(f"\n  FHX Converter running at  http://localhost:{port}\n")
    app.run(host='0.0.0.0', port=port, debug=False)
except ImportError:
    # Fallback: plain Python HTTP server (no backend, just serves index.html)
    import http.server, socketserver
    PORT = 5000
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    handler = http.server.SimpleHTTPRequestHandler
    with socketserver.TCPServer(('', PORT), handler) as httpd:
        print(f"\n  FHX Converter running at  http://localhost:{PORT}\n")
        httpd.serve_forever()
