import http.server
import socketserver

PORT = 5000  # Yahan port number change kiya

with socketserver.TCPServer(("", PORT), http.server.SimpleHTTPRequestHandler) as server:
    print(f"Server started at http://localhost:{PORT}")
    server.serve_forever()