"""Server lokal sederhana untuk dashboard HPS (fallback saat internet mati).
Jalankan: python server.py [port]   (default 8080)
"""
import sys, os, socket, http.server, socketserver

HERE = os.path.dirname(os.path.abspath(__file__))
PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8080


class H(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=HERE, **kw)

    def end_headers(self):
        # supaya data.js / data.meta.json tidak dicache browser ->update langsung terasa
        self.send_header("Cache-Control", "no-store, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        super().end_headers()

    def log_message(self, fmt, *a):
        pass


def lan_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        return s.getsockname()[0]
    except OSError:
        return "127.0.0.1"
    finally:
        s.close()


if __name__ == "__main__":
    socketserver.TCPServer.allow_reuse_address = True
    ip = lan_ip()
    with socketserver.TCPServer(("0.0.0.0", PORT), H) as httpd:
        print("=" * 52)
        print("  Dashboard HPS - server lokal")
        print("  Komputer ini : http://127.0.0.1:%d/" % PORT)
        print("  Jaringan     : http://%s:%d/" % (ip, PORT))
        print("  (Tutup jendela ini untuk menghentikan server)")
        print("=" * 52)
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            pass
