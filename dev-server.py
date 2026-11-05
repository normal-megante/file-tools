"""開発中に PC でアプリを確認するための、かんたんなサーバー。

使い方（file-tools フォルダで）:
    python dev-server.py
そのあとブラウザで http://localhost:8140 を開く。止めるときは Ctrl+C。

ふつうの `python -m http.server` との違い:
- ファイルを書き換えたら、再読み込みですぐ反映される（古い内容を覚えておかない）
- Windows でも .js ファイルを正しく「JavaScript」として送る
"""
import functools
import http.server
import os
import sys

PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8140
ROOT = os.path.dirname(os.path.abspath(__file__))


class DevHandler(http.server.SimpleHTTPRequestHandler):
    extensions_map = {
        **http.server.SimpleHTTPRequestHandler.extensions_map,
        ".js": "text/javascript",
        ".mjs": "text/javascript",
        ".css": "text/css",
        ".json": "application/json",
        ".webmanifest": "application/manifest+json",
        ".svg": "image/svg+xml",
    }

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()


if __name__ == "__main__":
    handler = functools.partial(DevHandler, directory=ROOT)
    with http.server.ThreadingHTTPServer(("127.0.0.1", PORT), handler) as server:
        print(f"起動しました: http://localhost:{PORT}  （止めるには Ctrl+C）", flush=True)
        server.serve_forever()
