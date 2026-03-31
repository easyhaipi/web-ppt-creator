#!/usr/bin/env python3
"""
PPT预览服务器 - PPT Preview Server
启动本地HTTP服务器预览生成的PPT
"""

import http.server
import socketserver
import argparse
import os
import webbrowser
from pathlib import Path


class CustomHandler(http.server.SimpleHTTPRequestHandler):
    """自定义HTTP处理器"""

    def end_headers(self):
        # 添加CORS头，允许跨域访问
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Cache-Control', 'no-cache')
        super().end_headers()

    def log_message(self, format, *args):
        # 自定义日志格式
        print(f"[预览] {self.address_string()} - {format % args}")


def open_browser(port):
    """自动打开浏览器"""
    url = f"http://localhost:{port}"
    print(f"\n🌐 正在打开浏览器: {url}")
    webbrowser.open(url)


def main():
    parser = argparse.ArgumentParser(description="PPT预览服务器")
    parser.add_argument("--html", "-i", default="ppt-output/index.html", help="HTML文件路径")
    parser.add_argument("--port", "-p", type=int, default=8080, help="端口号")
    parser.add_argument("--open", "-o", action="store_true", help="自动打开浏览器")

    args = parser.parse_args()

    # 获取HTML文件目录
    html_path = Path(args.html)
    if not html_path.exists():
        print(f"❌ 错误: 文件不存在 - {args.html}")
        return

    # 切换到HTML所在目录
    os.chdir(html_path.parent)

    # 启动服务器
    PORT = args.port
    Handler = CustomHandler

    print(f"=" * 50)
    print(f"🎨 PPT预览服务器")
    print(f"=" * 50)
    print(f"📁 文件目录: {html_path.parent.absolute()}")
    print(f"🌐 访问地址: http://localhost:{PORT}/{html_path.name}")
    print(f"🛑 按 Ctrl+C 停止服务器")
    print(f"=" * 50)

    # 自动打开浏览器
    if args.open:
        open_browser(PORT)

    with socketserver.TCPServer(("", PORT), Handler) as httpd:
        try:
            print("\n✅ 服务器已启动！\n")
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\n\n👋 服务器已停止")


if __name__ == "__main__":
    main()