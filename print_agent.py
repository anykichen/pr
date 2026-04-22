#!/usr/bin/env python3
"""
Zebra Print Agent - Windows 本地打印代理
发送 ZPL 到斑马打印机，支持系统托盘
"""

import sys
import os
import json
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse

import win32print
import pystray
from PIL import Image, ImageDraw

PORT = int(os.environ.get('PORT', 9100))


# ── 获取打印机列表 ────────────────────────────────────────────
def list_printers():
    printers = []
    try:
        for flags, desc, name, comment in win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        ):
            try:
                h = win32print.OpenPrinter(name)
                info = win32print.GetPrinter(h, 2)
                port = info.get('pPortName', name)
                win32print.ClosePrinter(h)
            except Exception:
                port = name
            printers.append({'name': name, 'port': port})
    except Exception as e:
        print(f'获取打印机列表失败: {e}')
    return printers


# ── 发送 ZPL ─────────────────────────────────────────────────
def send_zpl(zpl: str, printer_name: str):
    hPrinter = win32print.OpenPrinter(printer_name)
    try:
        win32print.StartDocPrinter(hPrinter, 1, ('ZPL Label', None, 'RAW'))
        try:
            win32print.StartPagePrinter(hPrinter)
            win32print.WritePrinter(hPrinter, zpl.encode('utf-8'))
            win32print.EndPagePrinter(hPrinter)
        finally:
            win32print.EndDocPrinter(hPrinter)
    finally:
        win32print.ClosePrinter(hPrinter)


# ── HTTP 处理器 ───────────────────────────────────────────────
class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass  # 关闭默认访问日志

    def send_json(self, code, data):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_GET(self):
        path = urlparse(self.path).path
        if path == '/health':
            self.send_json(200, {'status': 'ok', 'platform': 'win32', 'port': PORT})
        elif path == '/printers':
            self.send_json(200, {'printers': list_printers()})
        else:
            self.send_json(404, {'error': 'Not found'})

    def do_POST(self):
        path = urlparse(self.path).path
        if path != '/print':
            self.send_json(404, {'error': 'Not found'})
            return

        try:
            length = int(self.headers.get('Content-Length', 0))
            body = json.loads(self.rfile.read(length).decode('utf-8'))
        except Exception:
            self.send_json(400, {'error': '请求格式错误'})
            return

        zpl = body.get('zpl', '').strip()
        printer = body.get('printer', '').strip()

        if not zpl:
            self.send_json(400, {'error': '缺少 zpl 参数'})
            return
        if not printer:
            self.send_json(400, {'error': '未指定打印机名称'})
            return

        try:
            send_zpl(zpl, printer)
            print(f'✅ 已打印 → {printer} ({len(zpl)} bytes)')
            self.send_json(200, {'success': True})
        except Exception as e:
            print(f'❌ 打印失败: {e}')
            self.send_json(500, {'error': str(e), 'success': False})


# ── 托盘图标 ─────────────────────────────────────────────────
def create_tray_icon():
    """生成一个简单的打印机图标"""
    size = 64
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    # 蓝色背景圆
    d.ellipse([0, 0, size-1, size-1], fill=(37, 99, 235, 255))
    # 打印机机身
    d.rectangle([14, 22, 50, 40], fill='white')
    # 出纸槽
    d.rectangle([18, 38, 46, 52], fill='white')
    # 进纸口
    d.rectangle([22, 14, 42, 24], fill='white')
    return img


def run_server_in_thread(server):
    server.serve_forever()


def main():
    # 启动 HTTP 服务
    try:
        server = HTTPServer(('127.0.0.1', PORT), Handler)
    except OSError as e:
        # 端口被占用时弹出提示
        import ctypes
        ctypes.windll.user32.MessageBoxW(
            0,
            f'端口 {PORT} 已被占用，请关闭其他打印代理实例后重试。\n\n错误: {e}',
            'Zebra Print Agent 启动失败',
            0x10
        )
        sys.exit(1)

    t = threading.Thread(target=run_server_in_thread, args=(server,), daemon=True)
    t.start()

    print(f'🖨️  Zebra Print Agent 已启动')
    print(f'📡 监听: http://localhost:{PORT}')
    printers = list_printers()
    print(f'检测到 {len(printers)} 台打印机:')
    for p in printers:
        print(f'  - {p["name"]} ({p["port"]})')

    # 系统托盘
    def on_quit(icon, item):
        icon.stop()
        server.shutdown()
        os._exit(0)

    icon = pystray.Icon(
        name='ZebraPrintAgent',
        icon=create_tray_icon(),
        title=f'🖨️ Zebra Print Agent  |  端口: {PORT}',
        menu=pystray.Menu(
            pystray.MenuItem('Zebra Print Agent', None, enabled=False),
            pystray.MenuItem(f'监听端口: {PORT}', None, enabled=False),
            pystray.MenuItem(f'已检测到 {len(printers)} 台打印机', None, enabled=False),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem('退出', on_quit),
        )
    )
    icon.run()


if __name__ == '__main__':
    main()
