#!/usr/bin/env node
/**
 * Zebra Print Agent
 * 本地打印代理 - 监听 HTTP 请求，将 ZPL 发送到 USB 斑马打印机
 * 
 * 用法: node print-agent.js
 * 端口: 9100 (可通过环境变量 PORT 修改)
 * 
 * Windows: 打印机路径自动检测，或设置环境变量 PRINTER_PATH=\\.\COM3
 * Linux:   默认 /dev/usb/lp0，或设置环境变量 PRINTER_PATH=/dev/usb/lp0
 */

const http = require('http');
const fs = require('fs');
const os = require('os');
const { exec } = require('child_process');

const PORT = process.env.PORT || 9100;
const platform = os.platform();

// ── 打印机路径检测 ────────────────────────────────────────────
function getDefaultPrinterPath() {
  if (process.env.PRINTER_PATH) return process.env.PRINTER_PATH;
  if (platform === 'win32') return null; // Windows 用打印机名称
  // Linux/Mac
  const candidates = [
    '/dev/usb/lp0', '/dev/usb/lp1',
    '/dev/lp0', '/dev/lp1',
  ];
  for (const p of candidates) {
    if (fs.existsSync(p)) return p;
  }
  return '/dev/usb/lp0'; // fallback
}

// ── 发送 ZPL 到打印机 ─────────────────────────────────────────
function sendZpl(zpl, printerName) {
  return new Promise((resolve, reject) => {
    if (platform === 'win32') {
      // Windows: 用 PowerShell + .NET RawPrinterHelper 发送原始 ZPL
      // 这是 Windows 上发送 ZPL 最可靠的方式，绕过驱动渲染直接发 RAW
      const tmpFile = os.tmpdir() + '\\zpl_' + Date.now() + '.zpl';
      fs.writeFileSync(tmpFile, zpl, 'binary');

      const printer = (printerName || process.env.PRINTER_NAME || '').trim();
      if (!printer) {
        fs.unlinkSync(tmpFile);
        reject(new Error('未指定打印机名称，请在前端选择打印机'));
        return;
      }

      // PowerShell 脚本：用 .NET System.Drawing.Printing 发送 RAW 数据
      const psScript = `
$printerName = '${printer.replace(/'/g, "''")}'
$filePath = '${tmpFile.replace(/\\/g, '\\\\')}'
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.IO;
public class RawPrint {
    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Ansi)]
    public class DOCINFOA { [MarshalAs(UnmanagedType.LPStr)] public string pDocName; [MarshalAs(UnmanagedType.LPStr)] public string pOutputFile; [MarshalAs(UnmanagedType.LPStr)] public string pDataType; }
    [DllImport("winspool.drv", EntryPoint="OpenPrinterA", SetLastError=true, CharSet=CharSet.Ansi)] public static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, out IntPtr hPrinter, IntPtr pd);
    [DllImport("winspool.drv", EntryPoint="ClosePrinter", SetLastError=true)] public static extern bool ClosePrinter(IntPtr hPrinter);
    [DllImport("winspool.drv", EntryPoint="StartDocPrinterA", SetLastError=true, CharSet=CharSet.Ansi)] public static extern bool StartDocPrinter(IntPtr hPrinter, Int32 level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);
    [DllImport("winspool.drv", EntryPoint="EndDocPrinter", SetLastError=true)] public static extern bool EndDocPrinter(IntPtr hPrinter);
    [DllImport("winspool.drv", EntryPoint="StartPagePrinter", SetLastError=true)] public static extern bool StartPagePrinter(IntPtr hPrinter);
    [DllImport("winspool.drv", EntryPoint="EndPagePrinter", SetLastError=true)] public static extern bool EndPagePrinter(IntPtr hPrinter);
    [DllImport("winspool.drv", EntryPoint="WritePrinter", SetLastError=true)] public static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, Int32 dwCount, out Int32 dwWritten);
    public static bool SendFileToPrinter(string szPrinterName, string szFileName) {
        IntPtr hPrinter = new IntPtr(0); DOCINFOA di = new DOCINFOA(); bool bSuccess = false;
        di.pDocName = "ZPL Label"; di.pDataType = "RAW";
        if (OpenPrinter(szPrinterName.Normalize(), out hPrinter, IntPtr.Zero)) {
            if (StartDocPrinter(hPrinter, 1, di)) {
                if (StartPagePrinter(hPrinter)) {
                    byte[] bytes = File.ReadAllBytes(szFileName);
                    IntPtr pUnmanagedBytes = Marshal.AllocCoTaskMem(bytes.Length);
                    Marshal.Copy(bytes, 0, pUnmanagedBytes, bytes.Length);
                    int dwWritten = 0;
                    bSuccess = WritePrinter(hPrinter, pUnmanagedBytes, bytes.Length, out dwWritten);
                    Marshal.FreeCoTaskMem(pUnmanagedBytes);
                    EndPagePrinter(hPrinter);
                }
                EndDocPrinter(hPrinter);
            }
            ClosePrinter(hPrinter);
        }
        return bSuccess;
    }
}
'@
$result = [RawPrint]::SendFileToPrinter($printerName, $filePath)
Remove-Item $filePath -Force -ErrorAction SilentlyContinue
if (-not $result) { exit 1 } else { exit 0 }
`.trim();

      exec(`powershell -NoProfile -Command "${psScript.replace(/"/g, '\\"')}"`,
        { timeout: 15000, maxBuffer: 1024 * 1024 },
        (err, stdout, stderr) => {
          if (err) {
            // Fallback: copy /b 命令直接复制到端口
            const portOrName = printerName;
            const tmpFile2 = os.tmpdir() + '\\zpl_fb_' + Date.now() + '.zpl';
            fs.writeFileSync(tmpFile2, zpl, 'binary');
            exec(`copy /b "${tmpFile2}" "\\\\localhost\\${portOrName}"`, { timeout: 8000 }, (err2) => {
              try { fs.unlinkSync(tmpFile2); } catch {}
              if (err2) reject(new Error(`打印失败: ${err.message}`));
              else resolve();
            });
          } else {
            resolve();
          }
        }
      );
    } else {
      // Linux/Mac: 直接写入设备文件
      const path = printerName || getDefaultPrinterPath();
      fs.writeFile(path, zpl, 'binary', (err) => {
        if (err) reject(new Error(`写入 ${path} 失败: ${err.message}`));
        else resolve();
      });
    }
  });
}

// ── 获取可用打印机列表 ────────────────────────────────────────
function listPrinters() {
  return new Promise((resolve) => {
    if (platform === 'win32') {
      // 用 PowerShell 获取所有已安装打印机（比 wmic 更可靠）
      const ps = `powershell -NoProfile -Command "Get-Printer | Select-Object Name,PortName,DriverName,PrinterStatus | ConvertTo-Json -Compress"`;
      exec(ps, { timeout: 8000 }, (err, stdout) => {
        if (err || !stdout.trim()) {
          // PowerShell 失败时 fallback 到 wmic
          exec('wmic printer get name,portname /format:csv', { timeout: 8000 }, (err2, stdout2) => {
            if (err2 || !stdout2.trim()) { resolve([]); return; }
            const lines = stdout2.trim().split('\n').filter(l => l.trim());
            // wmic csv: Node,PortName,Name
            const printers = lines.slice(1)
              .map(l => {
                const parts = l.trim().split(',');
                if (parts.length < 3) return null;
                const name = parts[parts.length - 1]?.trim();
                const port = parts[parts.length - 2]?.trim();
                return name ? { name, port: port || name, driver: '' } : null;
              })
              .filter(Boolean);
            resolve(printers);
          });
          return;
        }

        try {
          let raw = stdout.trim();
          // PowerShell 返回单个对象时不是数组，统一包成数组
          if (!raw.startsWith('[')) raw = `[${raw}]`;
          const list = JSON.parse(raw);
          const printers = list
            .filter(p => p.Name)
            .map(p => ({
              name: p.Name,
              port: p.PortName || p.Name,
              driver: p.DriverName || '',
              status: p.PrinterStatus === 0 ? 'Ready' : String(p.PrinterStatus)
            }));
          resolve(printers);
        } catch {
          resolve([]);
        }
      });
    } else {
      // Linux: 扫描 USB 设备文件
      const devices = [];
      for (let i = 0; i < 4; i++) {
        const p = `/dev/usb/lp${i}`;
        if (fs.existsSync(p)) devices.push({ name: `USB Printer lp${i}`, port: p });
      }
      for (let i = 0; i < 4; i++) {
        const p = `/dev/lp${i}`;
        if (fs.existsSync(p)) devices.push({ name: `Parallel Printer lp${i}`, port: p });
      }
      if (devices.length === 0) {
        devices.push({ name: 'USB Printer (default)', port: '/dev/usb/lp0' });
      }
      resolve(devices);
    }
  });
}

// ── HTTP 服务器 ───────────────────────────────────────────────
const server = http.createServer(async (req, res) => {
  // CORS - 允许来自任何本地源的请求
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Content-Type', 'application/json');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  const url = new URL(req.url, `http://localhost:${PORT}`);

  // GET /health - 健康检查
  if (req.method === 'GET' && url.pathname === '/health') {
    res.writeHead(200);
    res.end(JSON.stringify({
      status: 'ok',
      platform,
      defaultPrinter: getDefaultPrinterPath(),
      version: '1.0.0'
    }));
    return;
  }

  // GET /printers - 获取打印机列表
  if (req.method === 'GET' && url.pathname === '/printers') {
    const printers = await listPrinters();
    res.writeHead(200);
    res.end(JSON.stringify({ printers }));
    return;
  }

  // POST /print - 打印 ZPL
  if (req.method === 'POST' && url.pathname === '/print') {
    let body = '';
    req.on('data', chunk => { body += chunk; });
    req.on('end', async () => {
      try {
        const { zpl, printer } = JSON.parse(body);
        if (!zpl) {
          res.writeHead(400);
          res.end(JSON.stringify({ error: '缺少 zpl 参数' }));
          return;
        }
        console.log(`[${new Date().toLocaleTimeString()}] 打印 ZPL (${zpl.length} bytes) → ${printer || 'default'}`);
        await sendZpl(zpl, printer);
        res.writeHead(200);
        res.end(JSON.stringify({ success: true }));
      } catch (e) {
        console.error('打印错误:', e.message);
        res.writeHead(500);
        res.end(JSON.stringify({ error: e.message }));
      }
    });
    return;
  }

  res.writeHead(404);
  res.end(JSON.stringify({ error: 'Not found' }));
});

server.listen(PORT, '127.0.0.1', async () => {
  console.log('');
  console.log('🖨️  Zebra Print Agent 已启动');
  console.log(`📡 监听地址: http://localhost:${PORT}`);
  console.log(`💻 系统平台: ${platform}`);
  console.log(`🔌 默认打印机: ${getDefaultPrinterPath()}`);
  console.log('');
  console.log('API 端点:');
  console.log(`  GET  http://localhost:${PORT}/health    - 健康检查`);
  console.log(`  GET  http://localhost:${PORT}/printers  - 打印机列表`);
  console.log(`  POST http://localhost:${PORT}/print     - 发送打印任务`);
  console.log('');
  console.log('等待打印任务...');

  const printers = await listPrinters();
  if (printers.length > 0) {
    console.log('检测到打印机:');
    printers.forEach(p => console.log(`  - ${p.name} (${p.port})`));
  }
});

server.on('error', (e) => {
  if (e.code === 'EADDRINUSE') {
    console.error(`❌ 端口 ${PORT} 已被占用，请关闭其他实例或设置 PORT 环境变量`);
  } else {
    console.error('服务器错误:', e.message);
  }
  process.exit(1);
});
