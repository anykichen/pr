# Print Agent

本地打印代理 - 监听 HTTP 请求，将 ZPL 发送到 Zebra 打印机

## 使用方法

1. 下载 Release 中的 `print-agent-win.exe`
2. 双击运行
3. 访问 http://localhost:9100

## API 接口

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/health` | 健康检查 |
| GET | `/printers` | 获取打印机列表 |
| POST | `/print` | 发送打印任务 |

### 打印示例

```javascript
fetch('http://localhost:9100/print', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({
    zpl: '^XA^FO50,50^ADN,36,20^FDHello^FS^XZ',
    printer: 'Zebra ZD410'
  })
});
```

## 环境变量

| 变量 | 默认值 | 说明 |
|------|--------|------|
| PORT | 9100 | 监听端口 |
| PRINTER_NAME | - | 默认打印机名称 |

## 自动构建

每次发布 tag 时自动构建 exe:

```bash
git tag v1.0.0
git push origin v1.0.0
```
