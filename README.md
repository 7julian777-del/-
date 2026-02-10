# 开单助手

本项目是“销货单”开单工具，支持：
- 本地数据库（SQLite），客户/产品/车辆数据自动补全
- 自动计算重量与金额
- 导出 Word 销货单（按模板风格，尺寸 20cm x 8cm）
- OCR 识别（可选，试验功能）

## 运行（开发）
1. 安装 Python 3.10+（建议 3.11）
2. 安装依赖：

```bash
pip install -r requirements.txt
```

3. 运行：

```bash
python app.py
```

## 打包 EXE（Windows）
```bash
pip install pyinstaller
pyinstaller --noconsole --onefile app.py
```

生成的 `app.exe` 在 `dist` 目录。

## 说明
- 数据库位置：`data/app.db`
- 导出位置：默认 `exports/`（可在“设置”里修改）
- OCR：默认 Tesseract 路径 `F:\OCR\tesseract.exe`（可在“设置”里修改）

## PWA（iOS Safari）
本项目已新增纯前端 PWA 版本（无后端、数据只保存在设备本机），入口文件为 `index.html`。

### 启动方式（仅静态托管）
你需要用任何静态文件服务打开本目录（不需要后端逻辑）：

```bash
python -m http.server 8000
```

然后在 iPhone Safari 打开 `http://<你的电脑IP>:8000/index.html`，点击“分享 → 添加到主屏幕”。

### 数据迁移（从旧 SQLite）
1. 在电脑上运行导出脚本：

```bash
python tools/export_db_json.py
```

2. 得到 `exports/kaidan-export-YYYYMMDD.json`。
3. 在 PWA 的“设置 → 数据迁移与备份”中导入该 JSON。

### 说明
- 数据存储在 iOS 设备的 IndexedDB 中，离线可用。
- Word 导出使用前端生成 `.docx`，iOS 会提示保存到“文件”。
- 如需备份，请在“设置”里导出 JSON。

> iOS 的 PWA 需要 HTTPS 才能注册离线缓存（service worker）。如果用局域网测试，请使用支持 HTTPS 的静态托管（如 GitHub Pages、Netlify 或本地 HTTPS 服务器）。

## Netlify 部署与 AI 代理（解决 iOS 跨域）
PWA 在浏览器里调用 AI 接口会遇到 CORS 限制。推荐使用 Netlify Functions 做代理。

### 1) 在 Netlify 上部署
- 直接把本项目部署到 Netlify（静态站点 + Functions）
- 已包含 `netlify/functions/ai-proxy.js` 与 `netlify.toml`

### 2) 设置环境变量（推荐方式）
在 Netlify 项目设置中新增：
- `AI_UPSTREAM_URL`：你的 AI 接口地址（例如 `https://.../v1/chat/completions`）
- `AI_API_KEY`：你的 API Key

这样前端无需保存 Key，安全性更高。

### 3) 前端设置
- “代理地址（Netlify Functions，可选）”：`/.netlify/functions/ai-proxy`
- “API 地址”：如果已在 Netlify 环境变量设置，则可留空；否则填真实上游地址
- “API Key”：如果已在 Netlify 环境变量设置，则可留空
- 其他如模型与提示词按需填写

### 4) 测试
- 在设置页点击“测试接口”，看到“成功”即可使用 AI 识图
