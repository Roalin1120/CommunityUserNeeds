# CommunityUserNeeds

版本：`V1.0.0`

本地 Web 后台：导入社群聊天导出文件，筛选原始消息，并用 Gemini 做用户需求 / 商机分析。

## 功能概览

详见 [功能清单.md](./功能清单.md)。

核心能力：

- 原始消息预览与 CSV 导出
- 用户数据分析（Gemini 结构化结果 + Excel 导出）
- 原文件**多选**，支持 `.txt` / `.html` / `.json`
- Telegram HTML 导出解析、WhatsApp 文本解析、常见 JSON 消息结构解析

## 目录约定

- 固定导出目录：`data/exports/`
- 分析库：`data/analysis.db`（本地运行生成，默认不入库）

## 运行

1. 安装依赖

```bash
pip install -r requirements.txt
```

2. 配置环境变量

- `GEMINI_API_KEY`：Gemini API Key（分析功能必需）
- `OPENCLAW_AUTH_TOKEN`：可选，OpenClaw 上传分析接口鉴权

3. 启动服务

```bash
python app.py
```

或双击 `start.bat` / `launch_with_browser.bat`。

4. 访问

- 分析后台：http://127.0.0.1:5000/analysis
- 原始导出：http://127.0.0.1:5000/raw

## 支持的导出格式

- WhatsApp `.txt`
- Telegram Desktop `.html` / `.htm`
- 常见聊天 `.json`（消息数组或 `messages` / `chats` 嵌套）

请将导出文件放入 `data/exports/`，或通过「导入」上传。

## 版本

当前发布版本：`V1.0.0`（见 `VERSION` 与 Git tag）。
