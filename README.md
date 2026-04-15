# 定时截图 + AI识别 + NapCat QQ私聊（文本推送）

Windows + Python 工具：定时截图（窗口或全屏），调用 OpenAI 兼容视觉模型分析，再把结果以纯文本发送到 QQ 私聊。

## 功能

- 定时截图
- 支持两种截图模式：
  - 指定窗口截图（可精确绑定 `hwnd/pid/class`）
  - 全屏截图（可选所有屏幕）
- 调用 OpenAI 兼容视觉接口分析截图
- 发送结果到 NapCat（OneBot HTTP）
- 自动重试（模型请求、NapCat 超时）
- `config.json` 每轮自动重载（改配置后下一轮生效）

## 文件说明

- `napcat_screenshot_ai.py`：主程序
- `config.py`：窗口选择器（把目标窗口写入 `config.json`）
- `run_config.bat`：一键运行窗口选择器（自动建虚拟环境）
- `run_venv.bat`：一键安装依赖并启动主程序
- `config.json.example`：配置模板

## 快速开始

1. 运行 `run_config.bat`（如果你要“窗口截图”）
2. 复制 `config.json.example` 为 `config.json`，填写参数
3. 运行 `run_venv.bat`

## 配置示例

```json
{
  "napcat_base_url": "http://127.0.0.1:3000",
  "napcat_access_token": "",
  "napcat_send_max_retries": 5,
  "target_qq": 123456789,
  "capture_fullscreen": false,
  "capture_all_screens": true,
  "window_title": "Window Title Here",
  "window_hwnd": null,
  "window_pid": null,
  "window_class": null,
  "interval_minutes": 5,
  "openai_base_url": "https://your-openai-compatible-endpoint/v1",
  "openai_api_key": "sk-xxxxxxxxxxxxxxxx",
  "model": "gpt-4o",
  "prompt": "请基于截图输出简短进度播报..."
}
```

## 关键参数说明

- `capture_fullscreen`
  - `true`：全屏截图（忽略 `window_*`）
  - `false`：按 `window_*` 截指定窗口
- `capture_all_screens`
  - `true`：多显示器全部截图（仅全屏模式生效）
  - `false`：仅主屏截图
- `window_*`
  - 建议由 `run_config.bat` 自动写入，不要手改

## NapCat 配置要点

1. 开启 OneBot HTTP 服务，端口与 `napcat_base_url` 一致
2. 如设置 token，`napcat_access_token` 必须一致
3. 建议 `target_qq` 使用真实好友，不建议给自己号发私聊

## 常见问题

### 1) `Ambiguous window ... matches`

窗口重名冲突。重新运行 `run_config.bat` 绑定一次窗口。

### 2) `retcode=200 ... Timeout`

NapCat/QQ 侧发消息回执超时。常见是目标不可私聊、好友关系问题、或 QQ 端状态异常。

### 3) `Upload failed, 403`

上游视觉接口拒绝图片上传，需在你的网关侧开通视觉能力或更换支持图片的模型路由。

## 手动命令行运行

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe config.py
.\.venv\Scripts\python.exe napcat_screenshot_ai.py
```
