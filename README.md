# 📬 Outlook 取件工具

多邮箱统一取件工具，支持 Microsoft Graph API + IMAP/POP3 双协议。网页端 + Android APP 双端同步。

## ✨ 功能特性

- 🔐 **多邮箱管理** — 同时管理 Outlook、QQ、163、126、Yahoo 等邮箱
- 📧 **双协议支持** — Outlook 走 Graph API，其他邮箱走 IMAP/POP3
- 🔄 **智能缓存** — 首次拉取 100 封，刷新只拉最新 10 封
- 🔍 **邮件搜索** — 按主题、发件人搜索
- 🌐 **网页端** — 浏览器直接访问
- 📱 **Android APP** — 独立 APK，手机取件
- 🔑 **API Token** — 可选鉴权，保护隐私

## 📦 安装

### 1. 克隆项目

```bash
git clone https://github.com/YouYangDaShu/outlook-mail-reader.git
cd outlook-mail-reader
```

### 2. 安装依赖

```bash
pip install flask flask-cors requests
```

### 3. 配置账号

复制示例文件并填入你的邮箱信息：

```bash
cp accounts.example.json accounts.json
```

编辑 `accounts.json`，按格式添加你的邮箱账号。

### 4. 启动服务

```bash
python app.py
```

默认端口 `8877`，访问 `http://localhost:8877`

## 📧 邮箱配置说明

### Outlook / Hotmail（Graph API）

需要先注册 Azure 应用获取 `client_id` 和 `refresh_token`：

1. 访问 [Azure Portal](https://portal.azure.com/)
2. 注册应用，获取 `client_id`
3. 授权 `Mail.Read` 权限
4. 获取 `refresh_token`

```json
{
  "type": "graph",
  "email": "your-email@outlook.com",
  "client_id": "你的Client ID",
  "client_secret": "",
  "refresh_token": "你的Refresh Token"
}
```

### QQ 邮箱（POP3）

1. 登录 QQ 邮箱 → 设置 → 账户
2. 开启 POP3/SMTP 服务
3. 生成授权码

```json
{
  "type": "imap",
  "email": "your-qq@qq.com",
  "password": "你的授权码"
}
```

> QQ 邮箱会自动识别为 `pop.qq.com:995`

### 163 / 126 邮箱（POP3）

1. 登录 163/126 邮箱 → 设置 → POP3/SMTP/IMAP
2. 开启 POP3 服务
3. 生成授权码

```json
{
  "type": "imap",
  "email": "your-163@163.com",
  "password": "你的授权码"
}
```

> 163 邮箱默认使用 `pop.163.com:995`，126 使用 `pop.126.com:995`

### Yahoo 邮箱

```json
{
  "type": "imap",
  "email": "your-email@yahoo.com",
  "password": "你的应用专用密码"
}
```

> Yahoo 使用 `pop.mail.yahoo.com:995`

### 自定义 POP3 服务器

如果自动识别不正确，可以手动指定：

```json
{
  "type": "imap",
  "email": "your-email@example.com",
  "password": "你的密码",
  "pop_server": "pop.example.com",
  "pop_port": 995
}
```

## 🔑 API Token（可选）

设置环境变量启用 API 鉴权：

```bash
export OUTLOOK_API_TOKEN="你的token"
python app.py
```

设置后：
- 网页端无法直接访问（保护隐私）
- API 请求需携带 Token：`Authorization: Bearer 你的token`

## 🌐 网页端

直接浏览器访问 `http://localhost:8877`

## 📱 Android APP

APK 文件位于 `app/` 目录（如果有的话），或自行编译。

APP 是网页壳，所有功能通过网页实现，服务端更新后 APP 刷新即可生效。

### APP 特性

- 独立运行，不依赖浏览器
- 支持多账号切换
- 自动同步网页端数据

## 🛠️ Systemd 服务（可选）

创建服务文件实现开机自启：

```bash
sudo tee /etc/systemd/system/outlook-mail.service > /dev/null <<EOF
[Unit]
Description=Outlook Mail Reader
After=network.target

[Service]
Type=simple
User=$USER
WorkingDirectory=$(pwd)
ExecStart=$(which python3) app.py
Restart=always
RestartSec=5
Environment=OUTLOOK_API_TOKEN=

[Install]
WantedBy=multi-user.target
EOF

sudo systemctl enable --now outlook-mail.service
```

## 📁 项目结构

```
outlook-mail-reader/
├── app.py                  # 后端服务
├── index.html              # 网页前端
├── accounts.example.json   # 账号配置示例
├── accounts.json           # 你的账号配置（不提交）
├── cache/                  # 邮件缓存（不提交）
├── .gitignore
└── README.md
```

## 📄 License

MIT License

## 🤝 贡献

欢迎提交 Issue 和 PR！
