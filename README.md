# 📬 Outlook 取件工具

多邮箱统一取件工具，支持 Microsoft Graph API + IMAP/POP3 双协议。网页端 + Android APP 双端同步。

## ✨ 功能特性

- 🔐 **多邮箱管理** — 同时管理 Outlook、QQ、163、126、Yahoo 等邮箱
- 📧 **双协议支持** — Outlook 走 Graph API，其他邮箱走 IMAP/POP3
- 🔄 **智能缓存** — 首次拉取 100 封，刷新只拉最新 10 封
- 🔍 **邮件搜索** — 按主题、发件人搜索
- 🌐 **网页端** — 浏览器直接访问
- 📱 **Android APP** — 独立 APK，手机取件（不依赖电脑）
- 🔑 **API Token** — 可选鉴权，保护隐私

---

## 🚀 快速开始（3 分钟搞定）

### 第一步：安装

```bash
# 克隆项目
git clone https://github.com/YouYangDaShu/outlook-mail-reader.git
cd outlook-mail-reader

# 安装依赖
pip install flask flask-cors requests
```

### 第二步：配置账号

```bash
# 复制示例文件
cp accounts.example.json accounts.json
```

编辑 `accounts.json`，添加你的邮箱。格式如下：

```json
[
  {
    "type": "imap",
    "email": "你的QQ邮箱@qq.com",
    "password": "你的POP3授权码"
  }
]
```

### 第三步：启动

```bash
python app.py
```

打开浏览器访问 `http://localhost:8877` ✅

---

## 📧 邮箱配置详细教程

### 1️⃣ QQ 邮箱（最简单，推荐新手）

**获取授权码步骤：**

1. 登录 [QQ邮箱](https://mail.qq.com/)
2. 点击 **设置** → **账户**
3. 往下找 **POP3/IMAP/SMTP/Exchange/CardDAV/CalDAV服务**
4. 开启 **POP3/SMTP服务**
5. 按提示用手机发短信，获取 **授权码**（16位字母）

**配置：**

```json
{
  "type": "imap",
  "email": "123456789@qq.com",
  "password": "abcdefghijklmnop"
}
```

> ✅ 会自动识别为 `pop.qq.com:995`，不用手动填服务器

---

### 2️⃣ 163 / 126 邮箱

**获取授权码步骤：**

1. 登录 [163邮箱](https://mail.163.com/) 或 [126邮箱](https://mail.126.com/)
2. 点击 **设置** → **POP3/SMTP/IMAP**
3. 开启 **POP3/SMTP服务**
4. 按提示用手机发短信，获取 **授权码**

**配置：**

```json
{
  "type": "imap",
  "email": "yourname@163.com",
  "password": "你的授权码"
}
```

> ✅ 163 自动识别为 `pop.163.com:995`，126 自动识别为 `pop.126.com:995`

---

### 3️⃣ Outlook / Hotmail（稍微复杂）

Outlook 需要通过 Microsoft Graph API，需要注册 Azure 应用。

**详细步骤：**

#### 第一步：注册 Azure 应用

1. 打开 [Azure Portal](https://portal.azure.com/)
2. 登录你的 Outlook 账号
3. 搜索 **Azure Active Directory** 或 **Microsoft Entra ID**
4. 左侧菜单点击 **应用注册** → **新注册**
5. 填写：
   - 名称：`MailReader`（随便写）
   - 受支持的账户类型：**任何组织目录中的账户和个人 Microsoft 账户**
   - 重定向 URI：`https://login.microsoftonline.com/common/oauth2/nativeclient`
6. 点击 **注册**

#### 第二步：获取 Client ID

1. 注册完成后，复制 **应用程序(客户端) ID**
2. 这就是你的 `client_id`

#### 第三步：获取 Refresh Token

1. 在 Azure 应用页面，点击 **API权限** → **添加权限**
2. 选择 **Microsoft Graph** → **委托的权限**
3. 搜索并添加 **Mail.Read** 权限
4. 打开浏览器，访问以下地址（替换 `YOUR_CLIENT_ID`）：

```
https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=YOUR_CLIENT_ID&response_type=code&redirect_uri=https://login.microsoftonline.com/common/oauth2/nativeclient&scope=Mail.Read
```

5. 登录并授权，浏览器会跳转到一个空白页面
6. 复制地址栏中的 `code` 参数（`code=` 后面的长字符串）

7. 用 curl 获取 refresh_token（替换参数）：

```bash
curl -X POST https://login.microsoftonline.com/common/oauth2/v2.0/token \
  -d "client_id=YOUR_CLIENT_ID" \
  -d "code=YOUR_CODE" \
  -d "redirect_uri=https://login.microsoftonline.com/common/oauth2/nativeclient" \
  -d "grant_type=authorization_code" \
  -d "scope=Mail.Read"
```

8. 返回的 JSON 中有 `refresh_token`，复制保存

**配置：**

```json
{
  "type": "graph",
  "email": "yourname@outlook.com",
  "client_id": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "client_secret": "",
  "refresh_token": "M.C547_BAY.xxxxx..."
}
```

---

### 4️⃣ Yahoo 邮箱

1. 登录 Yahoo 邮箱
2. 进入 **账户安全** → **生成应用专用密码**
3. 选择 **其他**，输入名称，生成密码

```json
{
  "type": "imap",
  "email": "yourname@yahoo.com",
  "password": "应用专用密码"
}
```

---

### 5️⃣ 自定义 POP3 服务器

如果自动识别不正确，可以手动指定：

```json
{
  "type": "imap",
  "email": "yourname@example.com",
  "password": "你的密码",
  "pop_server": "pop.example.com",
  "pop_port": 995
}
```

---

## 📱 Android APP 使用

### 方式一：使用预编译 APK

1. 下载 `static/Outlook取件-独立版.apk`
2. 安装到手机
3. 打开 APP，点击右上角 **设置**
4. 导入 `accounts.json` 文件
5. 返回主页，点击 **刷新**

### 方式二：使用网页版 APP

1. 在电脑上启动服务 `python app.py`
2. 手机浏览器访问 `http://电脑IP:8877`
3. 添加到主屏幕（浏览器菜单 → 添加到主屏幕）

---

## 🔑 API Token（可选，保护隐私）

如果不想让别人访问你的取件页面，可以设置 Token：

```bash
# Linux/Mac
export OUTLOOK_API_TOKEN="设置一个密码"
python app.py

# Windows
set OUTLOOK_API_TOKEN=设置一个密码
python app.py
```

设置后：
- 网页端会提示输入 Token
- API 请求需要携带：`Authorization: Bearer 你的token`

---

## 🛠️ 设置开机自启（Linux Systemd）

```bash
# 创建服务文件
sudo tee /etc/systemd/system/outlook-mail.service > /dev/null <<EOF
[Unit]
Description=Outlook Mail Reader
After=network.target

[Service]
Type=simple
User=$(whoami)
WorkingDirectory=$(pwd)
ExecStart=$(which python3) app.py
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
EOF

# 启用并启动
sudo systemctl enable --now outlook-mail.service

# 查看状态
sudo systemctl status outlook-mail.service
```

---

## 📁 项目结构

```
outlook-mail-reader/
├── app.py                  # 后端服务
├── index.html              # 网页前端
├── accounts.example.json   # 账号配置示例
├── accounts.json           # 你的账号配置（不提交到 Git）
├── cache/                  # 邮件缓存（不提交到 Git）
├── static/
│   ├── Outlook取件.apk     # 网页壳 APP
│   └── Outlook取件-独立版.apk  # 独立运行 APP
├── .gitignore
└── README.md
```

---

## ❓ 常见问题

### Q: QQ邮箱收不到邮件？
A: 确认已开启 POP3 服务，并使用授权码（不是QQ密码）。

### Q: Outlook 报错 "invalid_client"？
A: 检查 `client_id` 是否正确，redirect_uri 必须是 `https://login.microsoftonline.com/common/oauth2/nativeclient`

### Q: 网页打不开？
A: 检查端口 8877 是否被占用，或修改 `app.py` 中的端口号。

### Q: 手机 APP 连不上？
A: 确保手机和电脑在同一局域网，用电脑的局域网 IP 访问。

### Q: 怎么获取电脑的局域网 IP？
A: 
- Windows: `ipconfig`
- Linux/Mac: `ifconfig` 或 `ip addr`

---

## 📄 License

MIT License

## 🤝 贡献

欢迎提交 Issue 和 PR！
