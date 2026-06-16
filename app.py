#!/usr/bin/env python3
"""Outlook 取件工具 v5 — 支持 Graph API + IMAP 双协议"""

import json, os, requests, hashlib, urllib.parse, imaplib, email, poplib, smtplib
from email.header import decode_header
from datetime import datetime
from flask import Flask, Response, request, jsonify, send_from_directory
from flask_cors import CORS

app = Flask(__name__)
CORS(app)
BASE = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE, "accounts.json")
CACHE_DIR = os.path.join(BASE, "cache")
TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
GRAPH_API = "https://graph.microsoft.com/v1.0"
SCOPE = "https://graph.microsoft.com/.default offline_access"
API_TOKEN = os.environ.get("OUTLOOK_API_TOKEN", "")  # 环境变量设置token，为空则不验证

os.makedirs(CACHE_DIR, exist_ok=True)


def check_token():
    """验证API token，返回True表示通过"""
    if not API_TOKEN:  # 未设置token则不验证（本地使用）
        return True
    token = request.headers.get("Authorization", "").replace("Bearer ", "")
    if not token:
        token = request.args.get("token", "")
    return token == API_TOKEN


@app.before_request
def before_request():
    # 判断是否本地访问：只看 Host，避免 frp/Caddy 反代从 127.0.0.1 转发导致公网被误判成本地
    host = (request.host or '').split(':', 1)[0]
    is_local = host in ('127.0.0.1', 'localhost', '::1')
    
    # 如果设置了API_TOKEN，且不是本地访问，则需要验证
    if API_TOKEN and not is_local:
        # API接口需要token
        if request.path.startswith("/api/"):
            if not check_token():
                return jsonify({"error": "未授权，请提供有效的API Token"}), 401
        # 非本地访问禁止看网页
        elif not request.path.startswith("/static/"):
            return jsonify({"error": "禁止访问，请使用API接口或本地访问"}), 403


def cache_path(email_addr):
    safe = hashlib.md5(email_addr.encode()).hexdigest()
    return os.path.join(CACHE_DIR, safe + ".json")


def load_cache(email_addr):
    p = cache_path(email_addr)
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_cache(email_addr, mails):
    with open(cache_path(email_addr), "w", encoding="utf-8") as f:
        json.dump(mails, f, indent=2, ensure_ascii=False)


def load_accs():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_accs(a):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(a, f, indent=2, ensure_ascii=False)


def decode_mime_header(header_str):
    """解码 MIME 编码的邮件头"""
    if not header_str:
        return ""
    parts = decode_header(header_str)
    result = []
    for part, charset in parts:
        if isinstance(part, bytes):
            result.append(part.decode(charset or "utf-8", errors="replace"))
        else:
            result.append(part)
    return "".join(result)


def get_email_body(msg):
    """提取邮件正文"""
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            if ct == "text/plain":
                payload = part.get_payload(decode=True)
                charset = part.get_content_charset() or "utf-8"
                body = payload.decode(charset, errors="replace")
                break
            elif ct == "text/html" and not body:
                payload = part.get_payload(decode=True)
                charset = part.get_content_charset() or "utf-8"
                body = payload.decode(charset, errors="replace")
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            charset = msg.get_content_charset() or "utf-8"
            body = payload.decode(charset, errors="replace")
    return body


def fetch_pop3_emails(acc, top=20):
    """通过 POP3 拉取邮件"""
    # 根据邮箱域名自动选择POP3服务器
    email_addr = acc.get("email", "")
    if not acc.get("pop_server"):
        if "@qq.com" in email_addr:
            default_server = "pop.qq.com"
        elif "@126.com" in email_addr:
            default_server = "pop.126.com"
        elif "@yeah.net" in email_addr:
            default_server = "pop.yeah.net"
        elif "@sina.com" in email_addr:
            default_server = "pop.sina.com"
        elif "@gmx" in email_addr:
            default_server = "pop.gmx.com"
        elif "@yahoo" in email_addr:
            default_server = "pop.mail.yahoo.com"
        else:
            default_server = "pop.163.com"
    else:
        default_server = acc["pop_server"]
    server = default_server
    port = acc.get("pop_port", 995)
    
    try:
        pop = poplib.POP3_SSL(server, port)
        pop.user(acc["email"])
        pop.pass_(acc["password"])
        
        count, _ = pop.stat()
        top = min(int(top), count)
        
        emails = []
        for i in range(count, count - top, -1):
            try:
                resp, lines, size = pop.retr(i)
                raw = b"\r\n".join(lines)
                msg = email.message_from_bytes(raw)
                
                subject = decode_mime_header(msg.get("Subject", ""))
                from_addr = decode_mime_header(msg.get("From", ""))
                date_str = msg.get("Date", "")
                body = get_email_body(msg)
                
                from_email = ""
                if "<" in from_addr and ">" in from_addr:
                    from_email = from_addr[from_addr.index("<")+1:from_addr.index(">")]
                else:
                    from_email = from_addr
                
                emails.append({
                    "subject": subject or "(无主题)",
                    "from": from_email,
                    "from_name": from_addr.split("<")[0].strip() if "<" in from_addr else from_addr,
                    "time": date_str,
                    "preview": body[:300] if body else "",
                    "body": body,
                    "body_type": "text",
                    "is_read": True,
                })
            except Exception:
                continue
        
        pop.quit()
        return emails
    except Exception as e:
        return [{"error": str(e)}]


def smtp_server_for(email_addr):
    email_addr = (email_addr or '').lower()
    if '@qq.com' in email_addr:
        return 'smtp.qq.com', 465
    if '@126.com' in email_addr:
        return 'smtp.126.com', 465
    if '@yeah.net' in email_addr:
        return 'smtp.yeah.net', 465
    if '@sina.com' in email_addr:
        return 'smtp.sina.com', 465
    if '@gmx' in email_addr:
        return 'smtp.gmx.com', 465
    if '@yahoo' in email_addr:
        return 'smtp.mail.yahoo.com', 465
    return 'smtp.163.com', 465


def check_pop3_account(acc):
    server = acc.get('pop_server')
    if not server:
        email_addr = acc.get('email', '')
        if '@qq.com' in email_addr:
            server = 'pop.qq.com'
        elif '@126.com' in email_addr:
            server = 'pop.126.com'
        elif '@yeah.net' in email_addr:
            server = 'pop.yeah.net'
        elif '@sina.com' in email_addr:
            server = 'pop.sina.com'
        elif '@gmx' in email_addr:
            server = 'pop.gmx.com'
        elif '@yahoo' in email_addr:
            server = 'pop.mail.yahoo.com'
        else:
            server = 'pop.163.com'
    port = int(acc.get('pop_port', 995))
    pop = poplib.POP3_SSL(server, port, timeout=12)
    pop.user(acc['email'])
    pop.pass_(acc['password'])
    count, _ = pop.stat()
    pop.quit()
    return count


def check_smtp_account(acc):
    server, port = smtp_server_for(acc.get('email', ''))
    smtp = smtplib.SMTP_SSL(server, port, timeout=12)
    smtp.login(acc['email'], acc['password'])
    smtp.quit()
    return True


def fetch_imap_emails(acc, top=20, search=""):
    """通过 IMAP 拉取邮件"""
    server = acc.get("imap_server", "imap.163.com")
    port = acc.get("imap_port", 993)
    
    try:
        mail = imaplib.IMAP4_SSL(server, port)
        mail.login(acc["email"], acc["password"])
        mail.select("INBOX")
        
        if search:
            # IMAP 搜索
            status, data = mail.search(None, f'(OR SUBJECT "{search}" FROM "{search}")')
        else:
            status, data = mail.search(None, "ALL")
        
        if status != "OK":
            mail.logout()
            return []
        
        msg_ids = data[0].split()
        # 取最新的 top 封
        msg_ids = msg_ids[-int(top):]
        msg_ids.reverse()
        
        emails = []
        for mid in msg_ids:
            status, msg_data = mail.fetch(mid, "(RFC822)")
            if status != "OK":
                continue
            raw = msg_data[0][1]
            msg = email.message_from_bytes(raw)
            
            subject = decode_mime_header(msg.get("Subject", ""))
            from_addr = decode_mime_header(msg.get("From", ""))
            date_str = msg.get("Date", "")
            body = get_email_body(msg)
            
            # 提取纯地址
            from_email = ""
            if "<" in from_addr and ">" in from_addr:
                from_email = from_addr[from_addr.index("<")+1:from_addr.index(">")]
            else:
                from_email = from_addr
            
            emails.append({
                "subject": subject or "(无主题)",
                "from": from_email,
                "from_name": from_addr.split("<")[0].strip() if "<" in from_addr else from_addr,
                "time": date_str,
                "preview": body[:300] if body else "",
                "body": body,
                "body_type": "text",
                "is_read": True,  # IMAP 默认标记已读
            })
        
        mail.logout()
        return emails
    except Exception as e:
        return [{"error": str(e)}]


@app.route("/api/cache/<path:email_addr>")
def get_cache(email_addr):
    mails = load_cache(email_addr)
    return jsonify({"cached": True, "count": len(mails), "emails": mails})


@app.route("/api/translate", methods=["POST"])
def api_translate():
    text = request.json.get("text", "")
    if not text.strip():
        return jsonify({"translated": ""})
    text = text[:3000]
    try:
        url = "https://translate.googleapis.com/translate_a/single?client=gtx&sl=en&tl=zh-CN&dt=t&q=" + urllib.parse.quote(text)
        r = requests.get(url, timeout=15)
        result = "".join([s[0] for s in r.json()[0] if s[0]])
        return jsonify({"translated": result})
    except Exception as e:
        return jsonify({"error": str(e), "translated": ""})


@app.route("/")
def index():
    with open(os.path.join(BASE, "index.html"), "r", encoding="utf-8") as f:
        return Response(f.read(), content_type="text/html; charset=utf-8")


@app.route("/api/version")
def get_version():
    # 版本号，每次更新APK时改这里
    return jsonify({"version": "2.2", "apk_url": "https://qj.youyangai.top/static/OutlookMailReader.apk"})


@app.route("/static/<path:filename>")
def serve_static(filename):
    static_dir = os.path.join(BASE, "static")
    return send_from_directory(static_dir, filename)


@app.route("/api/accounts", methods=["GET"])
def get_accs():
    return jsonify([{
        "email": a["email"],
        "type": a.get("type", "graph"),
        "password": a.get("password", ""),
        "client_id": a.get("client_id", ""),
        "imap_server": a.get("imap_server", ""),
    } for a in load_accs()])


@app.route("/api/accounts", methods=["POST"])
def add_accs():
    accs = load_accs()
    exist = {a["email"] for a in accs}
    added = 0
    for a in request.json:
        if a.get("email"):
            # IMAP 账号：需要 password + imap_server
            if a.get("type") == "imap" and a.get("password"):
                if a["email"] not in exist:
                    a.setdefault("imap_server", "imap.163.com")
                    a.setdefault("imap_port", 993)
                    accs.append(a)
                    exist.add(a["email"])
                    added += 1
            # Graph 账号：需要 client_id + refresh_token
            elif a.get("client_id") and a.get("refresh_token"):
                if a["email"] not in exist:
                    accs.append(a)
                    exist.add(a["email"])
                    added += 1
    save_accs(accs)
    return jsonify({"ok": True, "added": added, "total": len(accs)})


@app.route("/api/accounts", methods=["PUT"])
def replace_accs():
    old = {a["email"]: a for a in load_accs()}
    merged = [old[a["email"]] for a in request.json if a.get("email") in old]
    save_accs(merged)
    return jsonify({"ok": True, "total": len(merged)})


@app.route("/api/accounts/<path:email_addr>", methods=["PATCH"])
def update_acc(email_addr):
    accs = load_accs()
    patch_data = request.json or {}
    updated = False
    for a in accs:
        if a.get("email") == email_addr:
            for key in ("email", "type", "password", "client_id", "refresh_token", "imap_server", "imap_port", "pop_server", "pop_port"):
                if key in patch_data:
                    a[key] = patch_data[key]
            updated = True
            break
    if not updated:
        return jsonify({"error": "账号不存在"}), 404
    save_accs(accs)
    return jsonify({"ok": True})


@app.route("/api/accounts/<path:email_addr>", methods=["DELETE"])
def delete_acc(email_addr):
    accs = load_accs()
    new_accs = [a for a in accs if a.get("email") != email_addr]
    if len(new_accs) == len(accs):
        return jsonify({"error": "账号不存在"}), 404
    save_accs(new_accs)
    cache_file = cache_path(email_addr)
    if os.path.exists(cache_file):
        os.remove(cache_file)
    return jsonify({"ok": True, "total": len(new_accs)})


@app.route("/api/accounts/<path:email_addr>/check")
def check_acc(email_addr):
    accs = load_accs()
    acc = next((a for a in accs if a.get("email") == email_addr), None)
    if not acc:
        return jsonify({"ok": False, "status": "missing", "message": "账号不存在"}), 404
    try:
        if acc.get("type") == "imap":
            count = check_pop3_account(acc)
            smtp_ok = False
            try:
                smtp_ok = check_smtp_account(acc)
            except Exception:
                smtp_ok = False
            return jsonify({"ok": True, "status": "ok", "message": f"收件正常，共 {count} 封" + ("，发件正常" if smtp_ok else "，发件未验证"), "count": count, "smtp": smtp_ok})
        else:
            tr = requests.post(TOKEN_URL, data={
                "client_id": acc["client_id"],
                "refresh_token": acc["refresh_token"],
                "grant_type": "refresh_token",
                "scope": SCOPE
            }, timeout=15)
            if tr.status_code == 200:
                return jsonify({"ok": True, "status": "ok", "message": "Graph Token 正常"})
            return jsonify({"ok": False, "status": "bad", "message": f"Token 刷新失败 {tr.status_code}"})
    except Exception as e:
        return jsonify({"ok": False, "status": "bad", "message": str(e)})


@app.route("/api/accounts/check-all")
def check_all_accs():
    accs = load_accs()
    suffix = request.args.get("suffix", "")
    if suffix:
        accs = [a for a in accs if (a.get("email", "").split("@")[-1] == suffix)]
    results = []
    for acc in accs:
        email_addr = acc.get("email", "")
        try:
            if acc.get("type") == "imap":
                count = check_pop3_account(acc)
                smtp_ok = False
                try:
                    smtp_ok = check_smtp_account(acc)
                except Exception:
                    smtp_ok = False
                results.append({"email": email_addr, "ok": True, "status": "ok", "message": f"收件正常 {count} 封" + (" 发件OK" if smtp_ok else ""), "count": count, "smtp": smtp_ok})
            else:
                tr = requests.post(TOKEN_URL, data={
                    "client_id": acc["client_id"],
                    "refresh_token": acc["refresh_token"],
                    "grant_type": "refresh_token",
                    "scope": SCOPE
                }, timeout=15)
                if tr.status_code == 200:
                    results.append({"email": email_addr, "ok": True, "status": "ok", "message": "Graph Token 正常"})
                else:
                    results.append({"email": email_addr, "ok": False, "status": "bad", "message": f"Token 刷新失败 {tr.status_code}"})
        except Exception as e:
            results.append({"email": email_addr, "ok": False, "status": "bad", "message": str(e)})
    bad = [r for r in results if not r.get("ok")]
    return jsonify({"ok": True, "total": len(results), "bad_count": len(bad), "results": results})


@app.route("/api/mail/<path:email_addr>")
def get_mail(email_addr):
    """拉取单个账号的邮件 — 支持 Graph 和 IMAP"""
    accs = load_accs()
    acc = next((a for a in accs if a["email"] == email_addr), None)
    if not acc:
        return jsonify({"error": "账号不存在"}), 404

    top = request.args.get("top", "20")
    search = request.args.get("q", "")
    
    # 判断协议类型
    acc_type = acc.get("type", "graph")
    
    if acc_type == "imap":
        # IMAP 协议
        emails = fetch_imap_emails(acc, top=top, search=search)
        if emails and "error" in emails[0]:
            # IMAP 失败，尝试 POP3
            emails = fetch_pop3_emails(acc, top=top)
        if emails and "error" in emails[0]:
            return jsonify({"error": emails[0]["error"], "emails": []})
        save_cache(email_addr, emails)
        return jsonify({"error": None, "emails": emails, "cached": False})
    else:
        # Graph API 协议
        try:
            tr = requests.post(TOKEN_URL, data={
                "client_id": acc["client_id"],
                "refresh_token": acc["refresh_token"],
                "grant_type": "refresh_token",
                "scope": SCOPE
            }, timeout=15)
            if tr.status_code != 200:
                return jsonify({"error": f"Token 刷新失败 ({tr.status_code})", "emails": []})
            td = tr.json()
        except Exception as e:
            return jsonify({"error": str(e), "emails": []})

        if "refresh_token" in td:
            acc["refresh_token"] = td["refresh_token"]
            save_accs(accs)

        headers = {"Authorization": f"Bearer {td['access_token']}"}
        url = f"{GRAPH_API}/me/messages?$top={top}&$orderby=receivedDateTime desc"
        url += "&$select=subject,from,receivedDateTime,body,bodyPreview,isRead"
        if search:
            url += f'&$search="{search}"'

        try:
            r = requests.get(url, headers=headers, timeout=15)
            if r.status_code != 200:
                return jsonify({"error": f"Graph API {r.status_code}", "emails": []})
            emails = []
            for m in r.json().get("value", []):
                s = m.get("from", {}).get("emailAddress", {})
                b = m.get("body", {})
                emails.append({
                    "subject": m.get("subject", "(无主题)"),
                    "from": s.get("address", ""),
                    "from_name": s.get("name", ""),
                    "time": m.get("receivedDateTime", ""),
                    "preview": m.get("bodyPreview", "")[:300],
                    "body": b.get("content", ""),
                    "body_type": b.get("contentType", "text"),
                    "is_read": m.get("isRead", False),
                })
            save_cache(email_addr, emails)
            return jsonify({"error": None, "emails": emails, "cached": False})
        except Exception as e:
            return jsonify({"error": str(e), "emails": []})


if __name__ == "__main__":
    if not os.path.exists(DATA_FILE):
        old = os.path.join(BASE, "..", "outlook_accounts.json")
        if os.path.exists(old):
            with open(old) as f:
                save_accs(json.load(f))
    print("访问 http://localhost:8877")
    app.run(host="0.0.0.0", port=8877, debug=False)
