#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LAMS 拆票留底 - 邮件代理（纯 Python 标准库，无需 pip install）

职责：接收前端 POST 的「原始上传文件 + 处理结果文件 + 元数据」，
      用配置的 SMTP 账号把两份附件发到管理员邮箱。

⚠️ 安全：SMTP 密码 / 授权码只通过环境变量传入，切勿写进代码或提交到公开仓库。

环境变量：
  SMTP_HOST        SMTP 服务器（默认 smtp.qiye.aliyun.com）
  SMTP_PORT        SSL 端口（默认 465）
  SMTP_USER        发件账号（如 jiajun_cai@hmglog.com）
  SMTP_PASS        发件授权码 / 密码
  RECIPIENT        收件邮箱（管理员）
  ALLOWED_ORIGIN   允许的来源 Origin，逗号分隔（留空 = 允许任意，仅建议内网/测试）
  MAX_MB           最大请求体 MB（默认 15）
  PORT             本服务监听端口（默认 8000）

运行：
  SMTP_HOST=smtp.qiye.aliyun.com SMTP_USER=jiajun_cai@hmglog.com \
  SMTP_PASS='***' RECIPIENT=jiajun_cai@hmglog.com \
  ALLOWED_ORIGIN='https://78ed90c...codebuddy.work,https://cjj202307-creator.github.io' \
  python audit_mail.py
"""
import os
import sys
import json
import ssl
import base64
import smtplib
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

SMTP_HOST = os.environ.get('SMTP_HOST', 'smtp.qiye.aliyun.com')
SMTP_PORT = int(os.environ.get('SMTP_PORT', '465'))
SMTP_USER = os.environ.get('SMTP_USER', '')
SMTP_PASS = os.environ.get('SMTP_PASS', '')
RECIPIENT = os.environ.get('RECIPIENT', '')
ALLOWED = [o.strip() for o in os.environ.get('ALLOWED_ORIGIN', '').split(',') if o.strip()]
MAX_MB = float(os.environ.get('MAX_MB', '15'))
PORT = int(os.environ.get('PORT', '8000'))


def _safe(name):
    return os.path.basename(name or 'file')


def send_email(payload):
    meta = payload.get('meta', {})
    ts = meta.get('time', '?')
    fname = meta.get('fileName', '?')
    rows = meta.get('rowCount', '?')
    tickets = meta.get('ticketCount', '?')
    inc = meta.get('inconsistentCount', '?')
    val = meta.get('validationPassed', '?')

    subject = '[LAMS拆票留底] %s 原始:%s' % (ts, _safe(fname))
    msg = MIMEMultipart()
    msg['From'] = SMTP_USER
    msg['To'] = RECIPIENT
    msg['Subject'] = subject
    body = (
        '业务人员已完成一次拆票，系统自动将留底发送至管理员邮箱。\n\n'
        '时间：%s\n原始文件：%s\n数据行：%s\n分票数：%s\n不一致行：%s\n校验：%s\n\n'
        '（本邮件由 LAMS 拆票系统自动发送，附件含原始上传文件与处理结果文件）'
        % (ts, fname, rows, tickets, inc, val)
    )
    msg.attach(MIMEText(body, 'plain', 'utf-8'))

    for key, namekey in (('originalFile', 'originalName'), ('resultFile', 'resultName')):
        data = payload.get(key)
        if not data:
            continue
        try:
            raw = base64.b64decode(data)
        except Exception:
            continue
        fn = _safe(payload.get(namekey, key))
        part = MIMEApplication(raw)
        part.add_header('Content-Disposition', 'attachment', filename=fn)
        msg.attach(part)

    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=30, context=ctx) as s:
        s.login(SMTP_USER, SMTP_PASS)
        s.sendmail(SMTP_USER, [RECIPIENT], msg.as_string())


class Handler(BaseHTTPRequestHandler):
    def _cors(self):
        origin = self.headers.get('Origin')
        if ALLOWED:
            allow = origin if origin in ALLOWED else ''
        else:
            allow = '*'
        self.send_header('Access-Control-Allow-Origin', allow or 'null')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

    def do_OPTIONS(self):
        self.send_response(204)
        self._cors()
        self.end_headers()

    def do_POST(self):
        if self.path.rstrip('/') not in ('/audit', '/audit/'):
            self.send_response(404)
            self.end_headers()
            return
        origin = self.headers.get('Origin')
        if ALLOWED and origin not in ALLOWED:
            self.send_response(403)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'ok': False, 'error': 'origin denied'}).encode('utf-8'))
            return
        try:
            length = int(self.headers.get('Content-Length', 0))
        except Exception:
            length = 0
        if length > MAX_MB * 1024 * 1024:
            self.send_response(413)
            self.end_headers()
            return
        body = self.rfile.read(length) if length > 0 else b''
        try:
            payload = json.loads(body.decode('utf-8'))
            send_email(payload)
            self.send_response(200)
            self._cors()
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'ok': True}).encode('utf-8'))
        except Exception as e:
            self.send_response(500)
            self._cors()
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'ok': False, 'error': str(e)}).encode('utf-8'))

    def log_message(self, *a):
        pass


if __name__ == '__main__':
    if not SMTP_USER or not SMTP_PASS or not RECIPIENT:
        print('缺少环境变量 SMTP_USER / SMTP_PASS / RECIPIENT，无法启动。')
        sys.exit(1)
    print('audit mail proxy 启动 :%d  收件=%s  允许来源=%s' % (PORT, RECIPIENT, ALLOWED or '任意'))
    ThreadingHTTPServer(('0.0.0.0', PORT), Handler).serve_forever()
