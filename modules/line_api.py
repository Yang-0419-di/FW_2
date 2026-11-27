# modules/line_api.py
# ===== LINE Messaging API（只用到：驗證 webhook 簽章、回覆訊息、推播訊息）=====
# 金鑰由環境變數提供（不要寫進程式或推上 GitHub）：
#   LINE_CHANNEL_ACCESS_TOKEN：LINE Developers → 你的 Messaging API channel →「Messaging API」分頁 → Channel access token (long-lived)
#   LINE_CHANNEL_SECRET      ：同一個 channel →「Basic settings」分頁 → Channel secret（驗證 webhook 用）
import base64
import hashlib
import hmac
import os

import requests

API = "https://api.line.me/v2/bot/message"
MAX_TEXT = 4900          # LINE 文字訊息上限 5000 字，保留一點空間


def _token():
    t = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "").strip()
    if not t:
        raise RuntimeError("未設定環境變數 LINE_CHANNEL_ACCESS_TOKEN")
    return t


def verify_signature(body_bytes, signature):
    """驗證 LINE 傳來的 X-Line-Signature，確認請求真的來自 LINE。"""
    secret = os.environ.get("LINE_CHANNEL_SECRET", "").strip()
    if not secret or not signature:
        return False
    digest = hmac.new(secret.encode("utf-8"), body_bytes, hashlib.sha256).digest()
    return hmac.compare_digest(base64.b64encode(digest).decode("utf-8"), signature)


def _post(path, payload):
    r = requests.post(f"{API}/{path}", json=payload, timeout=10,
                      headers={"Authorization": f"Bearer {_token()}", "Content-Type": "application/json"})
    if r.status_code != 200:
        raise RuntimeError(f"LINE API {path} 失敗：HTTP {r.status_code} {r.text[:300]}")
    return r


def _texts(text):
    text = text if len(text) <= MAX_TEXT else text[:MAX_TEXT - 20] + "\n…（內容過長已截斷）"
    return [{"type": "text", "text": text}]


def reply(reply_token, text):
    """回覆使用者剛傳的訊息（回覆訊息不計入每月免費則數）。"""
    return _post("reply", {"replyToken": reply_token, "messages": _texts(text)})


def push(user_id, text):
    """主動推播給指定使用者（計入每月訊息則數）。"""
    return _post("push", {"to": user_id, "messages": _texts(text)})
