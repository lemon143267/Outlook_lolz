import logging
import os
import re
import secrets
import sqlite3
from contextlib import asynccontextmanager
from typing import List, Optional

import httpx
from fastapi import Depends, FastAPI, HTTPException, status
from fastapi.responses import HTMLResponse
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from pydantic import BaseModel

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

DB_NAME = os.getenv("DB_NAME", "accounts.db")
ADMIN_USERNAME = os.getenv("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "admin")

security = HTTPBasic()


def authenticate(credentials: HTTPBasicCredentials = Depends(security)):
    correct_username = secrets.compare_digest(credentials.username, ADMIN_USERNAME)
    correct_password = secrets.compare_digest(credentials.password, ADMIN_PASSWORD)
    if not (correct_username and correct_password):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Неверный логин или пароль",
            headers={"WWW-Authenticate": "Basic"},
        )
    return credentials.username


def get_db():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS accounts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email TEXT NOT NULL,
                password TEXT NOT NULL,
                client_id TEXT NOT NULL,
                refresh_token TEXT NOT NULL,
                is_registered INTEGER DEFAULT 0,
                group_name TEXT DEFAULT 'Основная',
                is_valid INTEGER DEFAULT 1
            )
        """)
        cursor = conn.execute("PRAGMA table_info(accounts)")
        columns = [row["name"] for row in cursor.fetchall()]
        if "group_name" not in columns:
            conn.execute("ALTER TABLE accounts ADD COLUMN group_name TEXT DEFAULT 'Основная'")
        if "is_valid" not in columns:
            conn.execute("ALTER TABLE accounts ADD COLUMN is_valid INTEGER DEFAULT 1")
        conn.commit()


@asynccontextmanager
async def lifespan(app: FastAPI):
    init_db()
    yield


app = FastAPI(title="Mail Dashboard", lifespan=lifespan, dependencies=[Depends(authenticate)])


class BulkImportRequest(BaseModel):
    raw_data: str
    group_name: Optional[str] = "Основная"


class BulkSendMailRequest(BaseModel):
    account_ids: List[int]
    to_email: str
    subject: str
    body: str


UUID_PATTERN = re.compile(r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$")
EMAIL_PATTERN = re.compile(r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$")


def parse_account_line(line: str):
    cleaned = line.strip().replace("\r", "")
    parts = [p.strip() for p in re.split(r"[:|;\t]+", cleaned) if p.strip()]
    if len(parts) < 4:
        return None

    email_val = None
    client_id = None
    refresh_token = None
    password = None

    remaining = list(parts)

    for p in remaining:
        if EMAIL_PATTERN.match(p):
            email_val = p
            remaining.remove(p)
            break

    for p in remaining:
        if UUID_PATTERN.match(p):
            client_id = p
            remaining.remove(p)
            break

    if remaining:
        longest = max(remaining, key=len)
        if len(longest) > 60:
            refresh_token = longest
            remaining.remove(longest)

    if remaining:
        password = remaining[0]

    if email_val and password and client_id and refresh_token:
        return {
            "email": email_val,
            "password": password,
            "client_id": client_id,
            "refresh_token": refresh_token,
        }
    return None


async def get_access_token(client_id: str, refresh_token: str) -> str:
    url = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"
    data = {
        "client_id": client_id,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "scope": "offline_access https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.Send",
    }

    async with httpx.AsyncClient(timeout=15.0) as client:
        resp = await client.post(url, data=data)
        if resp.status_code != 200:
            resp = await client.post("https://login.microsoftonline.com/common/oauth2/v2.0/token", data=data)
        if resp.status_code != 200:
            resp = await client.post("https://login.microsoftonline.com/common/oauth2/v2.0/token", data=data)

        if resp.status_code != 200:
            logger.warning("Token acquisition failed (status %d)", resp.status_code)
            raise HTTPException(status_code=400, detail="Токен недействителен")

        return resp.json().get("access_token")


async def execute_send(client_id: str, refresh_token: str, to_email: str, subject: str, body: str):
    token = await get_access_token(client_id, refresh_token)
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    email_payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body},
            "toRecipients": [{"emailAddress": {"address": to_email}}],
        },
        "saveToSentItems": "true",
    }

    async with httpx.AsyncClient(timeout=20.0) as client:
        resp = await client.post("https://graph.microsoft.com/v1.0/me/sendMail", headers=headers, json=email_payload)
        if resp.status_code == 401:
            resp = await client.post(
                "https://outlook.office.com/api/v2.0/me/sendmail",
                headers=headers,
                json={"Message": email_payload["message"], "SaveToSentItems": True},
            )
        if resp.status_code not in (200, 202):
            raise RuntimeError(f"Send mail failed with status {resp.status_code}")


@app.post("/api/accounts/bulk")
async def bulk_import(payload: BulkImportRequest):
    lines = [line.strip() for line in payload.raw_data.strip().splitlines() if line.strip()]
    if not lines:
        raise HTTPException(status_code=400, detail="Список пуст")

    target_group = payload.group_name.strip() if payload.group_name and payload.group_name.strip() else "Основная"
    added = 0
    failed = 0

    with get_db() as conn:
        cursor = conn.cursor()
        for line in lines:
            parsed = parse_account_line(line)
            if not parsed:
                failed += 1
                continue

            is_valid = 1
            try:
                await get_access_token(parsed["client_id"], parsed["refresh_token"])
            except Exception:
                is_valid = 0

            try:
                cursor.execute(
                    """INSERT INTO accounts (email, password, client_id, refresh_token, is_registered, group_name, is_valid)
                       VALUES (?, ?, ?, ?, 0, ?, ?)""",
                    (parsed["email"], parsed["password"], parsed["client_id"], parsed["refresh_token"], target_group, is_valid),
                )
                added += 1
            except sqlite3.Error:
                failed += 1

        conn.commit()

    return {"added": added, "failed": failed}


@app.get("/api/accounts")
def list_accounts():
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT id, email, password, is_registered, 
                   COALESCE(group_name, 'Основная') as group_name,
                   COALESCE(is_valid, 1) as is_valid 
            FROM accounts ORDER BY id DESC
        """)
        return [dict(row) for row in cursor.fetchall()]


@app.delete("/api/accounts/{account_id}")
def delete_account(account_id: int):
    with get_db() as conn:
        conn.execute("DELETE FROM accounts WHERE id = ?", (account_id,))
        conn.commit()
    return {"status": "ok"}


@app.delete("/api/groups/{group_name}")
def delete_group(group_name: str):
    with get_db() as conn:
        conn.execute("DELETE FROM accounts WHERE group_name = ?", (group_name,))
        conn.commit()
    return {"status": "ok"}


@app.post("/api/accounts/{account_id}/mark-registered")
def mark_registered(account_id: int):
    with get_db() as conn:
        conn.execute("UPDATE accounts SET is_registered = 1 WHERE id = ?", (account_id,))
        conn.commit()
    return {"status": "ok"}


@app.get("/api/accounts/{account_id}/messages")
async def get_messages(account_id: int, folder: str = "inbox"):
    target_folder = "junkemail" if folder == "junk" else "inbox"

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT client_id, refresh_token FROM accounts WHERE id = ?", (account_id,))
        acc = cursor.fetchone()
        if not acc:
            raise HTTPException(status_code=404, detail="Аккаунт не найден")

    try:
        token = await get_access_token(acc["client_id"], acc["refresh_token"])
    except Exception:
        with get_db() as conn:
            conn.execute("UPDATE accounts SET is_valid = 0 WHERE id = ?", (account_id,))
            conn.commit()
        raise HTTPException(status_code=400, detail="Токен недействителен")

    headers = {"Authorization": f"Bearer {token}"}

    async with httpx.AsyncClient(timeout=15.0) as client:
        resp = await client.get(
            f"https://graph.microsoft.com/v1.0/me/mailFolders/{target_folder}/messages?$top=10&$select=sender,subject,receivedDateTime,bodyPreview",
            headers=headers,
        )

        if resp.status_code == 401:
            resp = await client.get(
                f"https://outlook.office.com/api/v2.0/me/mailfolders/{target_folder}/messages?$top=10&$select=Sender,Subject,ReceivedDateTime,BodyPreview",
                headers=headers,
            )
            if resp.status_code == 200:
                data = resp.json().get("value", [])
                return [
                    {
                        "sender": {
                            "emailAddress": {
                                "name": m.get("Sender", {}).get("EmailAddress", {}).get("Name", ""),
                                "address": m.get("Sender", {}).get("EmailAddress", {}).get("Address", ""),
                            }
                        },
                        "subject": m.get("Subject", "Без темы"),
                        "bodyPreview": m.get("BodyPreview", ""),
                    }
                    for m in data
                ]

        if resp.status_code != 200:
            raise HTTPException(status_code=resp.status_code, detail="Ошибка загрузки писем")

        return resp.json().get("value", [])


@app.post("/api/accounts/bulk-send")
async def bulk_send_mail(payload: BulkSendMailRequest):
    if not payload.account_ids:
        raise HTTPException(status_code=400, detail="Не выбрана ни одна почта")
    if not payload.to_email.strip():
        raise HTTPException(status_code=400, detail="Укажите адрес получателя")

    placeholders = ",".join("?" for _ in payload.account_ids)
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute(f"SELECT id, email, client_id, refresh_token FROM accounts WHERE id IN ({placeholders})", payload.account_ids)
        accounts = cursor.fetchall()

    sent = 0
    failed = 0

    for acc in accounts:
        try:
            await execute_send(acc["client_id"], acc["refresh_token"], payload.to_email, payload.subject, payload.body)
            sent += 1
        except Exception:
            logger.error("Failed to send mail from %s", acc["email"])
            failed += 1

    return {"sent": sent, "failed": failed}


@app.api_route("/", methods=["GET", "HEAD"], response_class=HTMLResponse)
def serve_index():
    with open("index.html", "r", encoding="utf-8") as f:
        return f.read()


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("main:app", host="0.0.0.0", port=port, ws="none")