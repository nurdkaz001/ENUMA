from __future__ import annotations

import difflib
import io
import json
import logging
import os
import random
import re
import secrets
import time
from collections import defaultdict, deque
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Optional

import httpx
import jwt
from fastapi import Body, Depends, FastAPI, File, Form, Header, HTTPException, Request, UploadFile
from fastapi.exceptions import RequestValidationError
from fastapi.responses import HTMLResponse, JSONResponse, Response
from pydantic import BaseModel, Field
import openpyxl
from docx import Document
from pypdf import PdfReader
from starlette.exceptions import HTTPException as StarletteHTTPException
from supabase import Client, create_client

# ==================== НАСТРОЙКИ ====================
log = logging.getLogger("studyai")
logging.basicConfig(level=logging.INFO)

DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "")
DEEPSEEK_MODEL = os.getenv("DEEPSEEK_MODEL", "deepseek-chat")
DEEPSEEK_BASE_URL = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com").rstrip("/")
SUPABASE_URL = os.getenv("SUPABASE_URL", "")
SUPABASE_SERVICE_KEY = os.getenv("SUPABASE_SERVICE_KEY", "")
JWT_SECRET = os.getenv("JWT_SECRET", "")
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_WEBHOOK_SECRET = os.getenv("TELEGRAM_WEBHOOK_SECRET", "")
DAILY_TOKENS = int(os.getenv("DAILY_TOKENS", "3"))
PRO_TOKENS = int(os.getenv("PRO_TOKENS", "50"))          # токенов в день по подписке
ADMIN_IDS = {int(x) for x in re.split(r"[,\s]+", os.getenv("ADMIN_TELEGRAM_IDS", "")) if x.isdigit()}
MAX_DOC_CHARS = int(os.getenv("MAX_DOC_CHARS", "30000"))   # сколько текста отдаём модели
MAX_STORE_CHARS = 200_000                                  # сколько текста PDF храним
MAX_UPLOAD_BYTES = 15 * 1024 * 1024
JWT_DAYS = 30
CODE_TTL_MIN = 10
DAILY_XP_CAP = 1500

_missing = [k for k, v in {
    "DEEPSEEK_API_KEY": DEEPSEEK_API_KEY, "SUPABASE_URL": SUPABASE_URL,
    "SUPABASE_SERVICE_KEY": SUPABASE_SERVICE_KEY, "JWT_SECRET": JWT_SECRET,
}.items() if not v]
if _missing:
    raise RuntimeError("Не заданы переменные окружения: " + ", ".join(_missing))
# ====================================================

app = FastAPI(title="StudyAI Hub")
sb: Client = create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)
INDEX_FILE = Path(__file__).parent / "templates" / "index.html"

HISTORY_TYPES = {"quiz", "flashcards", "summary", "simple", "plan", "teacher", "tickets"}
LANG_LINE = {
    "kz": "Жауапты қазақ тілінде бер.",
    "ru": "Отвечай на русском языке.",
    "en": "Answer in English.",
}


# ==================== ОШИБКИ В ФОРМАТЕ {"error": ...} ====================
@app.exception_handler(StarletteHTTPException)
async def http_error(_: Request, exc: StarletteHTTPException):
    return JSONResponse({"error": str(exc.detail)}, status_code=exc.status_code)


@app.exception_handler(RequestValidationError)
async def validation_error(_: Request, exc: RequestValidationError):
    return JSONResponse({"error": "Некорректный запрос."}, status_code=422)


@app.exception_handler(Exception)
async def unhandled_error(_: Request, exc: Exception):
    log.exception("Unhandled error")
    return JSONResponse({"error": "Внутренняя ошибка сервера."}, status_code=500)


# ==================== АВТОРИЗАЦИЯ ====================
def make_token(user_id: str) -> str:
    exp = datetime.now(timezone.utc) + timedelta(days=JWT_DAYS)
    return jwt.encode({"sub": user_id, "exp": exp}, JWT_SECRET, algorithm="HS256")


def today_utc() -> date:
    return datetime.now(timezone.utc).date()


def current_user(authorization: Optional[str] = Header(default=None)) -> dict:
    if not authorization or not authorization.lower().startswith("bearer "):
        raise HTTPException(401, "Нужно войти в аккаунт.")
    try:
        payload = jwt.decode(authorization[7:].strip(), JWT_SECRET, algorithms=["HS256"])
    except jwt.PyJWTError:
        raise HTTPException(401, "Сессия истекла. Войдите снова.")
    res = sb.table("profiles").select("*").eq("id", payload["sub"]).limit(1).execute()
    if not res.data:
        raise HTTPException(401, "Пользователь не найден.")
    user = res.data[0]
    # подписка закончилась: снимаем Premium и аннулируем токены
    until = user.get("subscription_until")
    if user.get("is_subscribed") and until and _dt(until) < datetime.now(timezone.utc):
        upd = {"is_subscribed": False, "subscription_until": None, "tokens": 0,
               "tokens_refilled_on": today_utc().isoformat()}
        sb.table("profiles").update(upd).eq("id", user["id"]).execute()
        user.update(upd)
    # ежедневное пополнение токенов
    if str(user.get("tokens_refilled_on")) < today_utc().isoformat():
        refilled = sb.rpc("refill_tokens", {"p_user": user["id"], "p_amount": PRO_TOKENS if user.get("is_subscribed") else DAILY_TOKENS}).execute().data
        if refilled is not None:
            user["tokens"] = refilled
    return user


def public_profile(u: dict) -> dict:
    return {
        "user_id": u["id"], "telegram_id": u.get("telegram_id"),
        "username": u.get("username"), "display_name": u.get("display_name") or u.get("username"),
        "tokens": u.get("tokens", 0), "score": u.get("score", 0),
        "is_subscribed": bool(u.get("is_subscribed")), "is_admin": u.get("telegram_id") in ADMIN_IDS, "subscription_until": u.get("subscription_until"), "created_at": u.get("created_at"),
    }


# простой ограничитель попыток входа по IP
_attempts: dict[str, deque] = defaultdict(deque)


def rate_limit(key: str, limit: int, window_sec: int) -> None:
    now = time.time()
    q = _attempts[key]
    while q and now - q[0] > window_sec:
        q.popleft()
    if len(q) >= limit:
        raise HTTPException(429, "Слишком много попыток. Подождите несколько минут.")
    q.append(now)


@app.post("/api/auth/verify-code")
def verify_code(request: Request, code: str = Form(...)):
    ip = request.client.host if request.client else "?"
    rate_limit(f"verify:{ip}", 10, 600)
    code = code.strip()
    now = datetime.now(timezone.utc).isoformat()
    found = (sb.table("auth_codes").select("*").eq("code", code).eq("used", False)
             .gt("expires_at", now).limit(1).execute().data)
    if not found:
        raise HTTPException(400, "Неверный или просроченный код.")
    row = found[0]
    claimed = sb.table("auth_codes").update({"used": True}).eq("code", code).eq("used", False).execute().data
    if not claimed:
        raise HTTPException(400, "Код уже использован.")

    prof = sb.table("profiles").upsert(
        {"telegram_id": row["telegram_id"], "username": row.get("username"),
         "display_name": row.get("display_name")},
        on_conflict="telegram_id",
    ).execute().data[0]
    return {**public_profile(prof), "token": make_token(prof["id"])}


@app.post("/api/auth/start")
def auth_start():
    return {"nonce": secrets.token_urlsafe(12)}


@app.get("/api/auth/poll")
def auth_poll(request: Request, nonce: str):
    ip = request.client.host if request.client else "?"
    rate_limit(f"poll:{ip}", 400, 600)
    if not re.fullmatch(r"[A-Za-z0-9_-]{8,40}", nonce):
        raise HTTPException(400, "bad nonce")
    now = datetime.now(timezone.utc).isoformat()
    found = (sb.table("auth_codes").select("*").eq("nonce", nonce).eq("used", False)
             .gt("expires_at", now).limit(1).execute().data)
    if not found:
        return {"ready": False}
    row = found[0]
    if not sb.table("auth_codes").update({"used": True}).eq("code", row["code"]).eq("used", False).execute().data:
        return {"ready": False}
    prof = sb.table("profiles").upsert(
        {"telegram_id": row["telegram_id"], "username": row.get("username"), "display_name": row.get("display_name")},
        on_conflict="telegram_id").execute().data[0]
    return {"ready": True, **public_profile(prof), "token": make_token(prof["id"])}


def tg_send(chat_id: int, text: str) -> bool:
    if not TELEGRAM_BOT_TOKEN:
        return False
    try:
        r = httpx.post(f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage",
                       json={"chat_id": chat_id, "text": text}, timeout=10)
        return r.status_code == 200
    except httpx.HTTPError:
        log.exception("telegram send failed")
        return False


@app.post("/api/telegram/webhook")
def telegram_webhook(update: dict = Body(...),
                     x_telegram_bot_api_secret_token: Optional[str] = Header(default=None)):
    """Бот отвечает на любое сообщение новым одноразовым кодом входа.
    Включить: https://api.telegram.org/bot<TOKEN>/setWebhook?url=<SITE>/api/telegram/webhook&secret_token=<SECRET>"""
    if TELEGRAM_WEBHOOK_SECRET and x_telegram_bot_api_secret_token != TELEGRAM_WEBHOOK_SECRET:
        raise HTTPException(403, "forbidden")
    msg = update.get("message") or {}
    parts = (msg.get("text") or "").split()
    nonce = parts[1] if len(parts) > 1 and parts[0].startswith("/start") and re.fullmatch(r"[A-Za-z0-9_-]{8,40}", parts[1]) else None
    chat_id = (msg.get("chat") or {}).get("id")
    frm = msg.get("from") or {}
    if not chat_id:
        return {"ok": True}
    sb.table("auth_codes").delete().eq("telegram_id", chat_id).eq("used", False).execute()
    name = " ".join(filter(None, [frm.get("first_name"), frm.get("last_name")])) or frm.get("username")
    for _ in range(5):
        code = f"{secrets.randbelow(1_000_000):06d}"
        try:
            sb.table("auth_codes").insert({
                "code": code, "telegram_id": chat_id, "username": frm.get("username"),
                "display_name": name, "nonce": nonce,
                "expires_at": (datetime.now(timezone.utc) + timedelta(minutes=CODE_TTL_MIN)).isoformat(),
            }).execute()
            break
        except Exception:
            code = ""
    if not code:
        tg_send(chat_id, "Не удалось создать код, попробуйте ещё раз.")
        return {"ok": True}
    tg_send(chat_id, (("✅ Готово! Вернитесь на сайт — вход выполнится автоматически.\n" if nonce else "")
                      + f"Ваш код для входа: {code}\nДействует {CODE_TTL_MIN} минут."))
    return {"ok": True}


@app.post("/api/telegram/reminder")
def telegram_reminder(task_title: str = Form(...), task_date: str = Form(""), user: dict = Depends(current_user)):
    if not user.get("telegram_id"):
        raise HTTPException(400, "Telegram не привязан.")
    ok = tg_send(user["telegram_id"], f"🔔 Напоминание: {task_title[:200]}\n📅 {task_date[:10]}")
    if not ok:
        raise HTTPException(502, "Не удалось отправить сообщение в Telegram.")
    return {"ok": True}


# ==================== ПОМОЩНИКИ ====================
def ask_ai(prompt: str, json_mode: bool = False, max_tokens: int = 4000, temperature: float = 0.7, timeout: float = 90.0) -> str:
    """DeepSeek (OpenAI-совместимый API). В json_mode модель обязана вернуть JSON-ОБЪЕКТ."""
    payload = {
        "model": DEEPSEEK_MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": max_tokens,
        "temperature": temperature,
    }
    if json_mode:
        payload["response_format"] = {"type": "json_object"}
    try:
        r = httpx.post(
            f"{DEEPSEEK_BASE_URL}/chat/completions",
            headers={"Authorization": f"Bearer {DEEPSEEK_API_KEY}"},
            json=payload, timeout=httpx.Timeout(timeout, connect=10.0),
        )
        if r.status_code == 402:
            log.error("DeepSeek: недостаточно средств на балансе")
            raise HTTPException(502, "ИИ временно недоступен (проверьте баланс DeepSeek).")
        r.raise_for_status()
        text = (r.json()["choices"][0]["message"]["content"] or "").strip()
    except HTTPException:
        raise
    except Exception:
        log.exception("deepseek failed")
        raise HTTPException(502, "ИИ сейчас недоступен, попробуйте ещё раз.")
    if not text:
        raise HTTPException(502, "ИИ вернул пустой ответ.")
    return text


def parse_json(text: str):
    s = text.replace("```json", "").replace("```", "").strip()
    try:
        return json.loads(s)
    except json.JSONDecodeError:
        pass
    starts = [i for i in (s.find("{"), s.find("[")) if i >= 0]
    if not starts:
        return None
    a = min(starts)
    b = s.rfind("}" if s[a] == "{" else "]")
    try:
        return json.loads(s[a:b + 1])
    except json.JSONDecodeError:
        return None


def spend_token(user_id: str) -> int:
    left = sb.rpc("spend_token", {"p_user": user_id}).execute().data
    if left is None or left < 0:
        raise HTTPException(402, "Токены закончились. Они обновятся завтра.")
    return left


def refund_token(user_id: str) -> None:
    try:
        sb.rpc("refund_token", {"p_user": user_id}).execute()
    except Exception:
        log.exception("refund failed")


def paid(user: dict, fn):
    """Списывает токен, выполняет fn, при ошибке возвращает токен."""
    left = spend_token(user["id"])
    try:
        return fn(), left
    except Exception:
        refund_token(user["id"])
        raise


def save_generation(user_id: str, type_: str, title: str, content: str, doc_id: Optional[str] = None) -> Optional[int]:
    try:
        row = sb.table("generations").insert({
            "user_id": user_id, "type": type_, "title": (title or "")[:120],
            "content": content, "doc_id": doc_id,
        }).execute().data
        return row[0]["id"] if row else None
    except Exception:
        log.exception("save_generation failed")
        return None


def get_doc_text(doc_id: str, user: dict) -> Optional[str]:
    """'' — документа нет (это допустимо), None — документ указан, но не найден."""
    if not doc_id or doc_id in ("default", "planner"):
        return ""
    if doc_id.startswith("topic_"):
        return f"Тема для изучения: {doc_id[6:206]}"
    res = sb.table("documents").select("text").eq("id", doc_id).eq("user_id", user["id"]).limit(1).execute().data
    return res[0]["text"] if res else None


def build_context(doc_id: str, title: str, user: dict) -> str:
    text = get_doc_text(doc_id, user)
    if text is None:
        raise HTTPException(400, "Документ не найден. Загрузите файл заново.")
    if not text and title and title.strip() and title.strip() != "Lesson":
        text = f"Тема для изучения: {title.strip()[:200]}"
    if not text:
        raise HTTPException(400, "Сначала загрузите PDF или введите тему.")
    return text[:MAX_DOC_CHARS]


def lang_line(lang: str) -> str:
    return LANG_LINE.get(lang, LANG_LINE["ru"])


# ==================== СТРАНИЦА / ЗАГРУЗКА / ГЕНЕРАЦИЯ ====================
@app.get("/", response_class=HTMLResponse)
def read_root():
    return HTMLResponse(INDEX_FILE.read_text(encoding="utf-8"))


@app.get("/health")
def health():
    return {"ok": True}


def read_upload(file: UploadFile) -> str:
    """Достаёт текст из PDF, Word (.docx), Excel (.xlsx) или .txt."""
    name = (file.filename or "").lower()
    data = file.file.read(MAX_UPLOAD_BYTES + 1)
    if len(data) > MAX_UPLOAD_BYTES:
        raise HTTPException(413, "Файл больше 15 МБ.")
    parts: list[str] = []
    try:
        if name.endswith(".pdf"):
            reader = PdfReader(io.BytesIO(data))
            if reader.is_encrypted:
                raise HTTPException(400, "PDF защищён паролем.")
            parts = list(filter(None, (p.extract_text() for p in reader.pages)))
        elif name.endswith(".docx"):
            d = Document(io.BytesIO(data))
            parts = [p.text for p in d.paragraphs if p.text.strip()]
            for t in d.tables:
                for row in t.rows:
                    parts.append(" | ".join(c.text.strip() for c in row.cells))
        elif name.endswith((".xlsx", ".xlsm")):
            wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
            size = 0
            for ws in wb.worksheets:
                parts.append(f"# Лист: {ws.title}")
                for row in ws.iter_rows(values_only=True):
                    cells = [str(c).strip() for c in row if c is not None and str(c).strip()]
                    if cells:
                        line = " | ".join(cells)
                        parts.append(line)
                        size += len(line)
                if size > MAX_STORE_CHARS:
                    break
        elif name.endswith(".txt"):
            parts = [data.decode("utf-8", "ignore")]
        else:
            raise HTTPException(400, "Поддерживаются PDF, Word (.docx), Excel (.xlsx) и .txt.")
    except HTTPException:
        raise
    except Exception:
        log.exception("read_upload failed")
        raise HTTPException(400, "Не удалось прочитать файл.")
    text = "\n".join(parts)
    if len(text.strip()) < 30:
        raise HTTPException(400, "В файле не найден текст (возможно, это скан).")
    return text


@app.post("/upload")
def upload_file(file: UploadFile = File(...), user: dict = Depends(current_user)):
    text = read_upload(file)
    doc_id = secrets.token_hex(8)
    sb.table("documents").insert({
        "id": doc_id, "user_id": user["id"], "filename": (file.filename or "")[:200],
        "text": text[:MAX_STORE_CHARS],
    }).execute()
    return {"doc_id": doc_id, "message": "Файл успешно прочитан!", "chars": len(text)}


PROMPTS = {
    "summary": "Сделай структурированный краткий конспект на основе этого текста. Выдели главные тезисы, формулы и термины. Используй заголовки ### и **жирный** для терминов.",
    "quiz": "Составь 10 качественных тестовых вопросов с вариантами (A, B, C, D) и укажи правильные ответы в самом конце списком.",
    "flashcards": "Создай 10 карточек (Вопрос — Ответ) для заучивания ключевых понятий из текста.",
    "simple": "Объясни сложные идеи из этого текста максимально простыми словами, используя жизненные примеры.",
    "plan": "Составь пошаговый план подготовки и изучения этого материала.",
}


@app.post("/generate")
def generate_content(
    action: str = Form(...), doc_id: str = Form("default"), lang: str = Form("kz"),
    title: str = Form(""), custom_prompt: str = Form(""), user: dict = Depends(current_user),
):
    # ИИ-планер: только для подписчиков, промпт собирает фронтенд, ответ — JSON {"tasks": [...]}
    if action == "planner":
        if not user.get("is_subscribed"):
            raise HTTPException(403, "ИИ-планер доступен по подписке Premium.")
        prompt = custom_prompt.strip()[:1500]
        if not prompt:
            raise HTTPException(400, "Введите запрос для планера.")
        return {"result": ask_ai(prompt, json_mode=True), "tokens_left": user["tokens"]}

    if action not in PROMPTS:
        raise HTTPException(400, "Неизвестный режим.")
    context = build_context(doc_id, title, user)
    prompt = f"{PROMPTS[action]}\n{lang_line(lang)}\n\nМатериал:\n{context}"
    result, left = paid(user, lambda: ask_ai(prompt))
    hid = save_generation(user["id"], action, title if title != "Lesson" else "", result, doc_id)
    return {"result": result, "tokens_left": left, "history_id": hid}


@app.post("/api/chat")
def chat(message: str = Form(...), doc_id: str = Form("default"), lang: str = Form("kz"),
         user: dict = Depends(current_user)):
    message = message.strip()[:2000]
    if not message:
        raise HTTPException(400, "Пустое сообщение.")
    ctx = get_doc_text(doc_id, user) or ""
    past = (sb.table("chat_messages").select("message,reply").eq("user_id", user["id"])
            .order("created_at", desc=True).limit(6).execute().data or [])
    dialog = "\n".join(f"Студент: {m['message']}\nНаставник: {m['reply']}" for m in reversed(past))
    prompt = (
        "Ты — дружелюбный ИИ-наставник для студентов. Отвечай понятно и по делу. "
        "Формулы пиши в LaTeX между $...$. " + lang_line(lang) + "\n\n"
        + (f"Материал студента:\n{ctx[:MAX_DOC_CHARS]}\n\n" if ctx else "")
        + (f"Предыдущий диалог:\n{dialog}\n\n" if dialog else "")
        + f"Студент: {message}\nНаставник:"
    )
    reply, left = paid(user, lambda: ask_ai(prompt))
    try:
        sb.table("chat_messages").insert({"user_id": user["id"], "doc_id": doc_id,
                                          "message": message, "reply": reply}).execute()
    except Exception:
        log.exception("chat save failed")
    return {"reply": reply, "tokens_left": left}


@app.post("/api/ai-teacher")
def ai_teacher(topic: str = Form(...), comment: str = Form(""), doc_id: str = Form("default"),
               lang: str = Form("kz"), user: dict = Depends(current_user)):
    topic = topic.strip()[:200]
    ctx = (get_doc_text(doc_id, user) or "")[:MAX_DOC_CHARS]
    prompt = (
        f"Ты — опытный учитель. Объясни тему «{topic}» простым и понятным языком (6–10 предложений, с примером). "
        f"Пожелания студента: {comment[:500] or 'нет'}. " + lang_line(lang) + "\n"
        + (f"Опирайся на материал студента:\n{ctx}\n" if ctx else "")
        + 'Верни ТОЛЬКО JSON такого вида: {"explanation":"...","video_query":"поисковый запрос для YouTube",'
          '"game_questions":[{"q":"вопрос","options":["A) ...","B) ...","C) ...","D) ..."],"answer":"A"}]}. '
          "Ровно 5 вопросов для закрепления, answer — буква правильного варианта."
    )

    def run():
        data = parse_json(ask_ai(prompt, json_mode=True))
        if not isinstance(data, dict) or not data.get("explanation"):
            raise HTTPException(502, "ИИ вернул неверный формат. Попробуйте ещё раз.")
        return data

    data, left = paid(user, run)
    save_generation(user["id"], "teacher", topic, json.dumps(data, ensure_ascii=False), doc_id)
    return {"result": data, "tokens_left": left}


@app.post("/api/tickets")
def tickets(theme: str = Form(...), comment: str = Form(""), count: int = Form(3),
            doc_id: str = Form("default"), lang: str = Form("kz"), user: dict = Depends(current_user)):
    theme = theme.strip()[:200]
    count = max(1, min(count, 10))
    ctx = (get_doc_text(doc_id, user) or "")[:MAX_DOC_CHARS]
    prompt = (
        f"Составь {count} экзаменационных билетов по теме «{theme}» в формате игры-викторины. "
        f"Дополнительные требования: {comment[:500] or 'нет'}. " + lang_line(lang) + "\n"
        + (f"Опирайся на материал студента:\n{ctx}\n" if ctx else "")
        + 'Верни ТОЛЬКО JSON-объект: {"tickets":[{"ticket_num":1,"question":"...","game_task":"короткое игровое задание",'
          '"options":["A) ...","B) ...","C) ..."],"correct":"A"}]}. correct — буква правильного варианта.'
    )

    def run():
        data = parse_json(ask_ai(prompt, json_mode=True))
        if isinstance(data, dict):
            data = data.get("tickets") or data.get("questions")
        if not isinstance(data, list) or not data:
            raise HTTPException(502, "ИИ вернул неверный формат. Попробуйте ещё раз.")
        return data

    data, left = paid(user, run)
    save_generation(user["id"], "tickets", theme, json.dumps(data, ensure_ascii=False), doc_id)
    return {"result": data, "tokens_left": left}


QUIZ_FMT = (
    'Верни ТОЛЬКО JSON-объект: {"questions":[{"q":"вопрос","options":["A) ...","B) ...","C) ...","D) ..."],'
    '"correct":"A","explain":"одно короткое предложение: почему это верно"}]}. '
    "Ровно 4 варианта, один верный; верные ответы равномерно распредели по A, B, C, D. "
    "Вопросы точные и однозначные, без вариантов «всё вышеперечисленное»."
)


def make_questions(prompt: str, n: int):
    def run():
        data = parse_json(ask_ai(prompt, json_mode=True, max_tokens=6000))
        if isinstance(data, dict):
            data = data.get("questions")
        if not isinstance(data, list) or len(data) < min(n, 4):
            raise HTTPException(502, "ИИ вернул неверный формат. Попробуйте ещё раз.")
        return data[:n]
    return run


@app.post("/api/quiz")
def quiz(doc_id: str = Form("default"), lang: str = Form("kz"), title: str = Form(""),
         user: dict = Depends(current_user)):
    context = build_context(doc_id, title, user)
    prompt = f"Составь 20 тестовых вопросов по материалу. {QUIZ_FMT}\n{lang_line(lang)}\n\nМатериал:\n{context}"
    data, left = paid(user, make_questions(prompt, 20))
    save_generation(user["id"], "quiz", title if title != "Lesson" else "",
                    json.dumps({"questions": data}, ensure_ascii=False), doc_id)
    return {"questions": data, "tokens_left": left}


class WrongIn(BaseModel):
    q: str = Field(max_length=500)
    correct: str = Field(max_length=300)
    chosen: str = Field(default="", max_length=300)


class ReviewIn(BaseModel):
    wrong: list[WrongIn] = Field(min_length=1, max_length=20)
    lang: str = "ru"


@app.post("/api/quiz/review")
def quiz_review(body: ReviewIn, user: dict = Depends(current_user)):
    listing = "\n".join(f"{i + 1}. Вопрос: {w.q}\n   Ответ студента: {w.chosen}\n   Верный ответ: {w.correct}"
                        for i, w in enumerate(body.wrong))
    prompt = (
        "Ты добрый учитель. Студент ошибся в этих вопросах:\n" + listing + "\n"
        "Для каждой ошибки объясни ПРОСТЫМИ словами, как другу, с бытовым примером: почему верный ответ именно такой "
        "и в чём была ловушка; дай короткий способ запомнить. Затем сделай «закрепление темы»: 4–6 коротких пунктов "
        "главного и 3 вопроса для самопроверки. " + lang_line(body.lang) +
        ' Верни ТОЛЬКО JSON: {"items":[{"explain":"...","tip":"..."}],"recap":["..."],"selfcheck":["..."]}. '
        "items — в том же порядке и количестве, что и ошибки."
    )

    def run():
        data = parse_json(ask_ai(prompt, json_mode=True, max_tokens=5000))
        if not isinstance(data, dict) or not data.get("items"):
            raise HTTPException(502, "ИИ вернул неверный формат. Попробуйте ещё раз.")
        return data

    data, left = paid(user, run)
    return {"result": data, "tokens_left": left}


@app.post("/api/tickets/open")
def ticket_open(theme: str = Form(""), comment: str = Form(""), num: int = Form(1), doc_id: str = Form("default"),
                lang: str = Form("kz"), questions: str = Form(""), user: dict = Depends(current_user)):
    """Открывает один билет: 20 точных вопросов с вариантами. Содержимое генерируется только при открытии."""
    num = max(1, min(num, 4))
    theme = theme.strip()[:200]
    ctx = (get_doc_text(doc_id, user) or "")[:MAX_DOC_CHARS]
    given = parse_json(questions) if questions else None
    if isinstance(given, list) and given:
        given = [str(q)[:300] for q in given[:20]]
        n = len(given)
        task = ("Для КАЖДОГО из этих экзаменационных вопросов составь тестовый вопрос (сохрани смысл) "
                "с 4 вариантами и верным ответом:\n" + "\n".join(f"{i + 1}. {q}" for i, q in enumerate(given)))
    else:
        if not theme and not ctx:
            raise HTTPException(400, "Введите тему или загрузите материал.")
        n = 20
        task = (f"Составь экзаменационный билет №{num} из 4 по теме «{theme or 'материал'}»: ровно 20 точных вопросов. "
                f"Четыре билета покрывают РАЗНЫЕ части темы, в этом билете сделай упор на часть №{num} из 4. "
                f"Пожелания: {comment[:500] or 'нет'}. Вариация: {secrets.token_hex(3)}.")
    prompt = f"{task}\n{QUIZ_FMT}\n{lang_line(lang)}\n" + (f"Опирайся на материал:\n{ctx}\n" if ctx else "")
    data, left = paid(user, make_questions(prompt, n))
    save_generation(user["id"], "tickets", f"{theme or 'Билет'} · №{num}",
                    json.dumps({"questions": data}, ensure_ascii=False), doc_id)
    return {"questions": data, "tokens_left": left}


@app.post("/api/tickets/custom")
def tickets_custom(file: UploadFile = File(...), user: dict = Depends(current_user)):
    """Свои вопросы из PDF/Word/Excel: ИИ вытаскивает вопросы, сервер случайно раскладывает их по 4 билетам."""
    text = read_upload(file)
    prompt = ("Из текста выдели все экзаменационные вопросы (не больше 80), без номеров, по одному на строку, "
              'смысл не меняй. Верни ТОЛЬКО JSON: {"questions":["..."]}.\n\n' + text[:MAX_DOC_CHARS])

    def run():
        data = parse_json(ask_ai(prompt, json_mode=True, max_tokens=4000))
        qs = data.get("questions") if isinstance(data, dict) else data
        qs = [str(q).strip() for q in (qs or []) if str(q).strip()][:80]
        if len(qs) < 4:
            raise HTTPException(400, "Не нашёл вопросов в файле (нужно минимум 4).")
        return qs

    qs, left = paid(user, run)
    random.shuffle(qs)
    per = min(20, -(-len(qs) // 4))
    groups = [g for g in (qs[i * per:(i + 1) * per] for i in range(4)) if g]
    return {"groups": groups, "tokens_left": left}


@app.post("/api/flashcards")
def flashcards(doc_id: str = Form("default"), lang: str = Form("kz"), title: str = Form(""),
               user: dict = Depends(current_user)):
    context = build_context(doc_id, title, user)
    prompt = ("Создай 12 карточек для заучивания ключевых понятий из материала. Вопрос короткий (до 15 слов), "
              "ответ — 1–3 предложения. " + lang_line(lang) +
              ' Верни ТОЛЬКО JSON: {"cards":[{"q":"...","a":"..."}]}.\n\nМатериал:\n' + context)

    def run():
        d = parse_json(ask_ai(prompt, json_mode=True, max_tokens=4000))
        cards = d.get("cards") if isinstance(d, dict) else d
        if not isinstance(cards, list) or len(cards) < 3:
            raise HTTPException(502, "ИИ вернул неверный формат. Попробуйте ещё раз.")
        return cards[:12]

    cards, left = paid(user, run)
    save_generation(user["id"], "flashcards", title if title != "Lesson" else "",
                    json.dumps({"cards": cards}, ensure_ascii=False), doc_id)
    return {"cards": cards, "tokens_left": left}


PARASITE_RE = re.compile(r"\b(?:короче|типа|как бы|ну|прикинь|кстати|собственно говоря|по сути|на самом деле|в общем-то)\b", re.I)


def _words(t: str) -> list[str]:
    return re.findall(r"\w+", t.lower())


@app.post("/api/humanize")
def humanize(text: str = Form(...), lang: str = Form("ru"), user: dict = Depends(current_user)):
    """Оценка «похоже на ИИ» + глубокое переписывание живым языком + список изменений. 1 токен."""
    text = text.strip()[:6000]
    if len(text) < 80:
        raise HTTPException(400, "Вставьте текст (минимум 80 символов).")
    base = (
        "Ты — опытный редактор и копирайтер, пишущий в живом, разговорном и публицистическом стиле. "
        "Выполни две задачи с текстом ниже.\n"
        "ЗАДАЧА 1 — анализ. Оцени по шкале 0–100, насколько текст похож на написанный нейросетью (ai_score), "
        "в одном предложении объясни вердикт и найди до 6 признаков: фраза из текста (до 8 слов) и почему она выдаёт ИИ.\n"
        "ЗАДАЧА 2 — сначала определи СТИЛЬ исходного текста (style): научный, официально-деловой, публицистический, "
        "учебный, художественный или разговорный. Затем ГЛУБОКО перепиши текст (human_text), строго оставаясь В ЭТОМ СТИЛЕ "
        "и в рамках норм речевого этикета: вежливо, грамотно, уважительно. Текст должен читаться так, будто его написал другой "
        "человек, но смысл, факты, числа и имена сохранены, ничего не выдумывай.\n"
        "ОБЯЗАТЕЛЬНО:\n"
        "• Меняй КАЖДОЕ предложение. Дословно оставлять предложение из оригинала ЗАПРЕЩЕНО: меняй порядок слов, начало фразы, "
        "объединяй короткие предложения и дроби длинные.\n"
        "• Заменяй слова на уместные для стиля синонимы: минимум каждое второе-третье значимое слово "
        "(существительные, глаголы, прилагательные, наречия).\n"
        "• Чередуй длину предложений волнообразно: одно длинное (с причастным оборотом или вводной конструкцией), следом короткое, "
        "если стиль это допускает. Связанные мысли можно разделять точкой.\n"
        "• Полностью исключи маркеры ИИ: «важно отметить», «таким образом», «в заключение», «несомненно», «ключевой аспект», "
        "«стоит упомянуть», «погрузиться в», «экосистема», «инновационный», «трансформация», «в современном мире», «играет роль».\n"
        "• Канцеляризмы меняй на простые глаголы («осуществляет процесс анализа» → «анализирует»), предпочитай активный залог.\n"
        "СТРОГО ЗАПРЕЩЕНО: слова-паразиты и вставки («короче», «типа», «как бы», «ну», «прикинь», «кстати», «собственно говоря», "
        "«по сути», «на самом деле»), сленг, жаргон, просторечие, мат, панибратство — даже если исходный текст разговорный. "
        "Живость достигай ритмом фраз и точными словами, а не паразитами.\n"
        "Пример (публицистический стиль). Было: «Идея про баланс звучит отлично, но на практике этот баланс каждый определяет сам.» "
        "Стало: «Идея баланса привлекательна. Однако границы его каждый выстраивает самостоятельно.»\n"
        "Язык результата — тот же, что у исходного текста. human_text содержит ТОЛЬКО очеловеченный текст: без комментариев, "
        "вступлений и кавычек вокруг. Оцени новый текст по той же шкале (human_score). "
        "changes — 4–6 пар «было → стало», где after ЗАМЕТНО отличается от before (не просто пунктуация). "
        "Поля verdict и why пиши на языке интерфейса: " + lang_line(lang) + "\n"
        'Верни ТОЛЬКО JSON: {"style":"...","ai_score":0,"verdict":"...","markers":[{"phrase":"...","why":"..."}],'
        '"human_text":"...","human_score":0,"changes":[{"before":"...","after":"..."}]}\n\nТЕКСТ:\n' + text
    )
    orig_words = _words(text)

    def run():
        best, best_ratio, best_pct, extra = None, 9.0, 1.0, ""
        for attempt in range(2):
            d = parse_json(ask_ai(base + extra, json_mode=True, max_tokens=7000, temperature=0.9))
            if not isinstance(d, dict) or not str(d.get("human_text", "")).strip():
                continue
            ratio = difflib.SequenceMatcher(None, orig_words, _words(d["human_text"]), autojunk=False).ratio()
            bad = sorted({m.lower() for m in PARASITE_RE.findall(d["human_text"])})
            score = ratio + (0.5 if bad else 0)
            if score < best_ratio:
                best, best_ratio, best_pct = d, score, ratio
            if ratio <= 0.5 and not bad:
                break
            extra = "\n\nВНИМАНИЕ: исправь прошлую попытку."
            if ratio > 0.5:
                extra += f" Она совпала с оригиналом на {int(ratio * 100)}% — перепиши РАДИКАЛЬНЕЕ: другие слова, другая структура фраз, другие начала предложений."
            if bad:
                extra += " В тексте остались слова-паразиты: " + ", ".join(bad) + " — убери их, сохрани выбранный стиль и этикет."
        if not best:
            raise HTTPException(502, "ИИ вернул неверный формат. Попробуйте ещё раз.")
        best["changes"] = [c for c in (best.get("changes") or []) if isinstance(c, dict)
                           and _words(str(c.get("before", ""))) != _words(str(c.get("after", "")))]
        best["changed_pct"] = max(0, min(100, round((1 - best_pct) * 100)))
        return best

    d, left = paid(user, run)
    return {"result": d, "tokens_left": left}


# ==================== ПОСТУПЛЕНИЕ ПОСЛЕ 11 КЛАССА ====================
# Проверенные данные: общий конкурс на грант 2025 (минимальный балл получивших грант) по univision.kz,
# официальный сайт КБТУ и заявление МНВО. kind: grant — проходной на грант; threshold — порог вуза (реальный проходной обычно выше).
# Обновляйте раз в год после публикации итогов конкурса грантов.
KB_YEAR = 2025
KB = {
    "kbtu_it": ("КБТУ", "Информационные технологии", 110, "grant", "other", True),
    "kaznu_it": ("КазНУ им. аль-Фараби", "Информационные технологии (порог вуза)", 90, "threshold", "other", True),
    "satbayev_it": ("Satbayev University", "Информационные технологии (порог вуза)", 85, "threshold", "other", True),
    "narxoz_it": ("Университет Нархоз", "Информационные технологии (порог вуза)", 105, "threshold", "other", False),
    "satbayev_it_rural": ("Satbayev University", "Информационные технологии, сельская квота", 75, "grant", "other", True),
    "astana_med": ("Медицинский университет Астана", "Общая медицина", 128, "grant", "health", True),
    "astana_ped": ("Медицинский университет Астана", "Педиатрия", 122, "grant", "health", True),
    "astana_pharm": ("Медицинский университет Астана", "Фармация", 126, "grant", "health", True),
    "kaznmu_med": ("КазНМУ им. Асфендиярова", "Общая медицина", 124, "grant", "health", True),
    "kaznmu_ped": ("КазНМУ им. Асфендиярова", "Педиатрия", 113, "grant", "health", True),
    "kaznmu_pharm": ("КазНМУ им. Асфендиярова", "Фармация", 112, "grant", "health", True),
    "kaznmu_tfp": ("КазНМУ им. Асфендиярова", "Технология фармацевтического производства", 75, "grant", "health", True),
    "semey_med": ("Медицинский университет Семей", "Общая медицина", 112, "grant", "health", True),
    "semey_ped": ("Медицинский университет Семей", "Педиатрия", 105, "grant", "health", True),
    "semey_pharm": ("Медицинский университет Семей", "Фармация", 107, "grant", "health", True),
    "zkmu_med": ("ЗКМУ им. Оспанова (Актобе)", "Общая медицина", 112, "grant", "health", True),
    "zkmu_ped": ("ЗКМУ им. Оспанова (Актобе)", "Педиатрия", 106, "grant", "health", True),
    "krmu_med": ("Казахстанско-Российский медицинский университет", "Общая медицина", 112, "grant", "health", False),
    "skma_med": ("Южно-Казахстанская медицинская академия (Шымкент)", "Общая медицина", 113, "grant", "health", True),
    "skma_ped": ("Южно-Казахстанская медицинская академия (Шымкент)", "Педиатрия", 105, "grant", "health", True),
    "skma_tfp": ("Южно-Казахстанская медицинская академия (Шымкент)", "Технология фармацевтического производства", 89, "grant", "health", True),
    "kaznu_journ": ("КазНУ им. аль-Фараби", "Журналистика", 115, "grant", "other", True),
    "kaznu_philo": ("КазНУ им. аль-Фараби", "Филология", 118, "grant", "other", True),
    "kaznu_transl": ("КазНУ им. аль-Фараби", "Переводческое дело", 111, "grant", "other", True),
    "kaznu_media": ("КазНУ им. аль-Фараби", "Аудиовизуальные средства и медиа", 122, "grant", "other", True),
    "kaznu_psy": ("КазНУ им. аль-Фараби", "Психология", 100, "grant", "other", True),
    "kaznu_law": ("КазНУ им. аль-Фараби", "Право", 90, "grant", "law", True),
    "kaznu_fin": ("КазНУ им. аль-Фараби", "Финансы, экономика, банковское дело", 98, "grant", "other", True),
    "kaznu_mgmt": ("КазНУ им. аль-Фараби", "Менеджмент и управление", 114, "grant", "other", True),
    "kaznu_mkt": ("КазНУ им. аль-Фараби", "Маркетинг и реклама", 89, "grant", "other", True),
    "kaznu_math": ("КазНУ им. аль-Фараби", "Математика и статистика", 95, "grant", "other", True),
    "kaznu_bio": ("КазНУ им. аль-Фараби", "Биологические науки", 81, "grant", "other", True),
    "kaznu_chem": ("КазНУ им. аль-Фараби", "Химия", 80, "grant", "other", True),
    "kaznu_phys": ("КазНУ им. аль-Фараби", "Физика", 70, "grant", "other", True),
    "kaznu_pol": ("КазНУ им. аль-Фараби", "Политология", 86, "grant", "other", True),
    "kaznu_hist": ("КазНУ им. аль-Фараби", "История", 93, "grant", "other", True),
    "ped_math": ("Педагогические вузы РК", "Учитель математики (данные МНВО)", 124, "grant", "pedagogy", True),
}
KB_PROMPT = "\n".join(f"{k} | {v[0]} | {v[1]} | {v[2]} ({'проходной на грант' if v[3] == 'grant' else 'порог вуза'})" for k, v in KB.items())


def min_threshold(field: str, national: bool) -> int:
    return max(65 if national else 50, {"pedagogy": 75, "law": 75, "health": 70}.get(field, 0))


def chance_level(score: int, ref: float, kind: str) -> str:
    hi, lo = {"grant": (6, -5), "threshold": (12, 0), "estimate": (10, -8)}[kind]
    d = score - ref
    return "high" if d >= hi else ("low" if d < lo else "mid")


class CareerIn(BaseModel):
    answers: list[dict] = Field(max_length=8)
    score: int = Field(ge=0, le=140)
    rural: bool = False
    study_lang: str = "kz"
    lang: str = "ru"


@app.post("/api/career")
def career(body: CareerIn, user: dict = Depends(current_user)):
    ans = "\n".join(f"- {str(a.get('q', ''))[:200]} → {str(a.get('a', ''))[:200]}" for a in body.answers)
    prompt = (
        "Ты — школьный профориентолог и консультант по поступлению в вузы Казахстана. Ученик 11 класса прошёл мини-опрос.\n"
        f"ОТВЕТЫ:\n{ans}\nБалл ЕНТ (макс. 140): {body.score}. Сельская квота: {'да' if body.rural else 'нет'}. "
        f"Язык обучения: {body.study_lang}.\n\n"
        f"ПРОВЕРЕННАЯ БАЗА (данные {KB_YEAR}): id | вуз | программа | балл\n{KB_PROMPT}\n\n"
        "ЗАДАЧА: определи профиль ученика и предложи 8–9 вариантов поступления: частично из БАЗЫ (укажи kb_id), частично другие реальные "
        "вузы и программы Казахстана (kb_id = null). Включи варианты разной сложности: амбициозные, подходящие и запасные "
        "относительно балла ученика. Правила:\n"
        "• Если указан kb_id — числа НЕ пиши, сервер подставит их сам.\n"
        "• Если kb_id = null — дай est_low и est_high: честный ориентир проходного балла на грант (разброс не более 15 баллов). "
        "Не уверен в вузе или программе — не включай их. Не выдумывай точные цифры.\n"
        "• Для kb_id = null укажи field: pedagogy (педагогические науки), law, health (здравоохранение) или other, и national: true, "
        "если вуз национальный (КазНУ, ЕНУ, Satbayev, КазНПУ Абая, КазНМУ, КазНЖПУ и т. п.).\n"
        "• Учитывай сельскую квоту (около 30% грантов, баллы по ней обычно ниже) и язык обучения.\n"
        "• Профессии: 5 штук, что делает специалист и в каком направлении (ГОП) учиться. Без цифр зарплат.\n"
        "• tips: 4–5 коротких практических советов по подготовке к ЕНТ и подаче заявления. note: 1–2 предложения о конкурсе на грант.\n"
        + lang_line(body.lang) + "\n"
        'Верни ТОЛЬКО JSON: {"profile":{"title":"...","emoji":"🧠","summary":"2–3 предложения","strengths":["..."]},'
        '"professions":[{"name":"...","emoji":"💼","what":"...","field":"..."}],'
        '"options":[{"kb_id":null,"uni":"...","program":"...","field":"other","national":true,"est_low":0,"est_high":0,"why":"1 фраза"}],'
        '"tips":["..."],"note":"..."}'
    )

    def run():
        d = parse_json(ask_ai(prompt, json_mode=True, max_tokens=5000, temperature=0.6))
        if not isinstance(d, dict) or not isinstance(d.get("options"), list) or not d["options"]:
            raise HTTPException(502, "ИИ вернул неверный формат. Попробуйте ещё раз.")
        return d

    d, left = paid(user, run)
    order = {"high": 0, "mid": 1, "low": 2, "below": 3}
    options, seen = [], set()
    for o in d["options"][:12]:
        if not isinstance(o, dict):
            continue
        kb = KB.get(o.get("kb_id"))
        if kb:
            uni, prog, ref, kind, field, nat = kb
            ref_txt, verified = str(ref), True
        else:
            try:
                lo, hi = int(o.get("est_low")), int(o.get("est_high"))
            except (TypeError, ValueError):
                continue
            if not (40 <= lo <= hi <= 140) or hi - lo > 25 or not o.get("uni") or not o.get("program"):
                continue
            uni, prog = str(o["uni"])[:90], str(o["program"])[:120]
            field = o.get("field") if o.get("field") in ("pedagogy", "law", "health") else "other"
            nat, kind, ref, verified = bool(o.get("national")), "estimate", (lo + hi) / 2, False
            ref_txt = f"{lo}–{hi}"
        if (uni, prog) in seen:
            continue
        seen.add((uni, prog))
        need = min_threshold(field, nat)
        level = "below" if body.score < need else chance_level(body.score, ref, kind)
        options.append({"uni": uni, "program": prog, "ref": ref_txt, "ref_num": ref, "kind": kind, "verified": verified,
                        "chance": level, "need": need, "why": str(o.get("why", ""))[:200]})
    options.sort(key=lambda x: (order[x["chance"]], -x["ref_num"]))
    return {
        "profile": d.get("profile") or {}, "professions": (d.get("professions") or [])[:6],
        "options": options, "tips": (d.get("tips") or [])[:6], "note": str(d.get("note", ""))[:400],
        "year": KB_YEAR, "score": body.score, "tokens_left": left,
    }


# ==================== ПРЕЗЕНТАЦИИ (как Gamma) ====================
SLIDE_FONTS = ["Inter", "Playfair Display", "Space Grotesk", "Montserrat", "Manrope", "Unbounded", "Lora", "Oswald"]
SLIDE_LAYOUTS = {"title", "section", "bullets", "two_col", "stats", "quote", "timeline", "compare", "closing"}
SLIDE_DECOS = {"orbs", "grid", "waves", "lines", "none"}
HEX_RE = re.compile(r"^#[0-9a-fA-F]{6}$")


def _T(name, desc, bg, bg2, tx, mut, ac, ac2, fh, fb, deco):
    return {"name": name, "desc": desc, "bg": bg, "bg2": bg2, "tx": tx, "mut": mut, "ac": ac, "ac2": ac2, "fh": fh, "fb": fb, "deco": deco}


PRESET_STYLES = [
    _T("Neon Night", "Тёмный фон, неоновый лайм и фиолетовый — дерзко и современно", "#0b0b12", "#1a1a2e", "#f5f5f0", "#9a9aa8", "#c8ff2e", "#8b5cf6", "Space Grotesk", "Inter", "orbs"),
    _T("Clean Light", "Светлый минимализм, синий акцент — чисто и серьёзно", "#ffffff", "#eef2f7", "#0f172a", "#64748b", "#2563eb", "#06b6d4", "Manrope", "Inter", "lines"),
    _T("Editorial", "Тёплая бумага и засечки — как страница журнала", "#faf5ee", "#f0e4d3", "#1c1917", "#78716c", "#b45309", "#7c2d12", "Playfair Display", "Lora", "none"),
    _T("Aurora", "Глубокий бирюзовый и северное сияние — спокойно и дорого", "#042f2e", "#134e4a", "#ecfeff", "#99c6c3", "#5eead4", "#a78bfa", "Montserrat", "Inter", "waves"),
    _T("Sunset Pop", "Сливовый закат с розовым и золотым — ярко и смело", "#1f1235", "#4a1942", "#fff7ed", "#c9b3c9", "#fb7185", "#fbbf24", "Unbounded", "Manrope", "orbs"),
    _T("Corporate Navy", "Тёмно-синий и голубой — для бизнеса и защиты проектов", "#0f172a", "#1e3a8a", "#f8fafc", "#a5b4cf", "#38bdf8", "#f59e0b", "Montserrat", "Inter", "grid"),
]


def clean_theme(t, fb=None) -> dict:
    fb = fb or PRESET_STYLES[0]
    t = t if isinstance(t, dict) else {}
    out = {k: (t.get(k) if HEX_RE.match(str(t.get(k, ""))) else fb[k]) for k in ("bg", "bg2", "tx", "mut", "ac", "ac2")}
    out["fh"] = t.get("fh") if t.get("fh") in SLIDE_FONTS else fb["fh"]
    out["fb"] = t.get("fb") if t.get("fb") in SLIDE_FONTS else fb["fb"]
    out["deco"] = t.get("deco") if t.get("deco") in SLIDE_DECOS else fb["deco"]
    out["name"] = str(t.get("name") or fb["name"])[:40]
    out["desc"] = str(t.get("desc") or "")[:160]
    return out


THEME_SCHEMA = ('{"name":"...","desc":"1 фраза о настроении","bg":"#hex","bg2":"#hex","tx":"#hex","mut":"#hex","ac":"#hex","ac2":"#hex",'
                '"fh":"шрифт заголовков","fb":"шрифт текста","deco":"orbs|grid|waves|lines|none"}')
THEME_RULES = ("Правила темы: цвета только HEX из 6 символов; tx на bg и на bg2 с контрастом не ниже 4.5:1; mut — приглушённый текст, "
               "тоже читаемый; ac и ac2 — яркие акценты; шрифты ТОЛЬКО из списка: " + ", ".join(SLIDE_FONTS) + ".")


@app.post("/api/slides/styles")
def slides_styles(topic: str = Form(""), lang: str = Form("ru"), user: dict = Depends(current_user)):
    """Предлагает варианты оформления под тему. Бесплатно, с ограничением частоты."""
    rate_limit(f"styles:{user['id']}", 8, 600)
    prompt = (
        "Ты — арт-директор презентаций уровня Gamma и Pitch. Предложи 5 РАЗНЫХ красивых стилей оформления для презентации на тему: "
        f"«{topic.strip()[:400] or 'общая тема'}». Стили должны отличаться настроением (тёмный/светлый, строгий/дерзкий/тёплый) "
        "и подходить теме и аудитории. " + THEME_RULES + " Названия коротко и красиво (1–3 слова). " + lang_line(lang) +
        ' Верни ТОЛЬКО JSON: {"styles":[' + THEME_SCHEMA + "]}"
    )
    try:
        data = parse_json(ask_ai(prompt, json_mode=True, max_tokens=2000, temperature=1.0))
        items = data.get("styles") if isinstance(data, dict) else None
        styles = [clean_theme(x, PRESET_STYLES[i % len(PRESET_STYLES)]) for i, x in enumerate((items or [])[:5]) if isinstance(x, dict)]
    except Exception:
        log.exception("slides_styles failed")
        styles = []
    return {"styles": styles or PRESET_STYLES, "ai": bool(styles)}


class SlidesIn(BaseModel):
    prompt: str = Field(min_length=3, max_length=12000)
    count: int = Field(ge=3, le=12)
    mode: str = "write"            # write | shorten | verbatim
    style_text: str = Field(default="", max_length=400)
    theme: Optional[dict] = None
    lang: str = "ru"


def paid_n(user: dict, n: int, fn):
    """Списывает n токенов (по одному за слайд). При ошибке всё возвращается."""
    if int(user.get("tokens", 0)) < n:
        raise HTTPException(402, f"Нужно {n} токенов, у вас {int(user.get('tokens', 0))}. Выберите меньше слайдов или откройте Premium.")
    spent, left = 0, 0
    try:
        for _ in range(n):
            left = spend_token(user["id"])
            spent += 1
        result = fn()
    except Exception:
        for _ in range(spent):
            refund_token(user["id"])
        raise
    return result, left, spent


def _str(v, n=300) -> str:
    return str(v if v is not None else "").strip()[:n]


def _list(v, n, m=200):
    return [_str(x, m) for x in (v if isinstance(v, list) else [])[:n] if _str(x, m)]


def clean_slide(sl: dict) -> Optional[dict]:
    if not isinstance(sl, dict):
        return None
    layout = sl.get("layout") if sl.get("layout") in SLIDE_LAYOUTS else "bullets"
    out = {"layout": layout, "title": _str(sl.get("title"), 140), "subtitle": _str(sl.get("subtitle"), 220), "icon": _str(sl.get("icon"), 8) or "✨",
           "notes": _str(sl.get("notes"), 900), "bullets": _list(sl.get("bullets"), 6, 260)}
    for side in ("left", "right"):
        b = sl.get(side) if isinstance(sl.get(side), dict) else {}
        out[side] = {"h": _str(b.get("h"), 80), "items": _list(b.get("items"), 5, 220)}
    out["stats"] = [{"v": _str(x.get("v"), 14), "l": _str(x.get("l"), 90)} for x in (sl.get("stats") or [])[:4] if isinstance(x, dict) and _str(x.get("v"), 14)]
    out["steps"] = [{"h": _str(x.get("h"), 60), "t": _str(x.get("t"), 160)} for x in (sl.get("steps") or [])[:5] if isinstance(x, dict) and _str(x.get("h"), 60)]
    out["quote"], out["author"] = _str(sl.get("quote"), 400), _str(sl.get("author"), 80)
    # слайд без содержимого под свой layout → безопасный bullets
    need = {"stats": out["stats"], "timeline": out["steps"], "quote": out["quote"], "two_col": out["left"]["items"] or out["right"]["items"],
            "compare": out["left"]["items"] or out["right"]["items"]}
    if layout in need and not need[layout]:
        out["layout"] = "bullets"
    return out


MODE_RULES = {
    "write": "РЕЖИМ ТЕКСТА: пользователь дал идею или тему — ты сам пишешь весь контент: точно, конкретно, без воды.",
    "shorten": ("РЕЖИМ ТЕКСТА: пользователь дал готовый текст — СОКРАТИ его до тезисов. Сохрани все ключевые факты, числа, имена и порядок "
                "мыслей, убери повторы и воду. Пункт — до 12 слов."),
    "verbatim": ("РЕЖИМ ТЕКСТА: пользователь дал готовый текст — ОСТАВЬ ЕГО КАК ЕСТЬ. Запрещено перефразировать, сокращать, "
                 "исправлять или добавлять слова. Раздели текст на слайды по смыслу и копируй предложения и абзацы ДОСЛОВНО в bullets/items; "
                 "ты придумываешь только заголовок слайда, подзаголовок и layout. Если текст длиннее слайда — переноси остаток на следующий слайд."),
}


@app.post("/api/slides")
def slides_generate(body: SlidesIn, user: dict = Depends(current_user)):
    n = body.count
    mode = body.mode if body.mode in MODE_RULES else "write"
    preset = clean_theme(body.theme) if isinstance(body.theme, dict) else None
    theme_rule = ("ТЕМА ОФОРМЛЕНИЯ уже выбрана пользователем — НЕ придумывай её, в поле theme верни пустой объект {}."
                  if preset else
                  "ТЕМА ОФОРМЛЕНИЯ: придумай её сам" + (f" по описанию стиля пользователя: «{body.style_text.strip()}»" if body.style_text.strip()
                                                       else ", подобрав под тему и аудиторию") + ". " + THEME_RULES + " Формат theme: " + THEME_SCHEMA)
    prompt = (
        "Ты — главный дизайнер презентаций и сторителлер уровня TED, Apple и Gamma. Создай презентацию РОВНО из "
        f"{n} слайдов (ни больше, ни меньше).\n\n"
        f"ЗАПРОС ПОЛЬЗОВАТЕЛЯ:\n<<<\n{body.prompt.strip()}\n>>>\n\n{MODE_RULES[mode]}\n\n"
        "ДРАМАТУРГИЯ:\n"
        "• Слайд 1 — layout title: цепляющее название (до 8 слов) и подзаголовок: для кого и зачем.\n"
        f"• Слайд {n} — layout closing: сильный вывод или конкретный призыв к действию. Не «Спасибо за внимание».\n"
        "• Остальные: одна мысль — один слайд. Заголовок — это ВЫВОД, а не тема («Выручка выросла на 40% за год», а не «Выручка»), до 9 слов.\n"
        "• Логика: проблема → решение → доказательства → шаги, либо хронология. Без повторов между слайдами. "
        "Для длинных презентаций вставляй section для смены главы.\n\n"
        "LAYOUT (подбирай под смысл, чередуй, не ставь один и тот же больше двух раз подряд):\n"
        "title (обложка) · section (разделитель глав: только title и subtitle) · bullets (3–5 пунктов в bullets) · "
        "two_col (две колонки left и right: h — заголовок, items — 2–4 пункта) · stats (2–4 крупных факта в stats: v — короткое значение "
        "вроде «85%» или «3×», l — подпись) · quote (цитата в quote и author) · timeline (3–5 шагов в steps: h — короткая метка, t — пояснение) · "
        "compare (сравнение «до/после» или «A/B»: left и right) · closing (финал: title, subtitle, можно bullets из 2–3 шагов).\n\n"
        "КАЧЕСТВО ТЕКСТА:\n"
        "• Конкретика вместо общих слов: примеры, имена, механика. НИКОГДА не выдумывай цифры, даты и цитаты: если данных нет — "
        "формулируй качественно, а stats и quote используй только когда они действительно есть в запросе или общеизвестны.\n"
        "• Пункты начинай с сильного глагола или факта, без вводных слов и штампов («важно отметить», «в современном мире», «таким образом»).\n"
        "• Параллельная структура пунктов, без точек в конце. Не более 5 пунктов на слайд.\n"
        "• icon — один эмодзи, точно отражающий смысл слайда. notes — 2–3 предложения того, что докладчик говорит вслух "
        "(в режиме «как есть» notes содержат только пояснение для докладчика, а не изменённый текст слайда).\n\n"
        f"{theme_rule}\n{lang_line(body.lang)}\n\n"
        'Верни ТОЛЬКО JSON: {"title":"название презентации","theme":{...},"slides":[{"layout":"bullets","title":"...","subtitle":"","icon":"🚀",'
        '"bullets":["..."],"left":{"h":"","items":[]},"right":{"h":"","items":[]},"stats":[{"v":"","l":""}],'
        '"steps":[{"h":"","t":""}],"quote":"","author":"","notes":"..."}]}'
    )

    def run():
        raw = parse_json(ask_ai(prompt, json_mode=True, max_tokens=8000, temperature=0.8, timeout=170.0))
        slides = [x for x in (clean_slide(sl) for sl in (raw.get("slides") if isinstance(raw, dict) else None) or []) if x]
        if len(slides) < max(3, n - 1):
            raise HTTPException(502, "ИИ вернул неполную презентацию. Попробуйте ещё раз — токены вернутся.")
        theme = preset or clean_theme((raw or {}).get("theme"))
        return {"title": _str(raw.get("title"), 120) or slides[0]["title"], "theme": theme, "slides": slides[:n]}

    deck, left, spent = paid_n(user, n, run)
    extra = spent - len(deck["slides"])      # ИИ дал меньше слайдов — возвращаем разницу
    for _ in range(max(0, extra)):
        refund_token(user["id"]); left += 1
    save_generation(user["id"], "slides", deck["title"], json.dumps(deck, ensure_ascii=False))
    return {"deck": deck, "tokens_left": left, "charged": len(deck["slides"])}


def build_pptx(deck: dict) -> bytes:
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Inches, Pt

    th = clean_theme(deck.get("theme"))
    rgb = lambda h: RGBColor.from_string(h[1:].upper())
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]

    def box(sl, x, y, w, h, text, size, color, bold=False, font=None, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, italic=False):
        tf = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)).text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        lines = text if isinstance(text, list) else [text]
        for i, line in enumerate(lines):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.alignment = align
            para.space_after = Pt(8)
            run = para.add_run()
            run.text = line
            run.font.size, run.font.bold, run.font.italic = Pt(size), bold, italic
            run.font.color.rgb = rgb(color)
            run.font.name = font or th["fb"]
        return tf

    def shape(sl, kind, x, y, w, h, color):
        sh = sl.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
        sh.fill.solid()
        sh.fill.fore_color.rgb = rgb(color)
        sh.line.fill.background()
        return sh

    for idx, d in enumerate(deck.get("slides") or []):
        sl = prs.slides.add_slide(blank)
        sl.background.fill.solid()
        sl.background.fill.fore_color.rgb = rgb(th["bg"])
        shape(sl, MSO_SHAPE.RECTANGLE, 0, 0, 0.22, 7.5, th["ac"])
        lay, title = d.get("layout", "bullets"), d.get("title", "")
        if lay in ("title", "closing", "section"):
            shape(sl, MSO_SHAPE.OVAL, 9.6, -1.4, 5.2, 5.2, th["bg2"])
            shape(sl, MSO_SHAPE.OVAL, 10.9, 4.6, 3.4, 3.4, th["ac2"])
            box(sl, 1.0, 2.0, 9.0, 0.9, d.get("icon", ""), 44, th["ac"])
            box(sl, 1.0, 2.9, 9.4, 2.0, title, 46 if lay != "section" else 52, th["tx"], True, th["fh"], anchor=MSO_ANCHOR.TOP)
            if d.get("subtitle"):
                box(sl, 1.0, 5.0, 9.0, 1.0, d["subtitle"], 22, th["mut"])
            if lay == "closing" and d.get("bullets"):
                box(sl, 1.0, 5.9, 11, 1.2, ["→ " + b for b in d["bullets"][:3]], 16, th["ac"])
        else:
            box(sl, 0.9, 0.55, 11.6, 1.3, title, 34, th["tx"], True, th["fh"], anchor=MSO_ANCHOR.MIDDLE)
            shape(sl, MSO_SHAPE.RECTANGLE, 0.95, 1.9, 1.2, 0.07, th["ac"])
            if lay in ("two_col", "compare"):
                for j, side in enumerate(("left", "right")):
                    sd = d.get(side) or {}
                    x = 0.9 + j * 6.0
                    shape(sl, MSO_SHAPE.ROUNDED_RECTANGLE, x, 2.3, 5.6, 4.5, th["bg2"])
                    box(sl, x + 0.3, 2.45, 5.0, 0.7, sd.get("h", ""), 22, th["ac"] if j == 0 else th["ac2"], True, th["fh"])
                    box(sl, x + 0.3, 3.2, 5.0, 3.4, ["• " + t for t in sd.get("items", [])], 17, th["tx"])
            elif lay == "stats":
                st = d.get("stats") or []
                w = 11.4 / max(1, len(st))
                for j, it in enumerate(st):
                    x = 0.9 + j * w
                    shape(sl, MSO_SHAPE.ROUNDED_RECTANGLE, x, 2.5, w - 0.25, 3.6, th["bg2"])
                    box(sl, x + 0.1, 2.9, w - 0.45, 1.6, it["v"], 54, th["ac"] if j % 2 == 0 else th["ac2"], True, th["fh"], PP_ALIGN.CENTER)
                    box(sl, x + 0.2, 4.6, w - 0.65, 1.3, it["l"], 16, th["tx"], align=PP_ALIGN.CENTER)
            elif lay == "quote":
                box(sl, 1.2, 2.3, 10.8, 3.2, "“" + d.get("quote", "") + "”", 30, th["tx"], False, th["fh"], italic=True, anchor=MSO_ANCHOR.MIDDLE)
                box(sl, 1.2, 5.7, 10.8, 0.8, "— " + d.get("author", ""), 18, th["ac"])
            elif lay == "timeline":
                stp = d.get("steps") or []
                w = 11.6 / max(1, len(stp))
                shape(sl, MSO_SHAPE.RECTANGLE, 0.9, 3.05, 11.6, 0.05, th["mut"])
                for j, it in enumerate(stp):
                    x = 0.9 + j * w
                    dot = shape(sl, MSO_SHAPE.OVAL, x, 2.75, 0.65, 0.65, th["ac"] if j % 2 == 0 else th["ac2"])
                    box(sl, x, 2.78, 0.65, 0.6, str(j + 1), 18, th["bg"], True, th["fh"], PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
                    box(sl, x, 3.65, w - 0.2, 0.8, it["h"], 20, th["tx"], True, th["fh"])
                    box(sl, x, 4.45, w - 0.25, 2.2, it["t"], 14, th["mut"])
            else:
                box(sl, 0.95, 2.35, 11.4, 4.6, ["▸  " + b for b in d.get("bullets", [])], 22, th["tx"])
        if d.get("notes"):
            sl.notes_slide.notes_text_frame.text = d["notes"]
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


@app.post("/api/slides/export")
def slides_export(deck: dict = Body(...), user: dict = Depends(current_user)):
    slides = [x for x in (clean_slide(sl) for sl in (deck.get("slides") or [])[:14]) if x]
    if not slides:
        raise HTTPException(400, "Пустая презентация.")
    data = build_pptx({"theme": deck.get("theme"), "slides": slides})
    name = re.sub(r"[^\w\-]+", "_", str(deck.get("title") or "presentation"))[:40] or "presentation"
    return Response(data, media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    headers={"Content-Disposition": f'attachment; filename="{name}.pptx"'})


# ==================== АДМИН-ПАНЕЛЬ (подписки) ====================
def admin_user(user: dict = Depends(current_user)) -> dict:
    if user.get("telegram_id") not in ADMIN_IDS:
        raise HTTPException(403, "Нет доступа.")
    return user


@app.get("/api/admin/users")
def admin_users(q: str = "", _: dict = Depends(admin_user)):
    query = sb.table("profiles").select("id,telegram_id,username,display_name,tokens,score,is_subscribed,subscription_until,created_at")
    q = re.sub(r"[,()%*]", " ", q).strip().lstrip("@")[:40]
    if q:
        flt = f"username.ilike.%{q}%,display_name.ilike.%{q}%"
        if q.isdigit():
            flt += f",telegram_id.eq.{q}"
        query = query.or_(flt)
    return {"users": query.order("created_at", desc=True).limit(50).execute().data}


class SubIn(BaseModel):
    user_id: str
    active: bool


SUB_DAYS = 30
ADMIN_FIELDS = ("id", "telegram_id", "username", "display_name", "tokens", "score", "is_subscribed", "subscription_until", "created_at")


@app.post("/api/admin/subscription")
def admin_subscription(body: SubIn, _: dict = Depends(admin_user)):
    cur = sb.table("profiles").select("tokens").eq("id", body.user_id).limit(1).execute().data
    if not cur:
        raise HTTPException(404, "Пользователь не найден.")
    if body.active:   # включить на 30 дней
        upd: dict = {"is_subscribed": True, "tokens": max(cur[0]["tokens"], PRO_TOKENS),
                     "subscription_until": (datetime.now(timezone.utc) + timedelta(days=SUB_DAYS)).isoformat()}
    else:             # выключить: токены аннулируются
        upd = {"is_subscribed": False, "subscription_until": None, "tokens": 0,
               "tokens_refilled_on": today_utc().isoformat()}
    row = sb.table("profiles").update(upd).eq("id", body.user_id).execute().data[0]
    return {"user": {k: row.get(k) for k in ADMIN_FIELDS}}


# ==================== ЛИЧНЫЙ КАБИНЕТ ====================
def _dt(ts: str) -> datetime:
    return datetime.fromisoformat(ts.replace("Z", "+00:00"))


@app.get("/api/me")
def me(user: dict = Depends(current_user)):
    return public_profile(user)


@app.get("/api/me/overview")
def overview(tz: int = 0, user: dict = Depends(current_user)):
    """tz — смещение пользователя от UTC в минутах (Алматы = 300)."""
    uid = user["id"]
    tz = max(-840, min(tz, 840))
    off = timedelta(minutes=tz)
    local_today = (datetime.now(timezone.utc) + off).date()

    gens = (sb.table("generations").select("type").eq("user_id", uid).limit(5000).execute().data or [])
    by_type: dict[str, int] = {}
    for g in gens:
        by_type[g["type"]] = by_type.get(g["type"], 0) + 1

    chats = sb.table("chat_messages").select("id", count="exact").eq("user_id", uid).limit(1).execute().count or 0
    tasks = sb.table("tasks").select("done").eq("user_id", uid).execute().data or []

    since = (datetime.now(timezone.utc) - timedelta(days=60)).isoformat()
    events = (sb.table("xp_events").select("amount,created_at").eq("user_id", uid)
              .gte("created_at", since).limit(5000).execute().data or [])
    recent_gens = (sb.table("generations").select("created_at").eq("user_id", uid)
                   .gte("created_at", since).limit(2000).execute().data or [])

    xp_by_day: dict[date, int] = defaultdict(int)
    active_days: set[date] = set()
    for e in events:
        d = (_dt(e["created_at"]) + off).date()
        xp_by_day[d] += e["amount"]
        active_days.add(d)
    for g in recent_gens:
        active_days.add((_dt(g["created_at"]) + off).date())

    activity = []
    for i in range(13, -1, -1):
        d = local_today - timedelta(days=i)
        activity.append({"date": d.isoformat(), "xp": xp_by_day.get(d, 0)})

    streak = 0
    d = local_today if local_today in active_days else local_today - timedelta(days=1)
    while d in active_days:
        streak += 1
        d -= timedelta(days=1)

    return {
        "profile": public_profile(user),
        "counts": {
            "materials": len(gens), "chats": chats, "by_type": by_type,
            "tasks_total": len(tasks), "tasks_done": sum(1 for t in tasks if t["done"]),
        },
        "streak": streak,
        "activity": activity,
    }


@app.get("/api/history")
def history(type: str = "", q: str = "", limit: int = 100, user: dict = Depends(current_user)):
    query = sb.table("generations").select("id,type,title,created_at").eq("user_id", user["id"])
    if type in HISTORY_TYPES:
        query = query.eq("type", type)
    q = q.strip()[:60].replace("%", "").replace(",", " ")
    if q:
        query = query.ilike("title", f"%{q}%")
    items = query.order("created_at", desc=True).limit(max(1, min(limit, 200))).execute().data
    return {"items": items}


@app.get("/api/history/{item_id}")
def history_item(item_id: int, user: dict = Depends(current_user)):
    res = (sb.table("generations").select("id,type,title,content,created_at")
           .eq("id", item_id).eq("user_id", user["id"]).limit(1).execute().data)
    if not res:
        raise HTTPException(404, "Не найдено.")
    return res[0]


@app.delete("/api/history/{item_id}")
def history_delete(item_id: int, user: dict = Depends(current_user)):
    sb.table("generations").delete().eq("id", item_id).eq("user_id", user["id"]).execute()
    return {"ok": True}


@app.delete("/api/history")
def history_clear(user: dict = Depends(current_user)):
    sb.table("generations").delete().eq("user_id", user["id"]).execute()
    return {"ok": True}


@app.get("/api/chat/history")
def chat_history(limit: int = 50, user: dict = Depends(current_user)):
    rows = (sb.table("chat_messages").select("id,message,reply,created_at").eq("user_id", user["id"])
            .order("created_at", desc=True).limit(max(1, min(limit, 200))).execute().data)
    return {"items": rows}


@app.delete("/api/chat/history")
def chat_clear(user: dict = Depends(current_user)):
    sb.table("chat_messages").delete().eq("user_id", user["id"]).execute()
    return {"ok": True}


# ---------- задачи ----------
class TaskIn(BaseModel):
    title: str = Field(min_length=1, max_length=120)
    date: Optional[str] = None
    desc: str = Field(default="", max_length=300)


class TasksIn(BaseModel):
    tasks: list[TaskIn] = Field(min_length=1, max_length=10)


class TaskPatch(BaseModel):
    done: Optional[bool] = None
    title: Optional[str] = Field(default=None, min_length=1, max_length=120)
    date: Optional[str] = None


def safe_date(s: Optional[str]) -> str:
    try:
        return date.fromisoformat((s or "")[:10]).isoformat()
    except ValueError:
        return today_utc().isoformat()


def task_out(r: dict) -> dict:
    return {"id": r["id"], "title": r["title"], "desc": r.get("description") or "",
            "date": r["due_date"], "done": r["done"]}


@app.get("/api/tasks")
def tasks_list(user: dict = Depends(current_user)):
    rows = sb.table("tasks").select("*").eq("user_id", user["id"]).order("due_date").limit(500).execute().data
    return {"tasks": [task_out(r) for r in rows]}


@app.post("/api/tasks")
def tasks_create(body: TasksIn, user: dict = Depends(current_user)):
    rows = sb.table("tasks").insert([
        {"user_id": user["id"], "title": t.title.strip(), "description": t.desc.strip(),
         "due_date": safe_date(t.date)} for t in body.tasks
    ]).execute().data
    return {"tasks": [task_out(r) for r in rows]}


@app.patch("/api/tasks/{task_id}")
def tasks_patch(task_id: int, body: TaskPatch, user: dict = Depends(current_user)):
    cur = sb.table("tasks").select("*").eq("id", task_id).eq("user_id", user["id"]).limit(1).execute().data
    if not cur:
        raise HTTPException(404, "Задача не найдена.")
    cur = cur[0]
    upd: dict = {}
    if body.title is not None:
        upd["title"] = body.title.strip()
    if body.date is not None:
        upd["due_date"] = safe_date(body.date)
    awarded = 0
    score = user["score"]
    if body.done is not None:
        upd["done"] = body.done
        if body.done and not cur["xp_given"]:
            score = sb.rpc("add_xp", {"p_user": user["id"], "p_amount": 50, "p_reason": "task"}).execute().data
            upd["xp_given"] = True
            awarded = 50
    row = sb.table("tasks").update(upd).eq("id", task_id).execute().data[0] if upd else cur
    return {"task": task_out(row), "xp_awarded": awarded, "score": score}


@app.delete("/api/tasks/{task_id}")
def tasks_delete(task_id: int, user: dict = Depends(current_user)):
    sb.table("tasks").delete().eq("id", task_id).eq("user_id", user["id"]).execute()
    return {"ok": True}


@app.delete("/api/tasks")
def tasks_clear(user: dict = Depends(current_user)):
    sb.table("tasks").delete().eq("user_id", user["id"]).execute()
    return {"ok": True}


# ---------- XP ----------
class XpIn(BaseModel):
    amount: int = Field(ge=1, le=200)
    reason: str = Field(default="quiz", max_length=40)


@app.post("/api/xp")
def xp_add(body: XpIn, user: dict = Depends(current_user)):
    midnight = datetime.now(timezone.utc).replace(hour=0, minute=0, second=0, microsecond=0).isoformat()
    today_rows = (sb.table("xp_events").select("amount").eq("user_id", user["id"])
                  .gte("created_at", midnight).limit(2000).execute().data or [])
    if sum(r["amount"] for r in today_rows) + body.amount > DAILY_XP_CAP:
        raise HTTPException(429, "Дневной лимит XP достигнут.")
    score = sb.rpc("add_xp", {"p_user": user["id"], "p_amount": body.amount, "p_reason": body.reason}).execute().data
    return {"score": score}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
