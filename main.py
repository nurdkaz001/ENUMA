from __future__ import annotations

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
from fastapi.responses import HTMLResponse, JSONResponse
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
def ask_ai(prompt: str, json_mode: bool = False, max_tokens: int = 4000) -> str:
    """DeepSeek (OpenAI-совместимый API). В json_mode модель обязана вернуть JSON-ОБЪЕКТ."""
    payload = {
        "model": DEEPSEEK_MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": max_tokens,
        "temperature": 0.7,
    }
    if json_mode:
        payload["response_format"] = {"type": "json_object"}
    try:
        r = httpx.post(
            f"{DEEPSEEK_BASE_URL}/chat/completions",
            headers={"Authorization": f"Bearer {DEEPSEEK_API_KEY}"},
            json=payload, timeout=httpx.Timeout(90.0, connect=10.0),
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


@app.post("/api/humanize")
def humanize(text: str = Form(...), lang: str = Form("ru"), user: dict = Depends(current_user)):
    """Оценка «похоже на ИИ» + переписывание живым языком + список изменений. 1 токен."""
    text = text.strip()[:6000]
    if len(text) < 80:
        raise HTTPException(400, "Вставьте текст (минимум 80 символов).")
    prompt = (
        "Ты — опытный редактор и копирайтер, пишущий в живом, разговорном и публицистическом стиле. "
        "Выполни две задачи с текстом ниже.\n"
        "ЗАДАЧА 1 — анализ. Оцени по шкале 0–100, насколько текст похож на написанный нейросетью "
        "(ai_score), в одном предложении объясни вердикт, и найди до 6 признаков: фраза из текста (до 8 слов) и почему она выдаёт ИИ.\n"
        "ЗАДАЧА 2 — полностью перепиши текст (human_text), сохранив смысл, факты, числа и имена; ничего не выдумывай.\n"
        "ПРАВИЛА СТРУКТУРЫ: 1) чередуй длину предложений волнообразно: одно очень длинное (с причастным оборотом или "
        "вводной конструкцией), следом экстремально короткое (1–3 слова), например «Это работает именно так. Без вариантов.»; "
        "2) намеренно используй парцелляцию — разделяй связанные мысли точкой.\n"
        "ПРАВИЛА ЛЕКСИКИ: 1) полностью исключи маркерные слова ИИ: «важно отметить», «таким образом», «в заключение», "
        "«несомненно», «ключевой аспект», «стоит упомянуть», «погрузиться в», «экосистема», «инновационный», «трансформация»; "
        "2) заменяй канцеляризмы простыми глаголами («осуществляет процесс анализа» → «анализирует»); "
        "3) пиши только в активном залоге («мы сделали», «автор написал»); "
        "4) используй лёгкие разговорные связки: «кстати», «собственно говоря», «на самом деле», «плюс ко всему», «а значит».\n"
        "Язык результата — тот же, что у исходного текста (для не русского применяй те же принципы). "
        "human_text содержит ТОЛЬКО очеловеченный текст: без комментариев, вступлений и кавычек. "
        "Также оцени новый текст по той же шкале (human_score) и дай 4–6 пар фрагментов «было → стало» (changes), "
        "где before — кусок исходного текста, after — соответствующий кусок нового. "
        "Поля verdict и why пиши на языке интерфейса: " + lang_line(lang) + "\n"
        'Верни ТОЛЬКО JSON: {"ai_score":0,"verdict":"...","markers":[{"phrase":"...","why":"..."}],'
        '"human_text":"...","human_score":0,"changes":[{"before":"...","after":"..."}]}\n\nТЕКСТ:\n' + text
    )

    def run():
        d = parse_json(ask_ai(prompt, json_mode=True, max_tokens=7000))
        if not isinstance(d, dict) or not str(d.get("human_text", "")).strip():
            raise HTTPException(502, "ИИ вернул неверный формат. Попробуйте ещё раз.")
        return d

    d, left = paid(user, run)
    return {"result": d, "tokens_left": left}


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
