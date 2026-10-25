from __future__ import annotations

import io
import json
import logging
import os
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
DAILY_TOKENS = int(os.getenv("DAILY_TOKENS", "5"))
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
    # ежедневное пополнение токенов
    if str(user.get("tokens_refilled_on")) < today_utc().isoformat():
        refilled = sb.rpc("refill_tokens", {"p_user": user["id"], "p_amount": DAILY_TOKENS}).execute().data
        if refilled is not None:
            user["tokens"] = refilled
    return user


def public_profile(u: dict) -> dict:
    return {
        "user_id": u["id"], "telegram_id": u.get("telegram_id"),
        "username": u.get("username"), "display_name": u.get("display_name") or u.get("username"),
        "tokens": u.get("tokens", 0), "score": u.get("score", 0),
        "is_subscribed": bool(u.get("is_subscribed")), "created_at": u.get("created_at"),
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
                "display_name": name,
                "expires_at": (datetime.now(timezone.utc) + timedelta(minutes=CODE_TTL_MIN)).isoformat(),
            }).execute()
            break
        except Exception:
            code = ""
    if not code:
        tg_send(chat_id, "Не удалось создать код, попробуйте ещё раз.")
        return {"ok": True}
    tg_send(chat_id, f"Ваш код для входа: {code}\nДействует {CODE_TTL_MIN} минут.")
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
def ask_ai(prompt: str, json_mode: bool = False) -> str:
    """DeepSeek (OpenAI-совместимый API). В json_mode модель обязана вернуть JSON-ОБЪЕКТ."""
    payload = {
        "model": DEEPSEEK_MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": 4000,
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


@app.post("/upload")
def upload_pdf(file: UploadFile = File(...), user: dict = Depends(current_user)):
    if not (file.filename or "").lower().endswith(".pdf"):
        raise HTTPException(400, "Пожалуйста, загрузите файл в формате PDF.")
    data = file.file.read(MAX_UPLOAD_BYTES + 1)
    if len(data) > MAX_UPLOAD_BYTES:
        raise HTTPException(413, "Файл больше 15 МБ.")
    try:
        reader = PdfReader(io.BytesIO(data))
        if reader.is_encrypted:
            raise HTTPException(400, "PDF защищён паролем.")
        text = "\n".join(filter(None, (p.extract_text() for p in reader.pages)))
    except HTTPException:
        raise
    except Exception:
        raise HTTPException(400, "Не удалось прочитать PDF.")
    if len(text.strip()) < 30:
        raise HTTPException(400, "В PDF не найден текст (возможно, это скан).")

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
    # ИИ-планер: бесплатный, промпт собирает фронтенд, ответ — JSON-объект {"tasks": [...]}
    if action == "planner":
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
    amount: int = Field(ge=1, le=50)
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
