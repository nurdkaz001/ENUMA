from __future__ import annotations

import difflib
import io
import json
import logging
import os
import random
import re
import base64
import functools
import math
import secrets
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from collections import OrderedDict, defaultdict, deque
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Optional
from urllib.parse import parse_qs, quote, urlparse

import httpx
import jwt
from fastapi import Body, Depends, FastAPI, File, Form, Header, HTTPException, Request, UploadFile
from fastapi.exceptions import RequestValidationError
from fastapi.responses import HTMLResponse, JSONResponse, Response, StreamingResponse
from pydantic import BaseModel, Field
import openpyxl
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from pypdf import PdfReader
from starlette.exceptions import HTTPException as StarletteHTTPException
from supabase import Client, create_client

# ==================== НАСТРОЙКИ ====================
log = logging.getLogger("studyai")
logging.basicConfig(level=logging.INFO)

DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "")
DEEPSEEK_MODEL = os.getenv("DEEPSEEK_MODEL", "deepseek-chat")
DEEPSEEK_BASE_URL = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com").rstrip("/")
# Модель с поддержкой изображений (любой OpenAI-совместимый API). DeepSeek Chat картинки не понимает.
VISION_API_KEY = os.getenv("VISION_API_KEY", "")
VISION_BASE_URL = os.getenv("VISION_BASE_URL", "https://generativelanguage.googleapis.com/v1beta/openai").rstrip("/")
VISION_MODEL = os.getenv("VISION_MODEL", "gemini-2.0-flash")
SUPABASE_URL = os.getenv("SUPABASE_URL", "")
SUPABASE_SERVICE_KEY = os.getenv("SUPABASE_SERVICE_KEY", "")
JWT_SECRET = os.getenv("JWT_SECRET", "")
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_WEBHOOK_SECRET = os.getenv("TELEGRAM_WEBHOOK_SECRET", "")
WELCOME_TOKENS = int(os.getenv("WELCOME_TOKENS", "50"))      # подарок при первой регистрации
DAILY_TOKENS = int(os.getenv("DAILY_TOKENS", "3"))            # бесплатно каждый день (добавляется к балансу)
PRO_TOKENS = int(os.getenv("PRO_TOKENS", "50"))          # токенов в день по подписке
ADMIN_IDS = {int(x) for x in re.split(r"[,\s]+", os.getenv("ADMIN_TELEGRAM_IDS", "")) if x.isdigit()}
MAX_DOC_CHARS = int(os.getenv("MAX_DOC_CHARS", "70000"))   # сколько текста отдаём модели (DeepSeek держит 64K+ токенов контекста)
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

app = FastAPI(title="EnumaAI Hub")
sb: Client = create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)


def _patch_supabase_retries() -> None:
    """Supabase иногда рвёт keep-alive соединение — из-за этого случайно падали запросы.
    Читающие запросы повторяем при любых сетевых сбоях, пишущие — только если запрос заведомо не дошёл."""
    try:
        import postgrest._sync.request_builder as rb
    except Exception:
        return
    connect_only = (httpx.ConnectError, httpx.ConnectTimeout, httpx.PoolTimeout)
    flaky = (httpx.ReadError, httpx.ReadTimeout, httpx.RemoteProtocolError, httpx.WriteError, httpx.WriteTimeout)

    def wrap(orig):
        @functools.wraps(orig)
        def wrapped(self, *a, **kw):
            method = str(getattr(self, "http_method", "GET")).upper()
            idempotent = method in ("GET", "HEAD", "PATCH", "PUT", "DELETE")
            err = None
            for attempt in range(3):
                try:
                    return orig(self, *a, **kw)
                except connect_only as e:
                    err = e
                except flaky as e:
                    if not idempotent:
                        raise
                    err = e
                time.sleep(0.35 * (attempt + 1))
            raise err
        wrapped._retry = True
        return wrapped

    for name in dir(rb):
        cls = getattr(rb, name)
        ex = cls.__dict__.get("execute") if isinstance(cls, type) else None
        if ex and not getattr(ex, "_retry", False):
            cls.execute = wrap(ex)


_patch_supabase_retries()
INDEX_FILE = Path(__file__).parent / "templates" / "index.html"
STATIC_DIR = Path(__file__).parent / "static"

HISTORY_TYPES = {"quiz", "flashcards", "summary", "simple", "plan", "teacher", "tickets", "abroad", "letter"}
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


@app.exception_handler(httpx.TransportError)
async def network_error(_: Request, exc: Exception):
    log.warning("network error: %s", type(exc).__name__)
    return JSONResponse({"error": "Связь с сервером данных нестабильна. Повторите через пару секунд."}, status_code=503)


try:
    from postgrest.exceptions import APIError as _PgError

    @app.exception_handler(_PgError)
    async def db_error(_: Request, exc: Exception):
        code = str(getattr(exc, "code", "") or "")
        log.error("DB error %s: %s", code, getattr(exc, "message", exc))
        hint = " Похоже, не выполнена SQL-миграция (migration_v3.sql / migration_v4.sql)." if code in ("42P01", "42703", "PGRST205", "PGRST204") else ""
        return JSONResponse({"error": f"Ошибка базы данных ({code or '—'}).{hint}"}, status_code=500)
except Exception:        # другая версия библиотеки — остаётся общий обработчик
    pass


# ==================== АВТОРИЗАЦИЯ ====================
def make_token(user_id: str) -> str:
    exp = datetime.now(timezone.utc) + timedelta(days=JWT_DAYS)
    return jwt.encode({"sub": user_id, "exp": exp}, JWT_SECRET, algorithm="HS256")


def today_utc() -> date:
    return datetime.now(timezone.utc).date()


_ACTIVE_DAY: dict[str, str] = {}


def _touch_activity(user: dict) -> None:
    """Записывает, что пользователь сегодня заходил (локальный день по его часовому поясу) — на этом строится серия 🔥.
    Пишем не чаще раза в сутки на пользователя; данные не зависят от истории материалов."""
    try:
        tz = user.get("tz_offset")
        day = (datetime.now(timezone.utc) + timedelta(hours=5 if tz is None else int(tz))).date().isoformat()
        if _ACTIVE_DAY.get(user["id"]) == day:
            return
        if len(_ACTIVE_DAY) > 20000:
            _ACTIVE_DAY.clear()
        _ACTIVE_DAY[user["id"]] = day
        sb.table("activity_days").upsert({"user_id": user["id"], "day": day}, on_conflict="user_id,day").execute()
    except Exception as e:
        log.warning("activity not recorded (%s) — выполните migration_v8.sql", type(e).__name__)


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
        refilled = sb.rpc("refill_tokens", {"p_user": user["id"], "p_amount": DAILY_TOKENS}).execute().data
        if refilled is not None:
            user["tokens"] = refilled
    _touch_activity(user)
    return user


def login_profile(row: dict) -> tuple[dict, int]:
    """Находит профиль по Telegram ID или создаёт новый. Возвращает (профиль, размер приветственного подарка: 0, если профиль уже был).
    Подарок выдаётся один раз на Telegram-аккаунт: удаление аккаунта и повторная регистрация его не возвращают."""
    tg = row["telegram_id"]
    names = {"username": row.get("username"), "display_name": row.get("display_name")}
    found = sb.table("profiles").select("*").eq("telegram_id", tg).limit(1).execute().data
    if found:
        p = found[0]
        if (p.get("username"), p.get("display_name")) != (names["username"], names["display_name"]):
            sb.table("profiles").update(names).eq("id", p["id"]).execute()
            p.update(names)
        return p, 0
    tokens = WELCOME_TOKENS
    try:
        sb.table("welcome_grants").insert({"telegram_id": tg}).execute()
    except Exception as e:
        if "23505" in str(e) or "duplicate" in str(e).lower():        # уже получал подарок раньше (удалил аккаунт и зарегистрировался снова)
            tokens = DAILY_TOKENS
        else:
            log.warning("welcome_grants unavailable (%s) — выполните migration_v7.sql", type(e).__name__)
    try:
        p = sb.table("profiles").insert({"telegram_id": tg, **names, "tokens": tokens}).execute().data[0]
    except Exception:                                                   # параллельный вход уже создал профиль
        again = sb.table("profiles").select("*").eq("telegram_id", tg).limit(1).execute().data
        if not again:
            raise
        return again[0], 0
    return p, (tokens if tokens > DAILY_TOKENS else 0)


def public_profile(u: dict) -> dict:
    return {
        "user_id": u["id"], "telegram_id": u.get("telegram_id"),
        "username": u.get("username"), "display_name": u.get("display_name") or u.get("username"),
        "tokens": u.get("tokens", 0), "score": u.get("score", 0),
        "is_subscribed": bool(u.get("is_subscribed")), "is_admin": u.get("telegram_id") in ADMIN_IDS, "subscription_until": u.get("subscription_until"), "created_at": u.get("created_at"),
        "consent_version": u.get("consent_version"), "consent_at": u.get("consent_at"),
        "public_id": u.get("public_id"), "telegram_id": u.get("telegram_id"), "onboarded": ("onboarded_at" not in u) or bool(u.get("onboarded_at")),
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

    prof, bonus = login_profile(row)
    return {**public_profile(prof), "token": make_token(prof["id"]), "welcome_bonus": bonus}


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
    prof, bonus = login_profile(row)
    return {"ready": True, **public_profile(prof), "token": make_token(prof["id"]), "welcome_bonus": bonus}


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
def _deepseek(messages: list, json_mode: bool, max_tokens: int, temperature: float, timeout: float) -> tuple[str, str]:
    payload = {"model": DEEPSEEK_MODEL, "messages": messages, "max_tokens": max_tokens, "temperature": temperature}
    if json_mode:
        payload["response_format"] = {"type": "json_object"}
    try:
        r = httpx.post(f"{DEEPSEEK_BASE_URL}/chat/completions", headers={"Authorization": f"Bearer {DEEPSEEK_API_KEY}"},
                       json=payload, timeout=httpx.Timeout(timeout, connect=10.0))
        if r.status_code == 402:
            log.error("DeepSeek: недостаточно средств на балансе")
            raise HTTPException(502, "ИИ временно недоступен (проверьте баланс DeepSeek).")
        r.raise_for_status()
        ch = r.json()["choices"][0]
        return (ch["message"]["content"] or "").strip(), ch.get("finish_reason") or ""
    except HTTPException:
        raise
    except Exception:
        log.exception("deepseek failed")
        raise HTTPException(502, "ИИ сейчас недоступен, попробуйте ещё раз.")


def ask_ai(prompt: str, json_mode: bool = False, max_tokens: int = 4000, temperature: float = 0.7, timeout: float = 90.0) -> str:
    """DeepSeek (OpenAI-совместимый API). В json_mode модель обязана вернуть JSON-ОБЪЕКТ.
    Если текстовый ответ оборвался на лимите токенов — просим продолжить (до 2 раз), чтобы не терять конец."""
    msgs = [{"role": "user", "content": prompt}]
    text, finish = _deepseek(msgs, json_mode, max_tokens, temperature, timeout)
    if not text:
        raise HTTPException(502, "ИИ вернул пустой ответ.")
    for _ in range(2):
        if json_mode or finish != "length":
            break
        more, finish = _deepseek(msgs + [{"role": "assistant", "content": text},
                                         {"role": "user", "content": "Продолжи ровно с того места, где оборвался ответ. Без повторов и вступлений."}],
                                 False, max_tokens, temperature, timeout)
        if not more:
            break
        text += more if text.endswith((" ", "\n")) or more.startswith((" ", "\n")) else " " + more
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


def fit_context(text: str, limit: int = 0) -> str:
    """Если материал не помещается в лимит — берём выдержки из ВСЕХ его частей (а не обрезаем хвост), чтобы ИИ видел весь документ."""
    limit = limit or MAX_DOC_CHARS
    text = (text or "").strip()
    if len(text) <= limit:
        return text
    parts = 8
    per = limit // parts - 60
    step = len(text) / parts
    out = [f"[Материал большой ({len(text)} символов): ниже выдержки из всех {parts} частей документа по порядку. Охвати их все.]"]
    for i in range(parts):
        if i == parts - 1:                                   # из последней части берём КОНЕЦ документа: там обычно выводы
            chunk = text[-per:]
            nl = chunk.find("\n")
            if 0 <= nl < per * 0.3:
                chunk = chunk[nl + 1:]
        else:
            start = int(i * step)
            chunk = text[start:start + per]
            cut = chunk.rfind("\n")
            if cut > per * 0.6:
                chunk = chunk[:cut]
        out.append(f"[Часть {i + 1} из {parts}]\n{chunk.strip()}")
    return "\n\n".join(out)


MATERIAL_RULES = (
    "СТРОГИЕ ПРАВИЛА РАБОТЫ С МАТЕРИАЛОМ (соблюдай все):\n"
    "1. Изучи материал ПОЛНОСТЬЮ — от первой до последней строки, включая таблицы, формулы, определения, даты и примеры. "
    "Ничего не пропускай: охвати каждый раздел (маркеры [Страница N] и [Часть N из M] показывают, где ты находишься).\n"
    "2. Опирайся только на материал. Не выдумывай факты, числа, имена и формулы. Чего в материале нет — не утверждай; "
    "общеизвестные пояснения допустимы только с пометкой «Дополнение».\n"
    "3. Числа, имена, даты, термины и формулы переноси точно, без искажений и округлений.\n"
    "4. Перед ответом мысленно сверь результат с материалом: все ли разделы отражены, нет ли противоречий, неточностей и выдуманных деталей.\n"
    "5. Пиши грамотно: без орфографических, грамматических и стилистических ошибок, без повторов и заглушек («и т.д.», «...»).\n"
    "6. Доводи ответ до конца: не обрывай списки и разделы на полуслове.\n"
)
QUESTION_RULES = (
    "ПРАВИЛА ДЛЯ ВОПРОСОВ: равномерно охвати ВЕСЬ материал от начала до конца (не больше двух вопросов на один абзац); "
    "каждый вопрос проверяем по тексту, ответ однозначен и есть в материале; неверные варианты правдоподобны, но однозначно неверны; "
    "в вопросах, вариантах и пояснениях — точные числа, термины и формулировки из материала, без ошибок.\n"
)


def _verify_questions(items: list, material: str, lang: str) -> list:
    """Второй проход: сверяет вопросы с материалом, исправляет неверный ответ и ошибки текста, убирает неоднозначные."""
    if not material or not items:
        return items
    key = "correct" if "correct" in items[0] else "answer"
    listing = "\n".join(
        f"{i + 1}. {it.get('q', '')}\n" + "\n".join(str(o) for o in it.get("options", [])) + f"\nУказан верный: {it.get(key, '')}\nПояснение: {it.get('explain', '')}"
        for i, it in enumerate(items))
    prompt = (
        "Ты — строгий эксперт-проверяющий. Ниже МАТЕРИАЛ и тестовые вопросы по нему. Проверь КАЖДЫЙ вопрос:\n"
        "1) верен ли указанный правильный вариант по материалу (если нет — укажи верную букву);\n"
        "2) нет ли второго правильного варианта и двусмысленности (если есть — в drop);\n"
        "3) нет ли орфографических, фактических и смысловых ошибок в вопросе и пояснении (если есть — дай исправленный текст);\n"
        "4) ответ на вопрос должен находиться в материале; если его там нет — в drop.\n"
        'Верни ТОЛЬКО JSON: {"fixes":[{"i":1,"correct":"B","q":"исправленный вопрос или пустая строка","explain":"исправленное пояснение или пустая строка"}],"drop":[номера]}. '
        'Если всё верно — {"fixes":[],"drop":[]}. Не меняй то, что верно. ' + lang_line(lang) + "\n\n"
        f"МАТЕРИАЛ:\n{fit_context(material, 60000)}\n\nВОПРОСЫ:\n{listing}"
    )
    try:
        d = parse_json(ask_ai(prompt, json_mode=True, max_tokens=3000, temperature=0.1, timeout=100.0))
    except Exception:
        log.warning("verify skipped")
        return items
    if not isinstance(d, dict):
        return items
    for f in d.get("fixes") or []:
        try:
            it = items[int(f["i"]) - 1]
        except (KeyError, ValueError, TypeError, IndexError):
            continue
        c = str(f.get("correct") or "").strip().upper()[:1]
        if c in "ABCD" and c:
            it[key] = c
        if str(f.get("q") or "").strip():
            it["q"] = _str(f["q"], 400)
        if str(f.get("explain") or "").strip():
            it["explain"] = _str(f["explain"], 500)
    drop = {int(x) - 1 for x in (d.get("drop") or []) if str(x).isdigit()}
    keep = [it for i, it in enumerate(items) if i not in drop]
    return keep if len(keep) >= max(4, int(len(items) * 0.6)) else items


def build_context(doc_id: str, title: str, user: dict) -> str:
    text = get_doc_text(doc_id, user)
    if text is None:
        raise HTTPException(400, "Документ не найден. Загрузите файл заново.")
    if not text and title and title.strip() and title.strip() != "Lesson":
        text = f"Тема для изучения: {title.strip()[:200]}"
    if not text:
        raise HTTPException(400, "Сначала загрузите PDF или введите тему.")
    return fit_context(text)


def lang_line(lang: str) -> str:
    return LANG_LINE.get(lang, LANG_LINE["ru"])


_KZ_LETTERS = re.compile(r"[әіңғүұқөһӘІҢҒҮҰҚӨҺ]")
_CYR = re.compile(r"[А-Яа-яЁё]")
_LAT = re.compile(r"[A-Za-z]")


def detect_lang(text: str) -> Optional[str]:
    """Язык запроса по алфавиту: kz (есть казахские буквы), ru (кириллица), en (латиница). None — если букв почти нет (цифры, формулы, код)."""
    t = text or ""
    kz = len(_KZ_LETTERS.findall(t))
    cyr = len(_CYR.findall(t)) + kz
    lat = len(_LAT.findall(t))
    if cyr + lat < 2:
        return None
    if cyr >= lat:
        return "kz" if kz >= 1 else "ru"
    return "en"


def request_lang_line(text: str, ui_lang: str = "ru") -> str:
    """Отвечать надо на языке самого запроса, а не интерфейса сайта."""
    base = ("ЯЗЫК ОТВЕТА: отвечай на том же языке, на котором написан запрос студента, независимо от выбранного языка интерфейса сайта. "
            "Если в запросе смешаны языки — на языке основной части. Ключи JSON и служебные поля не переводи. ")
    d = detect_lang(text)
    if d == "ru":
        hint = "Запрос написан по-русски — отвечай по-русски."
    elif d == "kz":
        hint = "Запрос написан по-казахски — отвечай на казахском языке (Жауапты қазақ тілінде бер)."
    elif d == "en":
        hint = "The request is written in English (or another Latin-script language) — answer in the same language as the request."
    else:
        hint = "Если язык запроса определить невозможно (только цифры, формулы или код) — " + LANG_LINE.get(ui_lang, LANG_LINE["ru"])
    return base + hint


# ==================== СТРАНИЦА / ЗАГРУЗКА / ГЕНЕРАЦИЯ ====================
@app.get("/", response_class=HTMLResponse)
def read_root():
    return HTMLResponse(INDEX_FILE.read_text(encoding="utf-8"))


@app.get("/privacy", response_class=HTMLResponse)
@app.get("/terms", response_class=HTMLResponse)
def read_privacy():
    return HTMLResponse(INDEX_FILE.read_text(encoding="utf-8"))


@app.get("/manifest.webmanifest")
def pwa_manifest():
    """Манифест: чтобы сайт можно было поставить на экран телефона как приложение."""
    data = {
        "name": "EnumaAI Hub", "short_name": "EnumaAI",
        "description": "ИИ-наставник: тесты, конспекты, планер, поступление",
        "start_url": "/?source=pwa", "scope": "/", "display": "standalone",
        "background_color": "#131314", "theme_color": "#131314", "lang": "ru",
        "icons": [
            {"src": "/static/icon-192.png", "sizes": "192x192", "type": "image/png", "purpose": "any"},
            {"src": "/static/icon-512.png", "sizes": "512x512", "type": "image/png", "purpose": "any"},
            {"src": "/static/icon-maskable-512.png", "sizes": "512x512", "type": "image/png", "purpose": "maskable"},
        ],
    }
    return Response(json.dumps(data, ensure_ascii=False), media_type="application/manifest+json",
                    headers={"Cache-Control": "public, max-age=3600"})


@app.get("/sw.js")
def pwa_sw():
    """Минимальный service worker (нужен Chrome для установки). Ничего не кэширует, API не трогает."""
    js = "self.addEventListener('install',e=>self.skipWaiting());self.addEventListener('activate',e=>e.waitUntil(self.clients.claim()));self.addEventListener('fetch',()=>{});"
    return Response(js, media_type="application/javascript", headers={"Cache-Control": "no-cache", "Service-Worker-Allowed": "/"})


@app.get("/static/{name}")
def static_file(name: str, request: Request):
    """Раздача медиа из папки static/ (аватар-анимация и т.п.) с поддержкой Range — без него видео не играет в Safari/iPhone."""
    ctypes = {".mp4": "video/mp4", ".webm": "video/webm", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png", ".webp": "image/webp"}
    ext = Path(name).suffix.lower()
    f = STATIC_DIR / name
    if not re.fullmatch(r"[\w.-]{1,80}", name) or ext not in ctypes or not f.is_file():
        raise HTTPException(404, "Not found")
    data, size = f.read_bytes(), f.stat().st_size
    headers = {"Accept-Ranges": "bytes", "Cache-Control": "public, max-age=86400"}
    m = re.fullmatch(r"bytes=(\d*)-(\d*)", (request.headers.get("range") or "").strip())
    if m and (m.group(1) or m.group(2)):
        a, b = m.groups()
        start = int(a) if a else max(0, size - int(b))
        end = min(int(b) if (a and b) else size - 1, size - 1)
        if start > end or start >= size:
            return Response(status_code=416, headers={"Content-Range": f"bytes */{size}"})
        return Response(data[start:end + 1], status_code=206, media_type=ctypes[ext],
                        headers={**headers, "Content-Range": f"bytes {start}-{end}/{size}"})
    return Response(data, media_type=ctypes[ext], headers=headers)


@app.get("/health")
def health():
    return {"ok": True}


def read_upload_ex(file: UploadFile) -> tuple[str, dict]:
    """Достаёт текст из PDF, Word (.docx), Excel (.xlsx) или .txt. Возвращает текст и сведения о качестве чтения."""
    name = (file.filename or "").lower()
    data = file.file.read(MAX_UPLOAD_BYTES + 1)
    if len(data) > MAX_UPLOAD_BYTES:
        raise HTTPException(413, "Файл больше 15 МБ.")
    parts: list[str] = []
    meta: dict = {"pages": 0, "empty_pages": 0}
    try:
        if name.endswith(".pdf"):
            reader = PdfReader(io.BytesIO(data))
            if reader.is_encrypted:
                raise HTTPException(400, "PDF защищён паролем.")
            meta["pages"] = len(reader.pages)
            for i, page in enumerate(reader.pages, 1):
                try:
                    t = (page.extract_text() or "").strip()
                except Exception:
                    t = ""
                if t:
                    parts.append(f"[Страница {i}]\n{t}")
                else:
                    meta["empty_pages"] += 1
        elif name.endswith(".docx"):
            d = Document(io.BytesIO(data))
            for child in d.element.body.iterchildren():             # порядок как в документе: абзацы и таблицы на своих местах
                tag = child.tag.rsplit("}", 1)[-1]
                if tag == "p":
                    para = Paragraph(child, d)
                    if para.text.strip():
                        style = (para.style.name if para.style is not None else "") or ""
                        parts.append(("## " if style.lower().startswith(("heading", "заголовок", "title")) else "") + para.text.strip())
                elif tag == "tbl":
                    for row in Table(child, d).rows:
                        cells = []
                        for c in row.cells:
                            tx = c.text.strip().replace("\n", " ")
                            if tx and (not cells or cells[-1] != tx):         # объединённые ячейки не дублируем
                                cells.append(tx)
                        if cells:
                            parts.append(" | ".join(cells))
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
            raw = data
            for enc in ("utf-8-sig", "cp1251", "utf-8"):
                try:
                    parts = [raw.decode(enc)]
                    break
                except UnicodeDecodeError:
                    continue
            else:
                parts = [raw.decode("utf-8", "ignore")]
        else:
            raise HTTPException(400, "Поддерживаются PDF, Word (.docx), Excel (.xlsx) и .txt.")
    except HTTPException:
        raise
    except Exception:
        log.exception("read_upload failed")
        raise HTTPException(400, "Не удалось прочитать файл.")
    text = "\n".join(parts)
    if len(text.strip()) < 30:
        raise HTTPException(400, "В файле не найден текст (возможно, это скан). Загрузите PDF с выделяемым текстом или сфотографируйте страницы и отправьте в чат.")
    return text, meta


def read_upload(file: UploadFile) -> str:
    return read_upload_ex(file)[0]


@app.post("/upload")
def upload_file(file: UploadFile = File(...), user: dict = Depends(current_user)):
    text, meta = read_upload_ex(file)
    doc_id = secrets.token_hex(8)
    sb.table("documents").insert({
        "id": doc_id, "user_id": user["id"], "filename": (file.filename or "")[:200],
        "text": text[:MAX_STORE_CHARS],
    }).execute()
    warn = []
    if meta.get("empty_pages"):
        warn.append(f"на {meta['empty_pages']} из {meta['pages']} страниц нет текста (похоже на сканы) — ИИ их не увидит; сфотографируйте их и отправьте в чат")
    if len(text) > MAX_STORE_CHARS:
        warn.append(f"файл очень большой — сохранены первые {MAX_STORE_CHARS} символов")
    elif len(text) > MAX_DOC_CHARS:
        warn.append("материал большой: ИИ увидит выдержки из всех его частей")
    return {"doc_id": doc_id, "message": "Файл успешно прочитан!", "chars": len(text), "pages": meta.get("pages", 0),
            "warning": "; ".join(warn)}


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
    # ИИ-планер: 1 токен за запрос, промпт собирает фронтенд, ответ — JSON {"tasks": [...]}
    if action == "planner":
        prompt = custom_prompt.strip()[:1500]
        if not prompt:
            raise HTTPException(400, "Введите запрос для планера.")
        result, left = paid(user, lambda: ask_ai(prompt, json_mode=True))
        return {"result": result, "tokens_left": left}

    if action not in PROMPTS:
        raise HTTPException(400, "Неизвестный режим.")
    context = build_context(doc_id, title, user)
    prompt = f"{PROMPTS[action]}\n{lang_line(lang)}\n{MATERIAL_RULES}\nМатериал:\n{context}"
    result, left = paid(user, lambda: ask_ai(prompt, max_tokens=6000, temperature=0.3))
    hid = save_generation(user["id"], action, title if title != "Lesson" else "", result, doc_id)
    return {"result": result, "tokens_left": left, "history_id": hid}


OCR_MARK = "[Текст, распознанный с фото]"
PHOTO_DEFAULT = {"ru": "Опиши и разбери это фото.", "kz": "Бұл фотоны сипаттап, талдап бер.", "en": "Describe and analyze this photo."}


def _chat_prompt(message: str, doc_id: str, lang: str, history: str, user: dict) -> tuple[str, str]:
    message = message.strip()[:2000]
    if not message:
        raise HTTPException(400, "Пустое сообщение.")
    own = message.split(OCR_MARK)[0].strip() or message          # язык определяем по словам пользователя, а не по тексту, распознанному с фото
    ctx = get_doc_text(doc_id, user) or ""
    # память — только текущая беседа, которую присылает страница (старые чаты тему не навязывают)
    past: list = []
    try:
        hist = json.loads(history) if history else []
        past = [{"message": _str(h.get("m"), 800), "reply": _str(h.get("r"), 1500)} for h in hist[-6:] if isinstance(h, dict)]
    except Exception:
        past = []
    dialog = "\n".join(f"Студент: {m['message']}\nНаставник: {m['reply']}" for m in past)
    prompt = (
        "Ты — дружелюбный ИИ-наставник для студентов. Отвечай понятно и по делу. "
        "Формулы пиши в LaTeX: короткие — между $...$, отдельные (длинные, предел, дробь, сумма) — между $$...$$ на отдельной строке. Не используй \\[ \\] и \\( \\). " + request_lang_line(own, lang) + "\n"
        "ВАЖНО: не выбирай тему за студента и не додумывай её. Если сообщение — приветствие или в нём нет конкретного вопроса/темы, "
        "коротко поздоровайся и спроси, какую тему или задачу он хочет разобрать (1–2 предложения, без списка тем и без лекции).\n\n"
        + (MATERIAL_RULES + f"Материал студента:\n{fit_context(ctx)}\nОтвечай по материалу; если ответа в нём нет — скажи об этом прямо.\n\n" if ctx else "")
        + (f"Предыдущий диалог:\n{dialog}\n\n" if dialog else "")
        + f"Студент: {message}\nНаставник:"
    )
    return message, prompt


def _save_chat(user: dict, doc_id: str, message: str, reply: str) -> None:
    try:
        sb.table("chat_messages").insert({"user_id": user["id"], "doc_id": doc_id, "message": message, "reply": reply}).execute()
    except Exception:
        log.exception("chat save failed")


@app.post("/api/chat")
def chat(message: str = Form(...), doc_id: str = Form("default"), lang: str = Form("kz"), history: str = Form(""),
         user: dict = Depends(current_user)):
    message, prompt = _chat_prompt(message, doc_id, lang, history, user)
    reply, left = paid(user, lambda: ask_ai(prompt))
    _save_chat(user, doc_id, message, reply)
    return {"reply": reply, "tokens_left": left}


def _read_images(files) -> list[str]:
    """Проверяет и превращает загруженные фото в data-URL (в базе не сохраняются)."""
    out: list[str] = []
    for f in (files or [])[:4]:
        ct = (f.content_type or "").lower()
        if not ct.startswith("image/"):
            continue
        if ct not in ("image/jpeg", "image/png", "image/webp", "image/gif"):
            raise HTTPException(400, "Поддерживаются фото JPG, PNG, WEBP и GIF.")
        raw = f.file.read(6_000_001)
        if len(raw) > 6_000_000:
            raise HTTPException(413, "Фото слишком большое (максимум 6 МБ).")
        if raw:
            out.append(f"data:{ct};base64,{base64.b64encode(raw).decode()}")
    return out


def _sse(obj: dict) -> str:
    return "data: " + json.dumps(obj, ensure_ascii=False) + "\n\n"


@app.post("/api/chat/stream")
def chat_stream(message: str = Form(""), doc_id: str = Form("default"), lang: str = Form("kz"), history: str = Form(""),
                images: list[UploadFile] = File(default=[]), user: dict = Depends(current_user)):
    """Чат с потоковым выводом (SSE): текст приходит кусками, как в ChatGPT. Токен списывается заранее и возвращается при ошибке.
    Если приложены фото — отвечает модель с поддержкой изображений (VISION_*)."""
    imgs = _read_images(images)
    if imgs and not VISION_API_KEY:                     # до списания токена: страница сама распознает текст на фото (OCR)
        return JSONResponse({"error": "Распознавание фото на сервере не настроено.", "code": "no_vision"}, status_code=501)
    message = message.strip() or (PHOTO_DEFAULT.get(lang, PHOTO_DEFAULT["ru"]) if imgs else "")
    message, prompt = _chat_prompt(message, doc_id, lang, history, user)
    left = spend_token(user["id"])                      # 402 вернётся обычным JSON ещё до начала потока
    if imgs:
        api_url, api_key, model = f"{VISION_BASE_URL}/chat/completions", VISION_API_KEY, VISION_MODEL
        content = [{"type": "text", "text": prompt + "\n(К сообщению студента приложены фото — используй их.)"}] + \
                  [{"type": "image_url", "image_url": {"url": u}} for u in imgs]
        saved_msg = f"📷×{len(imgs)} {message}"
    else:
        api_url, api_key, model = f"{DEEPSEEK_BASE_URL}/chat/completions", DEEPSEEK_API_KEY, DEEPSEEK_MODEL
        content, saved_msg = prompt, message

    def gen():
        parts: list[str] = []
        failed = False
        try:
            with httpx.stream(
                "POST", api_url,
                headers={"Authorization": f"Bearer {api_key}"},
                json={"model": model, "messages": [{"role": "user", "content": content}],
                      "max_tokens": 4000, "temperature": 0.7, "stream": True},
                timeout=httpx.Timeout(90.0, connect=10.0),
            ) as r:
                if r.status_code != 200:
                    r.read()
                    log.error("deepseek stream status %s", r.status_code)
                    raise RuntimeError("bad status")
                for line in r.iter_lines():
                    if not line.startswith("data:"):
                        continue
                    data = line[5:].strip()
                    if data == "[DONE]":
                        break
                    try:
                        piece = json.loads(data)["choices"][0]["delta"].get("content")
                    except Exception:
                        continue
                    if piece:
                        parts.append(piece)
                        yield _sse({"d": piece})
        except GeneratorExit:                           # пользователь закрыл вкладку / остановил ответ
            failed = not parts
            raise
        except Exception:
            log.exception("chat stream failed")
            failed = not parts
            if failed:
                yield _sse({"error": "ИИ сейчас недоступен, попробуйте ещё раз. Токен возвращён."})
        finally:
            reply = "".join(parts).strip()
            if reply:
                _save_chat(user, doc_id, saved_msg, reply)
            else:
                refund_token(user["id"])
        if "".join(parts).strip():
            yield _sse({"done": True, "tokens_left": left})
        elif not failed:                                # поток закончился, но текста нет
            yield _sse({"error": "ИИ вернул пустой ответ. Токен возвращён, попробуйте ещё раз."})

    return StreamingResponse(gen(), media_type="text/event-stream",
                             headers={"Cache-Control": "no-cache, no-transform", "X-Accel-Buffering": "no", "Connection": "keep-alive"})


@app.post("/api/tickets")
def tickets(theme: str = Form(...), comment: str = Form(""), count: int = Form(3),
            doc_id: str = Form("default"), lang: str = Form("kz"), user: dict = Depends(current_user)):
    theme = theme.strip()[:200]
    count = max(1, min(count, 10))
    ctx = fit_context(get_doc_text(doc_id, user) or "")
    prompt = (
        f"Составь {count} экзаменационных билетов по теме «{theme}» в формате игры-викторины. "
        f"Дополнительные требования: {comment[:500] or 'нет'}. " + request_lang_line(theme + " " + comment, lang) + "\n"
        + (MATERIAL_RULES + QUESTION_RULES + f"Материал студента:\n{ctx}\n" if ctx else "")
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


def make_questions(prompt: str, n: int, material: str = "", lang: str = "ru"):
    def run():
        data = parse_json(ask_ai(prompt, json_mode=True, max_tokens=6000, temperature=0.4 if material else 0.7))
        if isinstance(data, dict):
            data = data.get("questions")
        if not isinstance(data, list) or len(data) < min(n, 4):
            raise HTTPException(502, "ИИ вернул неверный формат. Попробуйте ещё раз.")
        data = [x for x in data if isinstance(x, dict) and x.get("q") and isinstance(x.get("options"), list)][:n]
        return _verify_questions(data, material, lang) if material else data
    return run


@app.post("/api/quiz")
def quiz(doc_id: str = Form("default"), lang: str = Form("kz"), title: str = Form(""),
         user: dict = Depends(current_user)):
    context = build_context(doc_id, title, user)
    prompt = f"Составь 20 тестовых вопросов по материалу. {QUIZ_FMT}\n{lang_line(lang)}\n{MATERIAL_RULES}{QUESTION_RULES}\nМатериал:\n{context}"
    data, left = paid(user, make_questions(prompt, 20, context, lang))
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
    ctx = fit_context(get_doc_text(doc_id, user) or "")
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
    prompt = f"{task}\n{QUIZ_FMT}\n{lang_line(lang)}\n" + (MATERIAL_RULES + QUESTION_RULES + f"Материал:\n{ctx}\n" if ctx else "")
    data, left = paid(user, make_questions(prompt, n, ctx, lang))
    save_generation(user["id"], "tickets", f"{theme or 'Билет'} · №{num}",
                    json.dumps({"questions": data}, ensure_ascii=False), doc_id)
    return {"questions": data, "tokens_left": left}


@app.post("/api/tickets/custom")
def tickets_custom(file: UploadFile = File(...), user: dict = Depends(current_user)):
    """Свои вопросы из PDF/Word/Excel: ИИ вытаскивает вопросы, сервер случайно раскладывает их по 4 билетам."""
    text = read_upload(file)
    prompt = ("Из текста выдели все экзаменационные вопросы (не больше 80), без номеров, по одному на строку, "
              'смысл не меняй. Верни ТОЛЬКО JSON: {"questions":["..."]}.\n\n' + fit_context(text))

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
    prompt = ("Создай 12 карточек для заучивания ключевых понятий из материала: равномерно охвати ВЕСЬ материал от начала до конца, "
              "ничего важного не пропусти. Вопрос короткий (до 15 слов), ответ — 1–3 предложения, точные термины, числа и формулы из материала. " + MATERIAL_RULES + lang_line(lang) +
              ' Верни ТОЛЬКО JSON: {"cards":[{"q":"...","a":"..."}]}.\n\nМатериал:\n' + context)

    def run():
        d = parse_json(ask_ai(prompt, json_mode=True, max_tokens=4000, temperature=0.3))
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


DISCOURSE_RE = re.compile(
    r"\b(?:таким образом|важно отметить|стоит отметить|следует отметить|кроме того|более того|помимо этого|в заключение|в целом|"
    r"во-первых|во-вторых|в-третьих|тем не менее|в первую очередь|иными словами|прежде всего|"
    r"moreover|furthermore|additionally|in addition|in conclusion|overall|however|therefore|thus|it is important to note|"
    r"осылайша|сонымен қатар|қорытындылай келе|атап өткен жөн|алайда|жалпы алғанда|яғни)\b", re.I)


def style_signals(text: str) -> dict:
    """Объективные признаки «машинного» стиля (по описанию в статье Wikipedia «Artificial intelligence content detection»):
    повторы слов, повторяющиеся вводные слова/связки, шаблонные начала предложений, слишком ровная длина предложений."""
    words = _words(text)
    long_w = [w for w in words if len(w) >= 5]
    cnt: dict[str, int] = defaultdict(int)
    for w in long_w:
        cnt[w] += 1
    rep_words = sorted(((w, n) for w, n in cnt.items() if n >= 3), key=lambda x: -x[1])[:5]
    rep_ratio = round((1 - len(cnt) / len(long_w)) * 100) if long_w else 0
    marks: dict[str, int] = defaultdict(int)
    for mm in DISCOURSE_RE.findall(text):
        marks[mm.lower()] += 1
    sents = [s for s in re.split(r"(?<=[.!?…])\s+", text.strip()) if s.strip()]
    lens = [len(_words(s)) for s in sents] or [0]
    avg = sum(lens) / len(lens)
    cv = round((sum((x - avg) ** 2 for x in lens) / len(lens)) ** 0.5 / avg, 2) if avg else 0.0
    openers: dict[str, int] = defaultdict(int)
    for s in sents:
        w = _words(s)[:1]
        if w:
            openers[w[0]] += 1
    return {
        "repeat_words": [f"{w} ×{n}" for w, n in rep_words], "repeat_pct": rep_ratio,
        "markers": [f"{k} ×{n}" for k, n in sorted(marks.items(), key=lambda x: -x[1])], "markers_total": sum(marks.values()),
        "same_openers": [f"{k} ×{n}" for k, n in openers.items() if n >= 3], "length_variation": cv,
    }


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
        "Ищи признаки, которые перечисляет статья Wikipedia «Artificial intelligence content detection»: "
        "(1) повторы одних и тех же слов и оборотов; (2) необычно ровная, «слишком правильная» грамматика и однотипная длина предложений; "
        "(3) шаблонные переходы («таким образом», «кроме того», «важно отметить»); (4) повторяющиеся вводные слова и связки; "
        "(5) повторяющиеся риторические шаблоны (одинаковая схема абзацев, «не только… но и», тезис–пример–вывод в каждом абзаце); "
        "(6) галлюцинации — выдуманные факты, цифры, ссылки, имена. "
        "Опирайся и на объективные сигналы, посчитанные программой (см. ниже), но не придумывай признаков, которых нет в тексте. "
        "Помни: детекторы ошибаются, а текст человека, для которого язык неродной, часто ложно выглядит «машинным», "
        "поэтому не завышай ai_score без явных признаков.\n"
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
        "• Убери повторы слов (подбирай синонимы или местоимения), не начинай подряд несколько предложений одним и тем же словом, "
        "не ставь вводные слова и связки в каждом предложении, ломай одинаковую схему абзацев.\n"
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
        '"human_text":"...","human_score":0,"changes":[{"before":"...","after":"..."}]}\n\n'
        "ОБЪЕКТИВНЫЕ СИГНАЛЫ ИСХОДНОГО ТЕКСТА (посчитаны программой): " + json.dumps(style_signals(text), ensure_ascii=False) + "\n\nТЕКСТ:\n" + text
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
        best["signals"] = {"before": style_signals(text), "after": style_signals(str(best.get("human_text", "")))}
        return best

    d, left = paid(user, run)
    return {"result": d, "tokens_left": left}


# ==================== ПОСТУПЛЕНИЕ ПОСЛЕ 11 КЛАССА ====================
# Данные прошлого года (конкурс грантов 2025): минимальный балл получивших грант в общем конкурсе.
# Официальные итоги публикует НЦТ — grant.testcenter.kz; сводка по вузам и ГОП (источник этих чисел) — univision.kz.
# kind: grant — проходной на грант; threshold — порог вуза (реальный проходной обычно выше).
# Обновляйте раз в год после публикации итогов конкурса грантов. Сверено вручную: КБТУ, КазНУ, ЕНУ, КазНПУ, Астана МУ.
KB_YEAR = 2025
SOURCES = {"official": "https://grant.testcenter.kz", "official_name": "grant.testcenter.kz (НЦТ)", "summary": "https://univision.kz", "summary_name": "univision.kz"}
# Итоги основного ЕНТ-2025 (сообщение МНВО РК от 12.07.2025: более 210 тыс. абитуриентов, 381 тыс. попыток)
ENT_STATS = {"year": 2025, "takers": 210000, "attempts": 381000, "avg": 69, "max": 140, "passed_pct": 76, "passed": 161000, "high130": 598,
             "source_name": "МНВО РК / НЦТ", "source": "https://www.nur.kz/society/2264556-obyavleny-itogi-osnovnogo-ent-v-kazahstane/"}
KB = {
    "kbtu_it": ("КБТУ", "Информационные технологии", 110, "grant", "other", True),
    "enu_it": ("ЕНУ им. Л.Н. Гумилёва", "Информационные технологии", 110, "grant", "other", True),
    "kaznu_it": ("КазНУ им. аль-Фараби", "Информационные технологии", 90, "grant", "other", True),
    "satbayev_it": ("Satbayev University", "Информационные технологии (порог вуза 85; на общий конкурс 2025 — от 75)", 85, "threshold", "other", True),
    "kaznu_law": ("КазНУ им. аль-Фараби", "Право", 90, "grant", "law", True),
    "astana_med": ("Медицинский университет Астана", "Общая медицина", 128, "grant", "health", True),
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
    "kaznpu_math": ("КазНПУ им. Абая", "Подготовка учителей математики", 106, "grant", "pedagogy", True),
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


@app.get("/api/career/info")
def career_info():
    """Публичная статистика прошлого года (ЕНТ-2025) и источники данных — для стартового экрана раздела."""
    return {"year": KB_YEAR, "ent_stats": ENT_STATS, "sources": SOURCES}


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
        "year": KB_YEAR, "score": body.score, "tokens_left": left, "sources": SOURCES, "ent_stats": ENT_STATS,
    }


# ==================== ПРЕЗЕНТАЦИИ (как Gamma) ====================
SLIDE_FONTS = ["Inter", "Playfair Display", "Space Grotesk", "Montserrat", "Manrope", "Unbounded", "Lora", "Oswald", "Comfortaa", "Rubik"]
SLIDE_LAYOUTS = {"title", "section", "bullets", "two_col", "stats", "quote", "timeline", "compare", "closing"}
IMG_LAYS = {"title", "section", "bullets", "quote", "closing"}
IMG_HOSTS = ("https://image.pollinations.ai/prompt/", "https://loremflickr.com/")
SLIDE_DECOS = {"orbs", "grid", "waves", "lines", "diag", "blocks", "strips", "gradblock", "outline", "organic", "curve", "none"}
HEX_RE = re.compile(r"^#[0-9a-fA-F]{6}$")


def _T(name, desc, bg, bg2, tx, mut, ac, ac2, fh, fb, deco, ttl="grad"):
    return {"name": name, "desc": desc, "bg": bg, "bg2": bg2, "tx": tx, "mut": mut, "ac": ac, "ac2": ac2, "fh": fh, "fb": fb, "deco": deco, "ttl": ttl}


PRESET_STYLES = [
    # --- по референсам пользователя ---
    _T("Жёлтый контраст", "Солнечно-жёлтый фон, чёрные диагонали и крупные цифры — энергично, как у музыкальных сервисов", "#ffcc00", "#ffcc00", "#0a0a0a", "#4a4630", "#0a0a0a", "#0057d9", "Montserrat", "Montserrat", "diag", "solid"),
    _T("Электро-синий", "Белый и лавандовый с яркими синими блоками — смело и по-деловому, для бизнес-планов", "#ffffff", "#eceaff", "#0a0a23", "#55557a", "#0000ff", "#8c8cff", "Montserrat", "Inter", "blocks", "solid"),
    _T("Корпоративный бирюзовый", "Белый фон, косые голубые полосы — строгий стиль для отчётов и услуг", "#ffffff", "#eaf6fb", "#0b2a3a", "#4f6b7a", "#16a6d9", "#0b5d8a", "Montserrat", "Inter", "strips", "solid"),
    _T("Бирюзовый градиент", "Обложки с плавным градиентом бирюзовый → синий и чистые белые слайды", "#ffffff", "#f1f8f8", "#0f2a2a", "#52706f", "#0e9f8e", "#1b5fb5", "Montserrat", "Inter", "gradblock", "solid"),
    _T("Кофейный контур", "Тёплая бумага, терракота и мята, огромные буквы-контуры — дружелюбно и игриво", "#f3ece4", "#e7d6c4", "#3b2a20", "#7a6455", "#b9704a", "#3fae9f", "Comfortaa", "Manrope", "outline", "solid"),
    _T("Персиковый flat", "Мягкие органические формы: персик, апельсин и зелень — как рисованная иллюстрация", "#fbe3d0", "#f6cdb0", "#4a2418", "#8a5a45", "#e8663d", "#7a9a3a", "Rubik", "Manrope", "organic", "solid"),
    _T("Красно-чёрный модерн", "Белый, чёрный и красный с плавными изгибами — стильно и современно", "#ffffff", "#f3f3f3", "#111111", "#666666", "#d90429", "#111111", "Montserrat", "Inter", "curve", "solid"),
    # --- новые ---
    _T("Лавандовый туман", "Нежный сиреневый с розовым акцентом — мягко и стильно", "#f6f3ff", "#e9e2ff", "#241b4b", "#6b6490", "#7c5cff", "#ff7ab8", "Manrope", "Inter", "orbs", "solid"),
    _T("Мятная свежесть", "Светлая мята и тёплый жёлтый, органичные формы — легко и свежо", "#e9fbf3", "#d3f4e4", "#0f3d2e", "#4f7a69", "#1fae7a", "#e0a100", "Rubik", "Manrope", "organic", "solid"),
    _T("Золото и ночь", "Чёрный и золото, засечные шрифты — премиально и торжественно", "#0e0e10", "#1b1710", "#f7f1e3", "#b5a98a", "#d4af37", "#f5e6a8", "Playfair Display", "Lora", "lines", "solid"),
    _T("Океан", "Глубокая синева и голубые волны — спокойно и масштабно", "#06283d", "#0b4f6c", "#eaf6ff", "#9cc4d9", "#4cc9f0", "#4895ef", "Montserrat", "Inter", "waves", "solid"),
    _T("Сакура", "Розовая пудра и тёплый персик — нежно и по-доброму", "#fff0f3", "#ffd9e2", "#4a1a2c", "#8a5468", "#e8457c", "#f08a24", "Comfortaa", "Manrope", "organic", "solid"),
    _T("Минимал ч/б", "Чёрно-белая строгость и серые блоки — для чистых идей", "#ffffff", "#f2f2f2", "#0a0a0a", "#6b6b6b", "#0a0a0a", "#9a9a9a", "Space Grotesk", "Inter", "blocks", "solid"),
    _T("Космос", "Тёмный космос с фиолетовым и бирюзовым свечением — для науки и технологий", "#05010f", "#1a0b3d", "#f4eaff", "#a99bc9", "#b388ff", "#00e5ff", "Unbounded", "Manrope", "orbs", "solid"),
    _T("Спорт-лайм", "Чёрный с кислотным лаймом и оранжевым — динамика и скорость", "#111111", "#111111", "#ffffff", "#a8a8a8", "#b6ff00", "#ff5a1f", "Montserrat", "Inter", "diag", "solid"),
    _T("Университет", "Строгий синий и золото — для учебных и научных докладов", "#f7f9fc", "#e8eef7", "#14213d", "#5a6b88", "#1e3a8a", "#c9a227", "Playfair Display", "Inter", "strips", "solid"),
    # --- классика ---
    _T("Neon Night", "Тёмный фон, неоновый лайм и фиолетовый — дерзко и современно", "#0b0b12", "#1a1a2e", "#f5f5f0", "#9a9aa8", "#c8ff2e", "#8b5cf6", "Space Grotesk", "Inter", "orbs"),
    _T("Clean Light", "Светлый минимализм, синий акцент — чисто и серьёзно", "#ffffff", "#eef2f7", "#0f172a", "#64748b", "#2563eb", "#06b6d4", "Manrope", "Inter", "lines"),
    _T("Editorial", "Тёплая бумага и засечки — как страница журнала", "#faf5ee", "#f0e4d3", "#1c1917", "#78716c", "#b45309", "#7c2d12", "Playfair Display", "Lora", "none"),
    _T("Aurora", "Глубокий бирюзовый и северное сияние — спокойно и дорого", "#042f2e", "#134e4a", "#ecfeff", "#99c6c3", "#5eead4", "#a78bfa", "Montserrat", "Inter", "waves"),
    _T("Sunset Pop", "Сливовый закат с розовым и золотым — ярко и смело", "#1f1235", "#4a1942", "#fff7ed", "#c9b3c9", "#fb7185", "#fbbf24", "Unbounded", "Manrope", "orbs"),
    _T("Corporate Navy", "Тёмно-синий и голубой — для бизнеса и защиты проектов", "#0f172a", "#1e3a8a", "#f8fafc", "#a5b4cf", "#38bdf8", "#f59e0b", "Montserrat", "Inter", "grid"),
]


for _i, _t in enumerate(PRESET_STYLES):
    _t["g"] = 1 if _i < 7 else (2 if _i < 16 else 3)


def clean_theme(t, fb=None) -> dict:
    fb = fb or PRESET_STYLES[0]
    t = t if isinstance(t, dict) else {}
    out = {k: (t.get(k) if HEX_RE.match(str(t.get(k, ""))) else fb[k]) for k in ("bg", "bg2", "tx", "mut", "ac", "ac2")}
    out["fh"] = t.get("fh") if t.get("fh") in SLIDE_FONTS else fb["fh"]
    out["fb"] = t.get("fb") if t.get("fb") in SLIDE_FONTS else fb["fb"]
    out["deco"] = t.get("deco") if t.get("deco") in SLIDE_DECOS else fb["deco"]
    out["ttl"] = t.get("ttl") if t.get("ttl") in ("solid", "grad") else fb.get("ttl", "grad")
    out["name"] = str(t.get("name") or fb["name"])[:40]
    out["desc"] = str(t.get("desc") or "")[:160]
    return out


THEME_SCHEMA = ('{"name":"...","desc":"1 фраза о настроении","bg":"#hex","bg2":"#hex","tx":"#hex","mut":"#hex","ac":"#hex","ac2":"#hex",'
                '"fh":"шрифт заголовков","fb":"шрифт текста","deco":"orbs|grid|waves|lines|diag|blocks|strips|gradblock|outline|organic|curve|none","ttl":"solid|grad"}')
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
    return {"styles": styles, "presets": PRESET_STYLES, "ai": bool(styles)}


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
    t = str(v if v is not None else "").encode("utf-8", "ignore").decode("utf-8", "ignore")
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]", "", t).strip()[:n]


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
    out["vq"] = _str(sl.get("vq"), 220) if layout in IMG_LAYS else ""
    out["kw"] = _str(sl.get("kw"), 70) if layout in IMG_LAYS else ""
    # слайд без содержимого под свой layout → безопасный bullets
    need = {"stats": out["stats"], "timeline": out["steps"], "quote": out["quote"], "two_col": out["left"]["items"] or out["right"]["items"],
            "compare": out["left"]["items"] or out["right"]["items"]}
    if layout in need and not need[layout]:
        out["layout"] = "bullets"
    return out


def salvage_slides(text: str) -> list:
    """Достаёт целые объекты-слайды из «обрезанного» или битого JSON (ответ ИИ мог оборваться на длинной презентации)."""
    i = text.find('"slides"')
    i = text.find("[", i) if i >= 0 else -1
    if i < 0:
        return []
    out, depth, start, in_str, esc_ = [], 0, -1, False, False
    for k in range(i + 1, len(text)):
        ch = text[k]
        if in_str:
            if esc_:
                esc_ = False
            elif ch == "\\":
                esc_ = True
            elif ch == '"':
                in_str = False
            continue
        if ch == '"':
            in_str = True
        elif ch == "{":
            if depth == 0:
                start = k
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0 and start >= 0:
                try:
                    obj = json.loads(text[start:k + 1])
                    if isinstance(obj, dict):
                        out.append(obj)
                except json.JSONDecodeError:
                    pass
                start = -1
        elif ch == "]" and depth == 0:
            break
    return out


MODE_RULES = {
    "write": "РЕЖИМ ТЕКСТА: пользователь дал идею или тему — ты сам пишешь весь контент: точно, конкретно, без воды.",
    "shorten": ("РЕЖИМ ТЕКСТА: пользователь дал готовый текст — СОКРАТИ его до тезисов. Сохрани все ключевые факты, числа, имена и порядок "
                "мыслей, убери повторы и воду. Пункт — до 12 слов."),
    "verbatim": ("РЕЖИМ ТЕКСТА: пользователь дал готовый текст — ОСТАВЬ ЕГО КАК ЕСТЬ. Запрещено перефразировать, сокращать, "
                 "исправлять или добавлять слова. Раздели текст на слайды по смыслу и копируй предложения и абзацы ДОСЛОВНО в bullets/items; "
                 "ты придумываешь только заголовок слайда, подзаголовок и layout. Если текст длиннее слайда — переноси остаток на следующий слайд."),
}


def _generate_deck(body: "SlidesIn", user: dict) -> dict:
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
        "КАРТИНКИ: для слайдов title, section, bullets, quote и closing задай vq — описание иллюстрации НА АНГЛИЙСКОМ (8–18 слов: "
        "предмет, действие, обстановка, настроение; без текста, букв, цифр и логотипов) и kw — 2–3 английских слова через запятую для поиска "
        "стокового фото (конкретные существительные, например \"solar panels,energy\"). Для остальных layout vq и kw оставь пустыми.\n\n"
        f"{theme_rule}\n{request_lang_line(body.prompt, body.lang)}\n\n"
        'Верни ТОЛЬКО JSON: {"title":"название презентации","theme":{...},"slides":[{"layout":"bullets","title":"...","subtitle":"","icon":"🚀",'
        '"bullets":["..."],"left":{"h":"","items":[]},"right":{"h":"","items":[]},"stats":[{"v":"","l":""}],'
        '"steps":[{"h":"","t":""}],"quote":"","author":"","vq":"","kw":"","notes":"..."}]}'
    )

    def run():
        best_slides, best_raw, last_err = [], None, None
        for attempt in range(3):                      # до трёх попыток: ИИ иногда отдаёт битый или неполный JSON
            try:
                txt = ask_ai(prompt, json_mode=True, max_tokens=8000, temperature=0.8 if attempt == 0 else 0.6, timeout=150.0)
            except HTTPException as e:
                last_err = e
                time.sleep(1.5)
                continue
            raw = parse_json(txt)
            items = raw.get("slides") if isinstance(raw, dict) else None
            if not isinstance(items, list) or len(items) < 3:
                items = salvage_slides(txt) or items or []
            slides = [x for x in (clean_slide(sl) for sl in items) if x]
            if len(slides) > len(best_slides):
                best_slides, best_raw = slides, raw
            if len(slides) >= max(3, n - 1):
                break
        if len(best_slides) < min(3, n):
            raise last_err or HTTPException(502, "ИИ вернул неполную презентацию. Попробуйте ещё раз — токены вернутся.")
        raw = best_raw if isinstance(best_raw, dict) else {}
        theme = preset or clean_theme(raw.get("theme"))
        return {"title": _str(raw.get("title"), 120) or best_slides[0]["title"], "theme": theme, "slides": best_slides[:n]}

    deck, left, spent = paid_n(user, n, run)
    extra = spent - len(deck["slides"])      # ИИ дал меньше слайдов — возвращаем разницу
    for _ in range(max(0, extra)):
        refund_token(user["id"]); left += 1
    save_generation(user["id"], "slides", deck["title"], json.dumps(deck, ensure_ascii=False))
    return {"deck": deck, "tokens_left": left, "charged": len(deck["slides"])}


@app.post("/api/slides")
def slides_generate(body: SlidesIn, user: dict = Depends(current_user)):
    return _generate_deck(body, user)


# Генерация в фоне: запрос не висит минутами (прокси обрывают долгие запросы), страница опрашивает статус.
SLIDE_JOBS: dict[str, dict] = {}
_JOB_POOL = ThreadPoolExecutor(max_workers=6)


def _run_slide_job(jid: str, body: "SlidesIn", user: dict) -> None:
    job = SLIDE_JOBS[jid]
    try:
        job.update(status="done", result=_generate_deck(body, user))
    except HTTPException as e:
        job.update(status="error", error=str(e.detail))
    except Exception:
        log.exception("slide job failed")
        job.update(status="error", error="Не удалось создать презентацию. Токены возвращены, попробуйте ещё раз.")
    job["finished"] = time.time()


@app.post("/api/slides/start")
def slides_start(body: SlidesIn, user: dict = Depends(current_user)):
    now = time.time()
    for k in [k for k, j in SLIDE_JOBS.items() if now - j["created"] > 1800]:
        SLIDE_JOBS.pop(k, None)
    if any(j["user"] == user["id"] and j["status"] == "running" for j in SLIDE_JOBS.values()):
        raise HTTPException(409, "Предыдущая презентация ещё создаётся. Подождите минуту.")
    rate_limit(f"slidejob:{user['id']}", 8, 600)
    jid = secrets.token_urlsafe(12)
    SLIDE_JOBS[jid] = {"user": user["id"], "status": "running", "created": now}
    _JOB_POOL.submit(_run_slide_job, jid, body, user)
    return {"job_id": jid}


@app.get("/api/slides/job/{jid}")
def slides_job(jid: str, user: dict = Depends(current_user)):
    job = SLIDE_JOBS.get(jid)
    if not job or job["user"] != user["id"]:
        raise HTTPException(404, "Задача не найдена (сервер мог перезапуститься). Попробуйте создать презентацию заново.")
    if job["status"] == "done":
        return {"status": "done", **job["result"]}
    if job["status"] == "error":
        return {"status": "error", "error": job["error"]}
    return {"status": "running"}


def build_pptx(deck: dict, imgs: Optional[dict] = None) -> bytes:
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Inches, Pt

    th = clean_theme(deck.get("theme"))
    deco = th["deco"]
    rgb = lambda h: RGBColor.from_string(h[1:].upper())
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]
    res = {"diag": 2.9, "blocks": 1.4, "curve": 2.7}.get(deco, 0)      # место справа под декор
    imgs = imgs or {}
    AW = 11.6 - res

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

    def decorate(sl, big):
        R, O, P = MSO_SHAPE.RECTANGLE, MSO_SHAPE.OVAL, MSO_SHAPE.PARALLELOGRAM
        if deco == "diag":
            shape(sl, P, 10.6 if big else 11.6, -0.1, 3.4 if big else 2.4, 7.7, th["ac"])
            shape(sl, MSO_SHAPE.RIGHT_TRIANGLE, 12.2, 5.9, 1.2, 1.7, th["ac2"])
        elif deco == "blocks":
            shape(sl, R, 0, 0, 0.55, 7.5, th["ac2"])
            shape(sl, R, 10.6 if big else 12.2, 0, 2.73 if big else 1.13, 7.5 if big else 1.13, th["ac"])
        elif deco == "strips":
            shape(sl, P, -0.5, 0, 1.6, 7.5, th["ac"])
            shape(sl, MSO_SHAPE.RIGHT_TRIANGLE, 11.0, 5.0, 2.33, 2.5, th["ac2"])
        elif deco == "gradblock":
            shape(sl, R, 0, 0, 0.4, 7.5, th["ac"])
            if not big:
                shape(sl, O, 12.0, -0.8, 2.2, 2.2, th["ac2"])
        elif deco == "outline":
            shape(sl, O, 12.0, 0.6, 0.5, 0.5, th["ac2"])
            shape(sl, O, 12.7, 1.1, 0.25, 0.25, th["ac"])
        elif deco == "organic":
            shape(sl, O, 11.0, -1.2, 3.4, 2.7, th["ac"])
            shape(sl, O, -1.0, 6.2, 3.6, 2.6, th["ac2"])
        elif deco == "curve":
            shape(sl, O, 11.3, -0.9, 4.0, 9.3, th["ac"])
            shape(sl, O, 12.3, -0.7, 1.8, 1.8, th["ac2"])
        else:
            shape(sl, R, 0, 0, 0.22, 7.5, th["ac"])
            if big:
                shape(sl, O, 9.6, -1.4, 5.2, 5.2, th["bg2"])
                shape(sl, O, 10.9, 4.6, 3.4, 3.4, th["ac2"])

    for idx, d in enumerate(deck.get("slides") or []):
        sl = prs.slides.add_slide(blank)
        lay, title = d.get("layout", "bullets"), d.get("title", "")
        big = lay in ("title", "closing", "section")
        flip = deco == "gradblock" and big          # обложка с градиентным фоном
        bg, tx, mut = (th["ac"], "#ffffff", "#e8f4ff") if flip else (th["bg"], th["tx"], th["mut"])
        sl.background.fill.solid()
        sl.background.fill.fore_color.rgb = rgb(bg)
        decorate(sl, big)
        left = 1.0 if deco not in ("blocks", "strips") else 1.4
        pic = imgs.get(idx) if lay in IMG_LAYS else None
        pw = (4.3 if not res else 3.2) if pic else 0
        if pic:
            px = 13.333 - 0.7 - res - pw
            sl.shapes.add_picture(io.BytesIO(pic), Inches(px), Inches(1.6 if not big else 1.5), Inches(pw), Inches(pw))
        if big:
            tw = min(9.4, AW - 1.0) if res else 9.4
            if pic:
                tw = min(tw, 13.333 - 0.7 - res - pw - left - 0.4)
            box(sl, left, 2.0, 9.0, 0.9, d.get("icon", ""), 44, th["ac"] if not flip else "#ffffff")
            fs = 44 if len(title) <= 28 else 36
            if tw < 5.6:
                fs = min(fs, 30)
            lines_n = max(1, -(-int(len(title) * 1.2) // max(1, int(tw * 72 / (fs * 0.78)))))
            box(sl, left, 2.9, tw, 0.4 + lines_n * fs * 1.35 / 72, title, fs, tx, True, th["fh"])
            sub_y = min(6.0, 2.9 + lines_n * fs * 1.35 / 72 + 0.3)
            if d.get("subtitle"):
                box(sl, left, sub_y, min(9.0, tw), 1.0, d["subtitle"], 22, mut)
            if lay == "closing" and d.get("bullets"):
                box(sl, left, min(6.2, sub_y + 0.9), AW - 0.5, 1.0, ["→ " + b for b in d["bullets"][:3]], 16, th["ac"] if not flip else "#ffffff")
        else:
            x0 = 0.9 if deco not in ("blocks", "strips") else 1.3
            box(sl, x0, 0.55, AW - 0.2, 1.3, title, 34, tx, True, th["fh"], anchor=MSO_ANCHOR.MIDDLE)
            AWp = AW - (pw + 0.6 if pic else 0)
            shape(sl, MSO_SHAPE.RECTANGLE, x0 + 0.05, 1.9, 1.2, 0.07, th["ac"])
            if lay in ("two_col", "compare"):
                w = (AW - 0.5) / 2
                for j, side in enumerate(("left", "right")):
                    sd = d.get(side) or {}
                    x = x0 + j * (w + 0.4)
                    shape(sl, MSO_SHAPE.ROUNDED_RECTANGLE, x, 2.3, w, 4.5, th["bg2"])
                    box(sl, x + 0.3, 2.45, w - 0.6, 0.7, sd.get("h", ""), 22, th["ac"] if j == 0 else th["ac2"], True, th["fh"])
                    box(sl, x + 0.3, 3.2, w - 0.6, 3.4, ["• " + t for t in sd.get("items", [])], 17, tx)
            elif lay == "stats":
                st = d.get("stats") or []
                w = (AW - 0.2) / max(1, len(st))
                for j, it in enumerate(st):
                    x = x0 + j * w
                    shape(sl, MSO_SHAPE.ROUNDED_RECTANGLE, x, 2.5, w - 0.25, 3.6, th["bg2"])
                    box(sl, x + 0.1, 2.9, w - 0.45, 1.6, it["v"], 54, th["ac"] if j % 2 == 0 else th["ac2"], True, th["fh"], PP_ALIGN.CENTER)
                    box(sl, x + 0.2, 4.6, w - 0.65, 1.3, it["l"], 16, tx, align=PP_ALIGN.CENTER)
            elif lay == "quote":
                box(sl, x0 + 0.3, 2.3, AWp - 0.8, 3.2, "“" + d.get("quote", "") + "”", 30 if not pic else 24, tx, False, th["fh"], italic=True, anchor=MSO_ANCHOR.MIDDLE)
                box(sl, x0 + 0.3, 5.7, AWp - 0.8, 0.8, "— " + d.get("author", ""), 18, th["ac"])
            elif lay == "timeline":
                stp = d.get("steps") or []
                w = AW / max(1, len(stp))
                shape(sl, MSO_SHAPE.RECTANGLE, x0, 3.05, AW - 0.2, 0.05, th["mut"])
                for j, it in enumerate(stp):
                    x = x0 + j * w
                    shape(sl, MSO_SHAPE.OVAL, x, 2.75, 0.65, 0.65, th["ac"] if j % 2 == 0 else th["ac2"])
                    box(sl, x, 2.78, 0.65, 0.6, str(j + 1), 18, th["bg"], True, th["fh"], PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
                    box(sl, x, 3.65, w - 0.2, 0.8, it["h"], 20, tx, True, th["fh"])
                    box(sl, x, 4.45, w - 0.25, 2.2, it["t"], 14, mut)
            else:
                box(sl, x0 + 0.05, 2.35, AWp - 0.4, 4.6, ["▸  " + b for b in d.get("bullets", [])], 22 if not pic else 19, tx)
        if d.get("notes"):
            sl.notes_slide.notes_text_frame.text = d["notes"]
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def fetch_images(urls: dict) -> dict:
    """Скачивает картинки для PPTX параллельно. Только разрешённые хосты; ошибки пропускаем."""
    if not urls:
        return {}
    from concurrent.futures import ThreadPoolExecutor

    def one(item):
        idx, url = item
        if url.startswith("/api/img?"):                      # свой прокси: берём напрямую, без HTTP
            qs = parse_qs(urlparse(url).query)
            try:
                got = get_slide_image((qs.get("k") or ["ai"])[0], (qs.get("q") or [""])[0][:600], abs(int((qs.get("s") or ["1"])[0])) % 1_000_000)
            except Exception:
                got = None
            return idx, (got[0] if got else None)
        try:
            r = httpx.get(url, timeout=httpx.Timeout(28.0, connect=6.0), follow_redirects=True, headers={"User-Agent": "EnumaAIHub/1.0"})
            ok = r.status_code == 200 and r.headers.get("content-type", "").startswith("image/") and len(r.content) < 6_000_000
            return idx, (r.content if ok else None)
        except Exception:
            return idx, None

    with ThreadPoolExecutor(max_workers=6) as ex:
        return {i: b for i, b in ex.map(one, urls.items()) if b}


@app.post("/api/slides/export")
def slides_export(deck: dict = Body(...), user: dict = Depends(current_user)):
    slides, urls = [], {}
    for sl in (deck.get("slides") or [])[:14]:
        c = clean_slide(sl)
        if c:
            u = str(sl.get("img") or "")
            if (u.startswith(IMG_HOSTS) or u.startswith("/api/img?")) and len(u) < 1500 and c["layout"] in IMG_LAYS:
                urls[len(slides)] = u
            slides.append(c)
    if not slides:
        raise HTTPException(400, "Пустая презентация.")
    imgs = fetch_images(urls)
    try:
        data = build_pptx({"theme": deck.get("theme"), "slides": slides}, imgs)
    except ImportError:
        log.exception("python-pptx is not installed")
        raise HTTPException(500, "На сервере не установлен python-pptx. Добавьте его в requirements.txt и сделайте новый деплой.")
    except Exception as e:
        log.exception("pptx export failed")
        raise HTTPException(500, f"Не удалось собрать PPTX ({type(e).__name__}).")
    title = re.sub(r"[^\w\-]+", "_", str(deck.get("title") or "presentation")).strip("_")[:40] or "presentation"
    ascii_name = re.sub(r"[^A-Za-z0-9_\-]+", "_", title).strip("_") or "presentation"
    return Response(
        data, media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f"attachment; filename=\"{ascii_name}.pptx\"; filename*=UTF-8''{quote(title)}.pptx"})


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


# ==================== ПОКУПКА ТОКЕНОВ (админ начисляет после оплаты через Kaspi) ====================
TOKEN_PACKS = {100: 300, 300: 800, 500: 1200}      # токенов -> тенге


class CreditIn(BaseModel):
    user_id: str
    amount: int          # 100 / 300 / 500 или -100 / -300 / -500 (отмена ошибочного начисления)


@app.post("/api/admin/tokens")
def admin_tokens(body: CreditIn, admin: dict = Depends(admin_user)):
    pack = abs(body.amount)
    if pack not in TOKEN_PACKS:
        raise HTTPException(400, "Допустимые пакеты: 100, 300, 500 токенов.")
    sign = 1 if body.amount > 0 else -1
    exists = sb.table("profiles").select("id,tokens").eq("id", body.user_id).limit(1).execute().data
    if not exists:
        raise HTTPException(404, "Пользователь не найден.")
    try:
        sb.rpc("add_tokens", {"p_user": body.user_id, "p_amount": sign * pack}).execute()
    except Exception:
        log.exception("add_tokens rpc failed, fallback to update")
        sb.table("profiles").update({"tokens": max(0, exists[0]["tokens"] + sign * pack)}).eq("id", body.user_id).execute()
    try:
        sb.table("token_purchases").insert({"user_id": body.user_id, "admin_id": admin["id"], "tokens": sign * pack,
                                            "price": sign * TOKEN_PACKS[pack]}).execute()
    except Exception:
        log.exception("token_purchases insert failed")
    row = sb.table("profiles").select(",".join(ADMIN_FIELDS)).eq("id", body.user_id).limit(1).execute().data[0]
    return {"user": row, "price": sign * TOKEN_PACKS[pack]}


@app.get("/api/admin/stats")
def admin_stats(_: dict = Depends(admin_user)):
    since = (datetime.now(timezone.utc) - timedelta(days=30)).isoformat()
    try:
        rows = sb.table("token_purchases").select("tokens,price").gte("created_at", since).limit(5000).execute().data
    except Exception:
        return {"tokens": 0, "revenue": 0, "count": 0}
    return {"tokens": sum(r["tokens"] for r in rows), "revenue": sum(r["price"] for r in rows), "count": sum(1 for r in rows if r["tokens"] > 0)}


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

    # пожизненная статистика и дни входа хранятся отдельно от истории — «Очистить» их не затрагивает
    life: dict[str, int] = {}
    try:
        life = {r["kind"]: int(r["n"]) for r in (sb.table("usage_counters").select("kind,n").eq("user_id", uid).limit(200).execute().data or [])}
    except Exception:
        log.warning("usage_counters unavailable — выполните migration_v8.sql")
    stored_days: list = []
    try:
        stored_days = (sb.table("activity_days").select("day").eq("user_id", uid)
                       .gte("day", (local_today - timedelta(days=400)).isoformat()).limit(1000).execute().data or [])
    except Exception:
        pass

    xp_by_day: dict[date, int] = defaultdict(int)
    active_days: set[date] = set()
    for r in stored_days:
        try:
            active_days.add(date.fromisoformat(str(r["day"])[:10]))
        except ValueError:
            pass
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

    materials_now = len(gens)
    for kind, n in life.items():
        if kind.startswith("gen:") and n > 0:
            by_type[kind[4:]] = max(by_type.get(kind[4:], 0), n)
    return {
        "profile": {**public_profile(user), "telegram_id": user.get("telegram_id")},
        "counts": {
            "materials": max(life.get("gen", 0), materials_now), "chats": max(life.get("chat", 0), chats), "by_type": by_type,
            "tasks_total": len(tasks), "tasks_done": max(life.get("task_done", 0), sum(1 for t in tasks if t["done"])),
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
    start: Optional[str] = None
    minutes: Optional[int] = Field(default=None, ge=5, le=480)


class TasksIn(BaseModel):
    tasks: list[TaskIn] = Field(min_length=1, max_length=40)


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
            "date": r["due_date"], "done": r["done"], "start": r.get("start_time"), "minutes": r.get("minutes")}


@app.get("/api/tasks")
def tasks_list(user: dict = Depends(current_user)):
    rows = sb.table("tasks").select("*").eq("user_id", user["id"]).order("due_date").limit(500).execute().data
    return {"tasks": [task_out(r) for r in rows]}


@app.post("/api/tasks")
def tasks_create(body: TasksIn, user: dict = Depends(current_user)):
    base = [{"user_id": user["id"], "title": t.title.strip(), "description": t.desc.strip(), "due_date": safe_date(t.date)} for t in body.tasks]
    full = [{**b, "start_time": _hm_norm(t.start), "minutes": t.minutes if _hm_norm(t.start) else None} for b, t in zip(base, body.tasks)]
    try:
        rows = sb.table("tasks").insert(full).execute().data
    except Exception as e:                      # колонок start_time/minutes ещё нет (migration_v5.sql не выполнена) — сохраняем без времени
        if "start_time" not in str(e) and "minutes" not in str(e) and "PGRST204" not in str(e):
            raise
        rows = sb.table("tasks").insert(base).execute().data
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


# ==================== СОГЛАСИЕ, ЭКСПОРТ И УДАЛЕНИЕ ДАННЫХ ====================
CONSENT_VERSION = "2026-10-29"


@app.post("/api/onboarded")
def mark_onboarded(user: dict = Depends(current_user)):
    """Пользователь прошёл (или пропустил) обучающий тур — больше не показываем."""
    try:
        sb.table("profiles").update({"onboarded_at": datetime.now(timezone.utc).isoformat()}).eq("id", user["id"]).execute()
    except Exception:
        log.exception("onboarded update failed")
    return {"ok": True}


@app.post("/api/consent")
def give_consent(user: dict = Depends(current_user)):
    """Фиксирует согласие пользователя с Политикой конфиденциальности и Условиями (версия + время)."""
    now = datetime.now(timezone.utc).isoformat()
    sb.table("profiles").update({"consent_at": now, "consent_version": CONSENT_VERSION}).eq("id", user["id"]).execute()
    return {"ok": True, "consent_version": CONSENT_VERSION, "consent_at": now}


@app.get("/api/me/export")
def export_my_data(user: dict = Depends(current_user)):
    """Выгрузка всех данных пользователя в JSON (право на доступ к своим данным)."""
    uid = user["id"]

    def rows(table: str, cols: str = "*", order: str = "created_at"):
        try:
            return sb.table(table).select(cols).eq("user_id", uid).order(order, desc=True).limit(5000).execute().data or []
        except Exception:
            log.exception("export %s failed", table)
            return []

    mem = sb.table("group_members").select("group_id,role,joined_at").eq("user_id", uid).execute().data or []
    data = {
        "exported_at": datetime.now(timezone.utc).isoformat(),
        "profile": {k: v for k, v in user.items() if k not in ("id",)},
        "documents": rows("documents"),
        "generations": rows("generations"),
        "chat_messages": rows("chat_messages"),
        "tasks": rows("tasks", order="due_date"),
        "xp_events": rows("xp_events"),
        "groups": mem,
    }
    return Response(json.dumps(data, ensure_ascii=False, default=str, indent=1), media_type="application/json",
                    headers={"Content-Disposition": 'attachment; filename="enuma_my_data.json"'})


@app.delete("/api/me")
def delete_my_account(user: dict = Depends(current_user)):
    """Полное удаление аккаунта и связанных данных (право на удаление)."""
    uid = user["id"]
    # группы, где пользователь владелец: передаём владение другому участнику или удаляем пустую группу
    for g in sb.table("groups").select("id").eq("owner_id", uid).execute().data or []:
        others = (sb.table("group_members").select("user_id").eq("group_id", g["id"]).neq("user_id", uid)
                  .order("joined_at").limit(1).execute().data)
        if others:
            sb.table("groups").update({"owner_id": others[0]["user_id"]}).eq("id", g["id"]).execute()
            sb.table("group_members").update({"role": "owner"}).eq("group_id", g["id"]).eq("user_id", others[0]["user_id"]).execute()
        else:
            sb.table("groups").delete().eq("id", g["id"]).execute()
    if user.get("telegram_id"):
        sb.table("auth_codes").delete().eq("telegram_id", user["telegram_id"]).execute()
    sb.table("profiles").delete().eq("id", uid).execute()      # остальное удалится каскадом
    return {"ok": True}


# ==================== КАРТИНКИ ДЛЯ СЛАЙДОВ (серверный прокси) ====================
# Браузер на телефоне часто не может загрузить картинки напрямую с pollinations/loremflickr
# (лимиты по IP, долгая генерация, блокировки). Сервер сам ходит за картинкой, повторяет попытки,
# при неудаче берёт стоковое фото и кэширует результат.
POLLINATIONS_API_KEY = os.getenv("POLLINATIONS_API_KEY", "")
_IMG_CACHE: "OrderedDict[str, tuple[bytes, str]]" = OrderedDict()
_IMG_LOCKS: dict[str, threading.Lock] = defaultdict(threading.Lock)
_IMG_CACHE_MAX = 150
_AI_SEM = threading.BoundedSemaphore(3)       # не более 3 одновременных генераций
_IMG_FAIL: dict[str, float] = {}               # ключ → когда была неудача (на минуту не долбим источник)


def client_ip(request: Request) -> str:
    fwd = request.headers.get("x-forwarded-for", "")
    return (fwd.split(",")[0].strip() if fwd else (request.client.host if request.client else "?")) or "?"


def _http_image(url: str, headers: Optional[dict] = None, timeout: float = 60.0) -> Optional[tuple[bytes, str]]:
    try:
        r = httpx.get(url, timeout=httpx.Timeout(timeout, connect=8.0), follow_redirects=True,
                      headers={"User-Agent": "Mozilla/5.0 EnumaAIHub/1.0", **(headers or {})})
        ctype = r.headers.get("content-type", "").split(";")[0].strip().lower()
        if r.status_code == 200 and ctype.startswith("image/") and 1000 < len(r.content) < 6_000_000:
            return r.content, ctype
        log.warning("img %s -> %s %s", url[:80], r.status_code, ctype)
    except Exception as e:
        log.warning("img fetch failed: %s", type(e).__name__)
    return None


def _stock_tags(q: str) -> str:
    words = [w for w in re.findall(r"[A-Za-z]{4,}", q) if w.lower() not in {"cute", "style", "illustration", "with", "that", "from", "this", "letters", "watermark", "text", "high", "detail", "colors", "vibrant", "soft", "lighting"}]
    return ",".join(words[:3]) or "study"


def get_slide_image(kind: str, q: str, seed: int) -> Optional[tuple[bytes, str]]:
    key = f"{kind}|{seed}|{q}"
    hit = _IMG_CACHE.get(key)
    if hit:
        _IMG_CACHE.move_to_end(key)
        return hit
    if time.time() - _IMG_FAIL.get(key, 0) < 45:
        return None
    with _IMG_LOCKS[key]:
        hit = _IMG_CACHE.get(key)
        if hit:
            return hit
        res = None
        if kind == "ai":
            enc = quote(q[:600], safe="")
            if POLLINATIONS_API_KEY:
                url = f"https://gen.pollinations.ai/image/{enc}?width=768&height=768&nologo=true&seed={seed}&model=flux"
                hdr = {"Authorization": f"Bearer {POLLINATIONS_API_KEY}"}
            else:
                url = f"https://image.pollinations.ai/prompt/{enc}?width=768&height=768&nologo=true&seed={seed}&model=flux"
                hdr = {}
            for attempt, tmo in enumerate((45.0, 30.0)):
                with _AI_SEM:
                    res = _http_image(url, hdr, timeout=tmo)
                if res:
                    break
                time.sleep(1.0)
                url = url.replace(f"seed={seed}", f"seed={seed + attempt + 1}")
        if not res:
            tags = quote(_stock_tags(q) if kind == "ai" else q[:80], safe=",")
            res = _http_image(f"https://loremflickr.com/768/768/{tags}?lock={seed}", timeout=20.0)
        if not res:
            _IMG_FAIL[key] = time.time()
            if len(_IMG_FAIL) > 500:
                _IMG_FAIL.clear()
        if res:
            _IMG_CACHE[key] = res
            while len(_IMG_CACHE) > _IMG_CACHE_MAX:
                _IMG_CACHE.popitem(last=False)
        _IMG_LOCKS.pop(key, None)
        return res


@app.get("/api/img")
def slide_image(request: Request, k: str = "ai", q: str = "", s: int = 1):
    q = q.strip()[:600]
    if k not in ("ai", "stock") or len(q) < 2:
        raise HTTPException(400, "bad request")
    s = abs(int(s)) % 1_000_000
    cached = _IMG_CACHE.get(f"{k}|{s}|{q}")
    if not cached:
        rate_limit(f"img:{client_ip(request)}", 160, 600)
    got = cached or get_slide_image(k, q, s)
    if not got:
        raise HTTPException(502, "Картинка недоступна.")
    return Response(got[0], media_type=got[1], headers={"Cache-Control": "public, max-age=86400"})


# ==================== ГРУППЫ: приглашения по нику и распределение задач ====================
MAX_GROUPS_OWNED = 10
MAX_MEMBERS = 20
MAX_GROUP_TASKS = 200
NICK_RE = re.compile(r"^[A-Za-z0-9_]{3,64}$")


class GroupIn(BaseModel):
    name: str = Field(min_length=2, max_length=60)


class InviteIn(BaseModel):
    nick: str = Field(min_length=1, max_length=80)


class GTaskIn(BaseModel):
    title: str = Field(min_length=1, max_length=200)
    description: str = Field(default="", max_length=1000)
    assignee_id: Optional[str] = None
    due_date: Optional[str] = None


class GTaskPatch(BaseModel):
    title: Optional[str] = Field(default=None, min_length=1, max_length=200)
    description: Optional[str] = Field(default=None, max_length=1000)
    assignee_id: Optional[str] = None
    unassign: bool = False
    due_date: Optional[str] = None
    done: Optional[bool] = None


class SplitIn(BaseModel):
    goal: str = Field(min_length=5, max_length=2000)
    lang: str = "ru"


def _gdate(s: Optional[str]) -> Optional[str]:
    if not s:
        return None
    try:
        return date.fromisoformat(str(s)[:10]).isoformat()
    except ValueError:
        return None


def _who(p: dict) -> dict:
    return {"user_id": p["id"], "username": p.get("username"), "display_name": p.get("display_name") or p.get("username") or "—"}


def _profiles(ids) -> dict:
    ids = list({i for i in ids if i})
    if not ids:
        return {}
    rows = sb.table("profiles").select("id,username,display_name,telegram_id").in_("id", ids).execute().data or []
    return {r["id"]: r for r in rows}


_ROLE_CACHE: dict[tuple[int, str], tuple[str, float]] = {}


def _forget_role(gid: int, uid: Optional[str] = None) -> None:
    for k in [k for k in _ROLE_CACHE if k[0] == gid and (uid is None or k[1] == uid)]:
        _ROLE_CACHE.pop(k, None)


def _my_role(gid: int, uid: str) -> Optional[str]:
    """Роль участника в группе. Положительный ответ кэшируем на 15 секунд — чат и задачи не бьют в БД на каждый запрос."""
    hit = _ROLE_CACHE.get((gid, uid))
    if hit and hit[1] > time.time():
        return hit[0]
    r = sb.table("group_members").select("role").eq("group_id", gid).eq("user_id", uid).limit(1).execute().data
    if not r:
        _ROLE_CACHE.pop((gid, uid), None)
        return None
    if len(_ROLE_CACHE) > 5000:
        _ROLE_CACHE.clear()
    _ROLE_CACHE[(gid, uid)] = (r[0]["role"], time.time() + 15)
    return r[0]["role"]


def _need_member(gid: int, user: dict) -> str:
    role = _my_role(gid, user["id"])
    if not role:
        raise HTTPException(404, "Группа не найдена.")
    return role


def _need_owner(gid: int, user: dict) -> None:
    if _need_member(gid, user) != "owner":
        raise HTTPException(403, "Это действие доступно только владельцу группы.")


def _count_members(gid: int) -> int:
    return len(sb.table("group_members").select("user_id").eq("group_id", gid).execute().data or [])


@app.get("/api/groups")
def groups_list(user: dict = Depends(current_user)):
    uid = user["id"]
    mem = sb.table("group_members").select("group_id,role").eq("user_id", uid).execute().data or []
    gids = [m["group_id"] for m in mem]
    roles = {m["group_id"]: m["role"] for m in mem}
    groups, counts, tasks = [], defaultdict(int), defaultdict(lambda: [0, 0])
    if gids:
        groups = sb.table("groups").select("*").in_("id", gids).order("created_at", desc=True).execute().data or []
        for r in sb.table("group_members").select("group_id").in_("group_id", gids).execute().data or []:
            counts[r["group_id"]] += 1
        for r in sb.table("group_tasks").select("group_id,done").in_("group_id", gids).limit(5000).execute().data or []:
            tasks[r["group_id"]][1 if r["done"] else 0] += 1
    inv = sb.table("group_invites").select("id,group_id,invited_by,created_at").eq("user_id", uid).eq("status", "pending").execute().data or []
    igroups = {}
    if inv:
        for g in sb.table("groups").select("id,name").in_("id", [i["group_id"] for i in inv]).execute().data or []:
            igroups[g["id"]] = g["name"]
    inviters = _profiles([i["invited_by"] for i in inv])
    return {
        "groups": [{"id": g["id"], "name": g["name"], "role": roles.get(g["id"]), "members": counts[g["id"]],
                    "open": tasks[g["id"]][0], "done": tasks[g["id"]][1],
                    "goal": g.get("goal") or "", "deadline": g.get("deadline")} for g in groups],
        "invites": [{"id": i["id"], "group_id": i["group_id"], "group_name": igroups.get(i["group_id"], "—"),
                     "from": _who(inviters[i["invited_by"]]) if i["invited_by"] in inviters else None,
                     "created_at": i["created_at"]} for i in inv if i["group_id"] in igroups],
    }


@app.post("/api/groups")
def groups_create(body: GroupIn, user: dict = Depends(current_user)):
    rate_limit(f"gcreate:{user['id']}", 10, 600)
    if len(sb.table("groups").select("id").eq("owner_id", user["id"]).execute().data or []) >= MAX_GROUPS_OWNED:
        raise HTTPException(400, f"Можно создать не больше {MAX_GROUPS_OWNED} групп.")
    g = sb.table("groups").insert({"name": body.name.strip(), "owner_id": user["id"]}).execute().data[0]
    sb.table("group_members").insert({"group_id": g["id"], "user_id": user["id"], "role": "owner"}).execute()
    return {"id": g["id"], "name": g["name"]}


@app.get("/api/groups/{gid}")
def group_detail(gid: int, user: dict = Depends(current_user)):
    """Вся группа за один заход: запросы к БД идут параллельно, профили участников подтягиваются тем же запросом."""
    role = _need_member(gid, user)
    f_group = _TEACH_POOL.submit(lambda: sb.table("groups").select("*").eq("id", gid).limit(1).execute().data)
    f_tasks = _TEACH_POOL.submit(lambda: sb.table("group_tasks").select("*").eq("group_id", gid).order("created_at").limit(MAX_GROUP_TASKS + 5).execute().data or [])
    f_inv = _TEACH_POOL.submit(lambda: (sb.table("group_invites").select("id,user_id,created_at").eq("group_id", gid).eq("status", "pending").execute().data or []) if role == "owner" else [])
    try:
        members = sb.table("group_members").select("user_id,role,joined_at,profiles(username,display_name,score)").eq("group_id", gid).order("joined_at").execute().data or []
        people = [{"user_id": m["user_id"], "role": m["role"], "username": (m.get("profiles") or {}).get("username"),
                   "display_name": (m.get("profiles") or {}).get("display_name") or (m.get("profiles") or {}).get("username") or "—",
                   "score": (m.get("profiles") or {}).get("score") or 0} for m in members]
    except Exception:                                               # запасной путь без встраивания профилей
        members = sb.table("group_members").select("user_id,role,joined_at").eq("group_id", gid).order("joined_at").execute().data or []
        prof = _profiles([m["user_id"] for m in members])
        people = [{**_who(prof[m["user_id"]]), "role": m["role"], "score": prof[m["user_id"]].get("score") or 0} for m in members if m["user_id"] in prof]
    g, tasks, invites = f_group.result(), f_tasks.result(), f_inv.result()
    if not g:
        raise HTTPException(404, "Группа не найдена.")
    iprof = _profiles([i["user_id"] for i in invites]) if invites else {}
    return {
        "group": {"id": gid, "name": g[0]["name"], "owner_id": g[0]["owner_id"], "goal": g[0].get("goal") or "", "deadline": g[0].get("deadline")},
        "me": {"user_id": user["id"], "role": role},
        "members": people,
        "tasks": [{"id": t["id"], "title": t["title"], "description": t.get("description", ""), "assignee_id": t.get("assignee_id"),
                   "due_date": t.get("due_date"), "done": bool(t.get("done"))} for t in tasks],
        "invites": [{"id": i["id"], "user": _who(iprof[i["user_id"]])} for i in invites if i["user_id"] in iprof],
    }


class GroupPatch(BaseModel):
    name: Optional[str] = Field(default=None, min_length=2, max_length=60)
    goal: Optional[str] = Field(default=None, max_length=300)
    deadline: Optional[str] = None
    clear_deadline: bool = False


@app.patch("/api/groups/{gid}")
def group_patch(gid: int, body: GroupPatch, user: dict = Depends(current_user)):
    _need_owner(gid, user)
    upd: dict = {}
    if body.name is not None:
        upd["name"] = body.name.strip()
    if body.goal is not None:
        upd["goal"] = body.goal.strip()
    if body.clear_deadline:
        upd["deadline"] = None
    elif body.deadline:
        upd["deadline"] = _gdate(body.deadline)
    if upd:
        sb.table("groups").update(upd).eq("id", gid).execute()
    return {"ok": True}


class MsgIn(BaseModel):
    text: str = Field(min_length=1, max_length=1000)


def _msg_out(m: dict) -> dict:
    p = m.get("profiles") or {}
    return {"id": m["id"], "user_id": m["user_id"], "text": m["text"], "at": m["created_at"],
            "name": p.get("display_name") or p.get("username") or "—", "username": p.get("username")}


@app.get("/api/groups/{gid}/messages")
def group_messages(gid: int, after: int = 0, user: dict = Depends(current_user)):
    """Чат группы. after=0 — последние 50 сообщений; after=N — только новые (страница опрашивает раз в несколько секунд)."""
    _need_member(gid, user)
    q = sb.table("group_messages").select("id,user_id,text,created_at,profiles(username,display_name)").eq("group_id", gid)
    if after > 0:
        rows = q.gt("id", after).order("id").limit(100).execute().data or []
    else:
        rows = list(reversed(q.order("id", desc=True).limit(50).execute().data or []))
    return {"messages": [_msg_out(m) for m in rows]}


@app.post("/api/groups/{gid}/messages")
def group_message_send(gid: int, body: MsgIn, user: dict = Depends(current_user)):
    _need_member(gid, user)
    rate_limit(f"gmsg:{user['id']}", 40, 60)
    text = body.text.strip()
    if not text:
        raise HTTPException(400, "Пустое сообщение.")
    m = sb.table("group_messages").insert({"group_id": gid, "user_id": user["id"], "text": text}).execute().data[0]
    return {"message": {"id": m["id"], "user_id": user["id"], "text": text, "at": m["created_at"],
                        "name": user.get("display_name") or user.get("username") or "—", "username": user.get("username")}}


@app.delete("/api/groups/{gid}")
def group_delete(gid: int, user: dict = Depends(current_user)):
    _need_owner(gid, user)
    sb.table("groups").delete().eq("id", gid).execute()
    _forget_role(gid)
    return {"ok": True}


def _find_invitee(raw: str) -> tuple[Optional[dict], str]:
    """Ищет пользователя по @нику Telegram, по ID из кабинета (public_id) или по Telegram ID."""
    raw = (raw or "").strip()
    cols = "id,username,display_name,telegram_id"
    if re.fullmatch(r"#?\d{1,15}", raw):
        num = int(raw.lstrip("#"))
        for col in ("public_id", "telegram_id"):
            try:
                rows = sb.table("profiles").select(cols).eq(col, num).limit(1).execute().data
            except Exception:            # колонки public_id нет (миграция v4 не выполнена) — пробуем следующий способ
                rows = []
            if rows:
                return rows[0], raw
        return None, raw
    nick = re.sub(r"^@+", "", raw)
    if not NICK_RE.match(nick):
        raise HTTPException(400, "Введите @ник Telegram (латиница, цифры, _) или числовой ID из кабинета.")
    rows = sb.table("profiles").select(cols).ilike("username", nick.replace("_", "\\_")).limit(5).execute().data or []
    found = [p for p in rows if (p.get("username") or "").lower() == nick.lower()]
    return (found[0] if found else None), "@" + nick


@app.post("/api/groups/{gid}/invite")
def group_invite(gid: int, body: InviteIn, user: dict = Depends(current_user)):
    _need_owner(gid, user)
    rate_limit(f"ginvite:{user['id']}", 30, 600)
    target, shown = _find_invitee(body.nick)
    if not target:
        raise HTTPException(404, f"Пользователь «{shown}» не найден. Он должен хотя бы раз войти в EnumaAI Hub. Введите @ник или ID из его кабинета.")
    if target["id"] == user["id"]:
        raise HTTPException(400, "Нельзя пригласить самого себя.")
    if _my_role(gid, target["id"]):
        raise HTTPException(400, "Этот пользователь уже в группе.")
    if _count_members(gid) >= MAX_MEMBERS:
        raise HTTPException(400, f"В группе не может быть больше {MAX_MEMBERS} участников.")
    if sb.table("group_invites").select("id").eq("group_id", gid).eq("user_id", target["id"]).eq("status", "pending").limit(1).execute().data:
        raise HTTPException(400, "Приглашение уже отправлено.")
    try:
        sb.table("group_invites").insert({"group_id": gid, "user_id": target["id"], "invited_by": user["id"]}).execute()
    except Exception as e:
        if "23505" in str(e) or "duplicate" in str(e).lower():
            raise HTTPException(400, "Приглашение уже отправлено.")
        raise
    gname = sb.table("groups").select("name").eq("id", gid).limit(1).execute().data[0]["name"]
    me_name = user.get("display_name") or user.get("username") or "Пользователь"
    notified = bool(target.get("telegram_id")) and tg_send(
        target["telegram_id"], f"👥 {me_name} приглашает вас в группу «{gname}» в EnumaAI Hub.\nОткройте сайт → «Группы», чтобы принять приглашение.")
    return {"ok": True, "user": _who(target), "notified": notified}


@app.delete("/api/groups/{gid}/invites/{iid}")
def group_invite_cancel(gid: int, iid: int, user: dict = Depends(current_user)):
    _need_owner(gid, user)
    sb.table("group_invites").delete().eq("id", iid).eq("group_id", gid).eq("status", "pending").execute()
    return {"ok": True}


def _my_invite(iid: int, user: dict) -> dict:
    r = sb.table("group_invites").select("*").eq("id", iid).eq("user_id", user["id"]).eq("status", "pending").limit(1).execute().data
    if not r:
        raise HTTPException(404, "Приглашение не найдено или уже обработано.")
    return r[0]


@app.post("/api/invites/{iid}/accept")
def invite_accept(iid: int, user: dict = Depends(current_user)):
    inv = _my_invite(iid, user)
    if _count_members(inv["group_id"]) >= MAX_MEMBERS:
        raise HTTPException(400, "В группе уже нет мест.")
    claimed = sb.table("group_invites").update({"status": "accepted"}).eq("id", iid).eq("status", "pending").execute().data
    if not claimed:
        raise HTTPException(400, "Приглашение уже обработано.")
    sb.table("group_members").upsert({"group_id": inv["group_id"], "user_id": user["id"], "role": "member"}, on_conflict="group_id,user_id").execute()
    _forget_role(inv["group_id"], user["id"])
    return {"ok": True, "group_id": inv["group_id"]}


@app.post("/api/invites/{iid}/decline")
def invite_decline(iid: int, user: dict = Depends(current_user)):
    _my_invite(iid, user)
    sb.table("group_invites").update({"status": "declined"}).eq("id", iid).eq("status", "pending").execute()
    return {"ok": True}


@app.delete("/api/groups/{gid}/members/{uid}")
def group_remove_member(gid: int, uid: str, user: dict = Depends(current_user)):
    role = _need_member(gid, user)
    if uid != user["id"] and role != "owner":
        raise HTTPException(403, "Исключать участников может только владелец.")
    if _my_role(gid, uid) == "owner":
        raise HTTPException(400, "Владелец не может выйти из группы — удалите группу целиком.")
    sb.table("group_members").delete().eq("group_id", gid).eq("user_id", uid).execute()
    _forget_role(gid, uid)
    sb.table("group_tasks").update({"assignee_id": None}).eq("group_id", gid).eq("assignee_id", uid).execute()
    return {"ok": True}


def _check_assignee(gid: int, assignee_id: Optional[str]) -> None:
    if assignee_id and not _my_role(gid, assignee_id):
        raise HTTPException(400, "Исполнитель должен быть участником группы.")


def _notify_assignee(assignee_id: Optional[str], me: dict, title: str, gname: str) -> None:
    if not assignee_id or assignee_id == me["id"]:
        return
    p = _profiles([assignee_id]).get(assignee_id)
    if p and p.get("telegram_id"):
        tg_send(p["telegram_id"], f"📌 {me.get('display_name') or me.get('username') or 'Участник'} назначил(а) вам задачу в группе «{gname}»:\n{title[:200]}")


@app.post("/api/groups/{gid}/tasks")
def gtask_create(gid: int, body: GTaskIn, user: dict = Depends(current_user)):
    _need_member(gid, user)
    _check_assignee(gid, body.assignee_id)
    if len(sb.table("group_tasks").select("id").eq("group_id", gid).execute().data or []) >= MAX_GROUP_TASKS:
        raise HTTPException(400, "В группе слишком много задач. Удалите выполненные.")
    t = sb.table("group_tasks").insert({
        "group_id": gid, "title": body.title.strip(), "description": body.description.strip(),
        "assignee_id": body.assignee_id or None, "created_by": user["id"], "due_date": _gdate(body.due_date),
    }).execute().data[0]
    gname = sb.table("groups").select("name").eq("id", gid).limit(1).execute().data[0]["name"]
    _notify_assignee(body.assignee_id, user, body.title, gname)
    return {"id": t["id"]}


@app.patch("/api/groups/{gid}/tasks/{tid}")
def gtask_patch(gid: int, tid: int, body: GTaskPatch, user: dict = Depends(current_user)):
    _need_member(gid, user)
    cur = sb.table("group_tasks").select("*").eq("id", tid).eq("group_id", gid).limit(1).execute().data
    if not cur:
        raise HTTPException(404, "Задача не найдена.")
    upd: dict = {}
    if body.title is not None:
        upd["title"] = body.title.strip()
    if body.description is not None:
        upd["description"] = body.description.strip()
    if body.done is not None:
        upd["done"] = body.done
    if body.due_date is not None:
        upd["due_date"] = _gdate(body.due_date)
    if body.unassign:
        upd["assignee_id"] = None
    elif body.assignee_id:
        _check_assignee(gid, body.assignee_id)
        upd["assignee_id"] = body.assignee_id
    if upd:
        sb.table("group_tasks").update(upd).eq("id", tid).execute()
        if body.assignee_id and body.assignee_id != cur[0].get("assignee_id"):
            gname = sb.table("groups").select("name").eq("id", gid).limit(1).execute().data[0]["name"]
            _notify_assignee(body.assignee_id, user, upd.get("title", cur[0]["title"]), gname)
    return {"ok": True}


@app.delete("/api/groups/{gid}/tasks/{tid}")
def gtask_delete(gid: int, tid: int, user: dict = Depends(current_user)):
    role = _need_member(gid, user)
    cur = sb.table("group_tasks").select("created_by").eq("id", tid).eq("group_id", gid).limit(1).execute().data
    if not cur:
        return {"ok": True}
    if role != "owner" and cur[0].get("created_by") != user["id"]:
        raise HTTPException(403, "Удалить задачу может её автор или владелец группы.")
    sb.table("group_tasks").delete().eq("id", tid).execute()
    return {"ok": True}


@app.post("/api/groups/{gid}/split")
def group_split(gid: int, body: SplitIn, user: dict = Depends(current_user)):
    """ИИ делит общую цель на задачи и раздаёт участникам. Стоит 1 токен."""
    _need_member(gid, user)
    members = sb.table("group_members").select("user_id").eq("group_id", gid).execute().data or []
    prof = _profiles([m["user_id"] for m in members])
    people = [p for p in prof.values()]
    if not people:
        raise HTTPException(400, "В группе нет участников.")
    names = {p["id"]: (("@" + p["username"]) if p.get("username") else (p.get("display_name") or "участник")) for p in people}
    roster = "\n".join(f"- {names[p['id']]}" for p in people)
    n = min(12, max(len(people), 3) * 2)
    prompt = (
        "Ты — тимлид учебной группы. Раздели общую цель на конкретные задачи и справедливо распредели между участниками.\n"
        f"ЦЕЛЬ ГРУППЫ:\n<<<\n{body.goal.strip()}\n>>>\n\nУЧАСТНИКИ (используй ники ТОЧНО как написано):\n{roster}\n\n"
        f"Сделай от {len(people)} до {n} задач. Каждому участнику — хотя бы одна задача, нагрузка примерно равная. "
        "Заголовок — короткое действие (до 10 слов), description — 1–2 предложения, что именно сделать. " + request_lang_line(body.goal, body.lang) + "\n"
        'Верни ТОЛЬКО JSON: {"tasks":[{"title":"...","description":"...","assignee":"@ник"}]}'
    )
    data, left = paid(user, lambda: parse_json(ask_ai(prompt, json_mode=True, max_tokens=2500, temperature=0.6)))
    items = (data or {}).get("tasks") if isinstance(data, dict) else None
    if not items:
        refund_token(user["id"])
        raise HTTPException(502, "ИИ вернул неполный ответ. Токен возвращён, попробуйте ещё раз.")
    by_nick = {names[p["id"]].lower(): p["id"] for p in people}
    rows, rr = [], 0
    for it in items[:12]:
        if not isinstance(it, dict) or not _str(it.get("title"), 200):
            continue
        aid = by_nick.get(str(it.get("assignee", "")).strip().lower())
        if not aid:
            aid, rr = people[rr % len(people)]["id"], rr + 1
        rows.append({"group_id": gid, "title": _str(it.get("title"), 200), "description": _str(it.get("description"), 1000),
                     "assignee_id": aid, "created_by": user["id"]})
    if not rows:
        refund_token(user["id"])
        raise HTTPException(502, "ИИ не смог составить задачи. Токен возвращён.")
    sb.table("group_tasks").insert(rows).execute()
    return {"added": len(rows), "tokens_left": left}


# ==================== РАСПИСАНИЕ, ПЛАНЕР С УЧЁТОМ ВРЕМЕНИ, СВОДКА В TELEGRAM ====================
DAY_START_MIN, DAY_END_MIN = 8 * 60, 22 * 60            # учебный день: 08:00–22:00
WD_SHORT = {"ru": ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"], "kz": ["Дс", "Сс", "Ср", "Бс", "Жм", "Сб", "Жс"],
            "en": ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]}
MONTHS_G = {"ru": ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря"],
            "kz": ["қаңтар", "ақпан", "наурыз", "сәуір", "мамыр", "маусым", "шілде", "тамыз", "қыркүйек", "қазан", "қараша", "желтоқсан"],
            "en": ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]}
WD_WORDS = {  # для разбора названий дней из расписания
    1: ("пн", "пон", "mon", "дс", "дүй"), 2: ("вт", "втор", "tue", "сс", "сей"), 3: ("ср", "сред", "wed", "ср", "сәр"),
    4: ("чт", "четв", "thu", "бс", "бей"), 5: ("пт", "пят", "fri", "жм", "жұм"), 6: ("сб", "суб", "sat", "сн", "сен"), 7: ("вс", "воскр", "sun", "жс", "жек")}


def _lang(v: str) -> str:
    return v if v in ("ru", "kz", "en") else "ru"


def _hm_norm(s) -> Optional[str]:
    m = re.fullmatch(r"\s*(\d{1,2})[:.](\d{2})\s*", str(s or ""))
    if not m:
        return None
    h, mi = int(m.group(1)), int(m.group(2))
    return f"{h:02d}:{mi:02d}" if h < 24 and mi < 60 else None


def _mins(hm: str) -> int:
    h, m = hm.split(":")
    return int(h) * 60 + int(m)


def _hm(minutes: int) -> str:
    minutes = max(0, min(24 * 60 - 1, int(minutes)))
    return f"{minutes // 60:02d}:{minutes % 60:02d}"


def user_now(user: dict) -> datetime:
    tz = user.get("tz_offset")
    tz = 5 if tz is None else max(-12, min(14, int(tz)))
    return datetime.now(timezone.utc) + timedelta(hours=tz)


def _weekday_of(v) -> Optional[int]:
    if isinstance(v, (int, float)) or (isinstance(v, str) and v.strip().isdigit()):
        n = int(v)
        return n if 1 <= n <= 7 else None
    s = str(v or "").strip().lower()
    for n, words in WD_WORDS.items():
        if any(s.startswith(w) for w in words):
            return n
    return None


def _clean_sched(items) -> list[dict]:
    out, seen = [], set()
    for it in items or []:
        if not isinstance(it, dict):
            continue
        wd = _weekday_of(it.get("weekday") or it.get("day"))
        st = _hm_norm(it.get("start") or it.get("start_time"))
        title = _str(it.get("title") or it.get("subject"), 80)
        if not (wd and st and title):
            continue
        en = _hm_norm(it.get("end") or it.get("end_time")) or ""
        if en and _mins(en) <= _mins(st):
            en = ""
        key = (wd, st, title.lower())
        if key in seen:
            continue
        seen.add(key)
        out.append({"weekday": wd, "start": st, "end": en, "title": title, "room": _str(it.get("room"), 30)})
    out.sort(key=lambda x: (x["weekday"], x["start"]))
    return out[:100]


def _sched_out(r: dict) -> dict:
    return {"weekday": r["weekday"], "start": r["start_time"], "end": r.get("end_time") or "", "title": r["title"], "room": r.get("room") or ""}


def _load_sched(uid: str) -> list[dict]:
    try:
        rows = sb.table("schedule_items").select("*").eq("user_id", uid).order("weekday").order("start_time").limit(200).execute().data or []
    except Exception:
        log.warning("schedule table unavailable")
        return []
    return [_sched_out(r) for r in rows]


def _busy(sched: list[dict], tasks: list[dict], d: date, extra: Optional[list] = None) -> list[tuple[int, int]]:
    """Занятые интервалы дня в минутах: уроки (+10 мин на дорогу) и уже запланированные задачи."""
    out = []
    for it in sched:
        if it["weekday"] == d.isoweekday():
            a = _mins(it["start"])
            b = _mins(it["end"]) if it.get("end") else a + 45
            out.append((max(0, a - 10), b + 10))
    for t in tasks:
        if t.get("due_date") == d.isoformat() and t.get("start_time") and not t.get("done"):
            a = _mins(t["start_time"])
            out.append((a, a + int(t.get("minutes") or 45) + 10))
    out += extra or []
    return sorted(out)


def _free(busy: list, lo: int = DAY_START_MIN, hi: int = DAY_END_MIN, min_len: int = 30) -> list[tuple[int, int]]:
    cur, res = lo, []
    for a, b in busy:
        if a > cur and min(a, hi) - cur >= min_len:
            res.append((cur, min(a, hi)))
        cur = max(cur, b)
    if hi - cur >= min_len:
        res.append((cur, hi))
    return res


def _place(busy: list, start: Optional[int], minutes: int, lo: int) -> Optional[int]:
    """Ставит задачу в свободное время: сначала туда, куда просит ИИ, иначе в первое подходящее окно."""
    def ok(s: int) -> bool:
        return s >= lo and s + minutes <= DAY_END_MIN and all(s + minutes <= a or s >= b for a, b in busy)
    if start is not None and ok(start):
        return start
    for c in sorted({lo, *[b for _, b in busy]}):
        c = (c + 4) // 5 * 5
        if ok(c):
            return c
    return None


class PlanIn(BaseModel):
    query: str = Field(min_length=2, max_length=600)
    lang: str = "ru"


@app.post("/api/planner/ask")
def planner_ask(body: PlanIn, user: dict = Depends(current_user)):
    """ИИ-планер: понимает запрос, видит расписание и занятые окна, ставит задачи в свободное время. 1 токен."""
    rate_limit(f"plan:{user['id']}", 20, 600)
    lang = _lang(body.lang)
    now = user_now(user)
    today = now.date()
    sched = _load_sched(user["id"])
    try:
        tasks = sb.table("tasks").select("title,due_date,done,start_time,minutes").eq("user_id", user["id"]) \
            .gte("due_date", today.isoformat()).order("due_date").limit(120).execute().data or []
    except Exception:
        tasks = sb.table("tasks").select("title,due_date,done").eq("user_id", user["id"]).gte("due_date", today.isoformat()).limit(120).execute().data or []
    lines = []
    for i in range(14):
        d = today + timedelta(days=i)
        lo = DAY_START_MIN if i else min(DAY_END_MIN, max(DAY_START_MIN, (now.hour * 60 + now.minute + 20 + 4) // 5 * 5))
        busy = _busy(sched, tasks, d)
        free = ", ".join(f"{_hm(a)}–{_hm(b)}" for a, b in _free(busy, lo)) or "нет свободного времени"
        lessons = "; ".join(f"{it['start']}–{it['end'] or '…'} {it['title']}" for it in sched if it["weekday"] == d.isoweekday()) or "занятий нет"
        lines.append(f"{d.isoformat()} ({WD_SHORT['ru'][d.weekday()]}): занятия: {lessons}. СВОБОДНО: {free}")
    existing = "\n".join(f"- {t['due_date']} {t.get('start_time') or ''} {t['title']}" for t in tasks if not t.get("done"))[:1500] or "ничего"
    prompt = (
        f"Ты — личный планировщик студента. Сегодня {today.isoformat()} ({WD_SHORT['ru'][today.weekday()]}), сейчас {now.strftime('%H:%M')}.\n"
        f"ЗАПРОС СТУДЕНТА (выполни ровно его): «{body.query.strip()}»\n\n"
        "ПРАВИЛА:\n"
        "1. Следуй запросу буквально. Если указаны число дней, предмет, объём, дедлайн, длительность или время — соблюдай их. "
        "Одна просьба — одна задача; «план на N дней» — распредели по дням начиная с сегодня (если сегодня уже поздно — с завтра). "
        "Не добавляй общих советов и лишних задач.\n"
        "2. Ставь задачи ТОЛЬКО в свободные окна из списка ниже. start — HH:MM внутри окна, minutes — от 20 до 120 (кратно 5). "
        "Не больше 4 часов учёбы в день, между задачами — перерыв от 10 минут. Не нарушай дедлайны из запроса.\n"
        "3. title — конкретное действие до 70 символов («Решить 10 задач по пределам», а не «Учёба»). desc — одно короткое предложение: что именно сделать.\n"
        "4. Максимум 20 задач. Не дублируй уже запланированное.\n\n"
        f"ЗАНЯТОСТЬ И СВОБОДНОЕ ВРЕМЯ НА 14 ДНЕЙ (расписание повторяется каждую неделю):\n" + "\n".join(lines) + "\n\n"
        f"УЖЕ ЗАПЛАНИРОВАНО:\n{existing}\n\n"
        'Верни ТОЛЬКО JSON: {"message":"одно короткое предложение студенту — что именно ты запланировал","tasks":[{"title":"...","date":"YYYY-MM-DD","start":"HH:MM","minutes":45,"desc":"..."}]}. '
        + request_lang_line(body.query, lang)
    )

    def run():
        data = parse_json(ask_ai(prompt, json_mode=True, max_tokens=2600, temperature=0.4, timeout=75.0))
        items = data.get("tasks") if isinstance(data, dict) else None
        if not isinstance(items, list) or not items:
            raise HTTPException(502, "ИИ не смог составить задачи по запросу. Токен возвращён — переформулируйте запрос.")
        return data

    data, left = paid(user, run)
    placed, extra = [], defaultdict(list)
    for it in data["tasks"][:20]:
        if not isinstance(it, dict) or not _str(it.get("title"), 80):
            continue
        try:
            d = date.fromisoformat(str(it.get("date"))[:10])
        except ValueError:
            d = today
        d = min(max(d, today), today + timedelta(days=90))
        minutes = max(20, min(180, int(it.get("minutes") or 45) // 5 * 5)) if str(it.get("minutes") or "45").isdigit() or isinstance(it.get("minutes"), (int, float)) else 45
        lo = DAY_START_MIN if d > today else min(DAY_END_MIN, max(DAY_START_MIN, (now.hour * 60 + now.minute + 20 + 4) // 5 * 5))
        st = _hm_norm(it.get("start"))
        busy = _busy(sched, tasks, d, extra[d.isoformat()])
        s = _place(busy, _mins(st) if st else None, minutes, lo)
        if s is not None:
            extra[d.isoformat()].append((s, s + minutes + 10))
        placed.append({"title": _str(it.get("title"), 80), "date": d.isoformat(), "start": _hm(s) if s is not None else None,
                       "minutes": minutes if s is not None else None, "desc": _str(it.get("desc"), 200)})
    if not placed:
        refund_token(user["id"])
        raise HTTPException(502, "ИИ не смог составить задачи по запросу. Токен возвращён.")
    placed.sort(key=lambda x: (x["date"], x["start"] or "99:99"))
    return {"message": _str(data.get("message"), 240), "tasks": placed, "tokens_left": left}


# ---------- расписание ----------
class SchedItem(BaseModel):
    weekday: int = Field(ge=1, le=7)
    start: str
    end: str = ""
    title: str = Field(min_length=1, max_length=80)
    room: str = Field(default="", max_length=30)


class SchedIn(BaseModel):
    items: list[SchedItem] = Field(max_length=100)


class NotifyIn(BaseModel):
    enabled: bool = True
    hour: int = Field(default=7, ge=0, le=23)
    minute: int = Field(default=30, ge=0, le=59)
    tz: int = Field(default=5, ge=-12, le=14)
    lang: str = "ru"


def _settings_out(u: dict) -> dict:
    return {"enabled": u.get("notify_enabled", True) is not False, "hour": u.get("notify_hour", 7) if u.get("notify_hour") is not None else 7,
            "minute": u.get("notify_minute", 30) if u.get("notify_minute") is not None else 30,
            "tz": u.get("tz_offset", 5) if u.get("tz_offset") is not None else 5, "has_telegram": bool(u.get("telegram_id"))}


@app.get("/api/schedule")
def schedule_get(user: dict = Depends(current_user)):
    return {"items": _load_sched(user["id"]), "settings": _settings_out(user)}


@app.put("/api/schedule")
def schedule_put(body: SchedIn, user: dict = Depends(current_user)):
    items = _clean_sched([i.model_dump() for i in body.items])
    sb.table("schedule_items").delete().eq("user_id", user["id"]).execute()
    if items:
        sb.table("schedule_items").insert([{"user_id": user["id"], "weekday": i["weekday"], "start_time": i["start"], "end_time": i["end"],
                                            "title": i["title"], "room": i["room"]} for i in items]).execute()
    return {"items": items}


@app.put("/api/settings/notify")
def notify_put(body: NotifyIn, user: dict = Depends(current_user)):
    sb.table("profiles").update({"notify_enabled": body.enabled, "notify_hour": body.hour, "notify_minute": body.minute,
                                 "tz_offset": body.tz, "lang": _lang(body.lang)}).eq("id", user["id"]).execute()
    return {"ok": True}


def _vision_complete(prompt: str, imgs: list[str]) -> str:
    content = [{"type": "text", "text": prompt}] + [{"type": "image_url", "image_url": {"url": u}} for u in imgs]
    try:
        r = httpx.post(f"{VISION_BASE_URL}/chat/completions", headers={"Authorization": f"Bearer {VISION_API_KEY}"},
                       json={"model": VISION_MODEL, "messages": [{"role": "user", "content": content}], "max_tokens": 4000, "temperature": 0.1},
                       timeout=httpx.Timeout(90.0, connect=10.0))
        r.raise_for_status()
        return r.json()["choices"][0]["message"]["content"]
    except Exception:
        log.exception("vision failed")
        raise HTTPException(502, "Не удалось прочитать фото. Токен возвращён — попробуйте более чёткое фото.")


@app.post("/api/schedule/parse")
def schedule_parse(text: str = Form(""), lang: str = Form("ru"), file: Optional[UploadFile] = File(None),
                   images: list[UploadFile] = File(default=[]), user: dict = Depends(current_user)):
    """ИИ превращает расписание (фото, PDF, Word, Excel, текст) в список занятий недели. 1 токен. Сохраняет клиент после проверки."""
    rate_limit(f"schparse:{user['id']}", 8, 600)
    src = text.strip()
    if file is not None and file.filename:
        src += "\n" + read_upload(file)
    imgs = _read_images(images)
    if imgs and not VISION_API_KEY:
        return JSONResponse({"error": "Распознавание фото на сервере не настроено.", "code": "no_vision"}, status_code=501)
    if not src.strip() and not imgs:
        raise HTTPException(400, "Загрузите файл или вставьте текст расписания.")
    prompt = (
        "Ты разбираешь расписание занятий студента или школьника. Извлеки ВСЕ занятия недели.\n"
        'Верни ТОЛЬКО JSON: {"items":[{"weekday":1,"start":"08:30","end":"09:15","title":"Математика","room":"305"}]}\n'
        "weekday: 1=понедельник … 7=воскресенье. Время в 24-часовом формате HH:MM. Если время окончания не указано — оставь end пустым. "
        "title — короткое название предмета, room — кабинет/аудитория или пустая строка. Если есть чётная/нечётная неделя — бери первую. "
        "Ничего не выдумывай: чего нет в расписании — пропускай. Если указаны номера уроков без времени — используй стандартные звонки (1-й урок 08:30, по 45 минут + перемена 10 минут).\n"
        + lang_line(_lang(lang)) + "\n\nРАСПИСАНИЕ:\n" + (src[:12000] or "(см. фото)")
    )

    def run():
        raw = _vision_complete(prompt, imgs) if imgs else ask_ai(prompt, json_mode=True, max_tokens=4000, temperature=0.1, timeout=90.0)
        data = parse_json(raw)
        items = _clean_sched(data.get("items") if isinstance(data, dict) else data)
        if not items:
            raise HTTPException(502, "Не удалось найти занятия в расписании. Токен возвращён — проверьте файл или вставьте текст.")
        return items

    items, left = paid(user, run)
    return {"items": items, "tokens_left": left}


# ---------- ежедневная сводка в Telegram ----------
DIGEST_T = {
    "ru": {"hi": "☀️ Доброе утро, {name}! Сегодня {date}.", "lessons": "📚 Уроки сегодня:", "tasks": "✅ Задачи на сегодня:", "over": "⏳ Просрочено задач: {n} — загляните в планер.",
           "window": "💡 Лучшее окно для подготовки: {a}–{b}.", "bye": "Удачного дня! Отключить эти сообщения можно в планере на сайте.", "mins": "мин"},
    "kz": {"hi": "☀️ Қайырлы таң, {name}! Бүгін {date}.", "lessons": "📚 Бүгінгі сабақтар:", "tasks": "✅ Бүгінгі тапсырмалар:", "over": "⏳ Мерзімі өткен тапсырмалар: {n} — жоспарлаушыны ашыңыз.",
           "window": "💡 Дайындалуға ең ыңғайлы уақыт: {a}–{b}.", "bye": "Сәтті күн! Бұл хабарламаларды сайттағы жоспарлаушыдан өшіруге болады.", "mins": "мин"},
    "en": {"hi": "☀️ Good morning, {name}! Today is {date}.", "lessons": "📚 Today’s classes:", "tasks": "✅ Today’s tasks:", "over": "⏳ Overdue tasks: {n} — open the planner.",
           "window": "💡 Best window to study: {a}–{b}.", "bye": "Have a great day! You can turn these messages off in the planner on the site.", "mins": "min"},
}


def build_digest(u: dict, sched: list[dict], tasks: list[dict], today: date, now_min: int) -> Optional[str]:
    lang = _lang(u.get("lang") or "ru")
    T = DIGEST_T[lang]
    lessons = [it for it in sched if it["weekday"] == today.isoweekday()]
    todays = [t for t in tasks if t.get("due_date") == today.isoformat() and not t.get("done")]
    overdue = [t for t in tasks if (t.get("due_date") or "9") < today.isoformat() and not t.get("done")]
    if not lessons and not todays:
        return None
    nice = f"{WD_SHORT[lang][today.weekday()]}, {today.day} {MONTHS_G[lang][today.month - 1]}"
    name = (u.get("display_name") or u.get("username") or "").split(" ")[0] or "друг"
    parts = [T["hi"].format(name=name, date=nice)]
    if lessons:
        parts.append(T["lessons"] + "\n" + "\n".join(
            f"• {it['start']}{'–' + it['end'] if it.get('end') else ''} — {it['title']}" + (f" ({it['room']})" if it.get("room") else "") for it in lessons))
    if todays:
        todays.sort(key=lambda t: t.get("start_time") or "99:99")
        parts.append(T["tasks"] + "\n" + "\n".join(
            "• " + (f"{t['start_time']} · " if t.get("start_time") else "") + t["title"] + (f" ({t['minutes']} {T['mins']})" if t.get("minutes") else "") for t in todays))
    if overdue:
        parts.append(T["over"].format(n=len(overdue)))
    free = [w for w in _free(_busy(sched, tasks, today), max(DAY_START_MIN, now_min)) if w[1] - w[0] >= 60]
    if free:
        a, b = max(free, key=lambda w: w[1] - w[0])
        parts.append(T["window"].format(a=_hm(a), b=_hm(b)))
    parts.append(T["bye"])
    return "\n\n".join(parts)[:3900]


def _tasks_for_digest(uid: str) -> list[dict]:
    try:
        return sb.table("tasks").select("title,due_date,done,start_time,minutes").eq("user_id", uid).eq("done", False).limit(300).execute().data or []
    except Exception:
        return sb.table("tasks").select("title,due_date,done").eq("user_id", uid).eq("done", False).limit(300).execute().data or []


def _send_digest(u: dict, force: bool = False) -> bool:
    now = user_now(u)
    text = build_digest(u, _load_sched(u["id"]), _tasks_for_digest(u["id"]), now.date(), now.hour * 60 + now.minute)
    return bool(text) and bool(u.get("telegram_id")) and tg_send(u["telegram_id"], text)


@app.post("/api/digest/test")
def digest_test(user: dict = Depends(current_user)):
    """Отправить сводку на сегодня прямо сейчас (проверка настроек)."""
    if not user.get("telegram_id"):
        raise HTTPException(400, "Telegram не привязан.")
    rate_limit(f"digest:{user['id']}", 5, 600)
    now = user_now(user)
    text = build_digest(user, _load_sched(user["id"]), _tasks_for_digest(user["id"]), now.date(), now.hour * 60 + now.minute)
    if not text:
        return {"sent": False, "reason": "empty"}
    if not tg_send(user["telegram_id"], text):
        raise HTTPException(502, "Не удалось отправить сообщение в Telegram. Откройте бота и нажмите /start.")
    return {"sent": True}


def _digest_tick() -> None:
    rows = (sb.table("profiles").select("id,telegram_id,display_name,username,tz_offset,notify_hour,notify_minute,last_digest_on,lang")
            .eq("notify_enabled", True).not_.is_("telegram_id", "null").limit(5000).execute().data or [])
    for u in rows:
        try:
            now = user_now(u)
            due = (u.get("notify_hour") if u.get("notify_hour") is not None else 7) * 60 + (u.get("notify_minute") if u.get("notify_minute") is not None else 30)
            passed = now.hour * 60 + now.minute - due
            ld = now.date().isoformat()
            if passed < 0 or u.get("last_digest_on") == ld:
                continue
            q = sb.table("profiles").update({"last_digest_on": ld}).eq("id", u["id"])           # «захватываем» отправку: защита от дублей при нескольких воркерах
            q = q.is_("last_digest_on", "null") if u.get("last_digest_on") is None else q.eq("last_digest_on", u["last_digest_on"])
            if not q.execute().data:
                continue
            if passed <= 180:                                                                      # если сервер был выключен — устаревшую утреннюю сводку не шлём
                _send_digest(u)
        except Exception:
            log.exception("digest for %s failed", u.get("id"))


def _digest_loop() -> None:
    time.sleep(15)
    while True:
        try:
            _digest_tick()
            time.sleep(60)
        except Exception as e:
            log.warning("digest tick failed (%s) — проверьте, что выполнена migration_v5.sql", type(e).__name__)
            time.sleep(300)


@app.on_event("startup")
def _start_background() -> None:
    if os.getenv("DISABLE_DIGEST") == "1":
        return
    threading.Thread(target=_digest_loop, daemon=True, name="digest").start()


# ==================== ИИ-УЧИТЕЛЬ: объяснение → видео → 10 вопросов (+10 за 5 токенов) ====================
YOUTUBE_API_KEY = os.getenv("YOUTUBE_API_KEY", "")
_VIDEO_CACHE: "OrderedDict[str, Optional[dict]]" = OrderedDict()
_TEACH_POOL = ThreadPoolExecutor(max_workers=8)
MORE_QUESTIONS_COST = 5


def _iso_seconds(d: str) -> int:
    m = re.fullmatch(r"PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?", d or "")
    return (int(m.group(1) or 0) * 3600 + int(m.group(2) or 0) * 60 + int(m.group(3) or 0)) if m else 0


def find_video(query: str, lang: str) -> Optional[dict]:
    """Лучшее обучающее видео по запросу (нужен YOUTUBE_API_KEY): встраиваемое, 3–25 минут, выше — с большим числом просмотров и лайков."""
    if not YOUTUBE_API_KEY or not query:
        return None
    key = f"{lang}|{query.lower()}"
    if key in _VIDEO_CACHE:
        return _VIDEO_CACHE[key]
    best = None
    try:
        rl = {"ru": "ru", "kz": "kk", "en": "en"}[_lang(lang)]
        r = httpx.get("https://www.googleapis.com/youtube/v3/search", timeout=10.0, params={
            "part": "snippet", "q": query, "type": "video", "maxResults": 10, "videoEmbeddable": "true", "videoSyndicated": "true",
            "safeSearch": "strict", "relevanceLanguage": rl, "key": YOUTUBE_API_KEY})
        ids = [i["id"]["videoId"] for i in r.json().get("items", []) if i.get("id", {}).get("videoId")]
        if ids:
            r2 = httpx.get("https://www.googleapis.com/youtube/v3/videos", timeout=10.0, params={
                "part": "statistics,contentDetails,snippet", "id": ",".join(ids), "key": YOUTUBE_API_KEY})
            top = -1e9
            for v in r2.json().get("items", []):
                st, secs = v.get("statistics", {}), _iso_seconds(v.get("contentDetails", {}).get("duration", ""))
                views, likes = int(st.get("viewCount", 0) or 0), int(st.get("likeCount", 0) or 0)
                score = math.log10(views + 10) + 40 * min(likes / max(views, 1), 0.08) - (3 if secs < 180 or secs > 1500 else 0)
                if score > top:
                    top, best = score, {"id": v["id"], "title": _str(v["snippet"].get("title"), 140), "channel": _str(v["snippet"].get("channelTitle"), 80),
                                        "minutes": max(1, round(secs / 60))}
    except Exception:
        log.warning("youtube search failed")
        return None                       # не кэшируем неудачу
    _VIDEO_CACHE[key] = best
    while len(_VIDEO_CACHE) > 200:
        _VIDEO_CACHE.popitem(last=False)
    return best


def _teacher_explain(topic: str, comment: str, ctx: str, lang: str) -> dict:
    prompt = (
        f"Ты — лучший преподаватель: объясняешь так, что тема «щёлкает» с первого раза и читать интересно. Тема: «{topic}». "
        f"Пожелания студента: {comment or 'нет'}. " + request_lang_line(topic + " " + comment, lang) + "\n"
        + (MATERIAL_RULES + f"Материал студента:\n{ctx}\n" if ctx else "")
        + "Напиши объяснение в Markdown, чтобы его было легко и приятно читать. Строго такая структура (заголовки с эмодзи):\n"
          "### 🎯 Суть в одной фразе — 1–2 предложения.\n"
          "### 🧩 Простыми словами — яркая аналогия из жизни, которая сразу даёт понимание. Начни с интересного факта или вопроса, который цепляет.\n"
          "### 📚 Разбор по шагам — 3–6 пунктов нумерованным списком: идея + короткое пояснение. Термины выделяй **жирным**.\n"
          "### 🔢 Пример — один конкретный пример с разбором по шагам (для нематематических тем — реальный кейс, ситуация или кусочек кода).\n"
          "### ⚠️ Частые ошибки — 2–3 пункта, где обычно ошибаются.\n"
          "### ✅ Запомнить — 3–5 коротких тезисов.\n"
          "Короткие абзацы, живой дружелюбный тон, 250–450 слов, без лишних вступлений. Формулы: короткие в $...$, отдельные в $$...$$ на своей строке.\n"
          "Также придумай video_query — точный запрос для YouTube, чтобы найти лучшее обучающее видео именно по этой теме (на языке ответа, добавь слово «объяснение» или «explained», без кавычек).\n"
          'Верни ТОЛЬКО JSON: {"explanation":"...Markdown...","video_query":"..."}'
    )
    for _ in range(2):
        data = parse_json(ask_ai(prompt, json_mode=True, max_tokens=3500, temperature=0.6, timeout=100.0))
        if isinstance(data, dict) and len(str(data.get("explanation") or "")) > 150:
            return {"explanation": str(data["explanation"]).strip(), "video_query": _str(data.get("video_query"), 120)}
    raise HTTPException(502, "ИИ вернул неверный формат. Попробуйте ещё раз — токены вернутся.")


def _clean_questions(items) -> list[dict]:
    out, seen = [], set()
    for it in items or []:
        if not isinstance(it, dict):
            continue
        q = _str(it.get("q") or it.get("question"), 400)
        opts_raw = it.get("options") or it.get("choices")
        if not q or not isinstance(opts_raw, list) or q.lower() in seen:
            continue
        opts = [re.sub(r"^\s*[A-Da-d][\).:]\s+", "", _str(o, 220)) for o in opts_raw][:4]
        if len(opts) != 4 or not all(opts) or len({o.lower() for o in opts}) != 4:
            continue
        ans = str(it.get("answer") if it.get("answer") is not None else it.get("correct", "")).strip()
        idx = "ABCD".find(ans[:1].upper()) if re.match(r"^[A-Da-d]([\).:\s]|$)", ans) else next((i for i, o in enumerate(opts) if o.lower() == re.sub(r"^\s*[A-Da-d][\).:]\s+", "", ans).lower()), -1)
        if idx < 0:
            continue
        correct = opts[idx]
        random.shuffle(opts)                         # модели любят ставить правильный ответ на B/C — перемешиваем сами
        seen.add(q.lower())
        lvl = str(it.get("level", "medium")).lower()
        out.append({"q": q, "options": [f"{'ABCD'[i]}) {o}" for i, o in enumerate(opts)], "answer": "ABCD"[opts.index(correct)],
                    "explain": _str(it.get("explain") or it.get("explanation"), 500), "level": lvl if lvl in ("easy", "medium", "hard") else "medium"})
    return out[:10]


def _gen_questions(topic: str, comment: str, ctx: str, lang: str, asked: Optional[list] = None) -> list[dict]:
    avoid = ("\nНЕ повторяй и не перефразируй уже заданные вопросы:\n" + "\n".join(f"- {a}" for a in (asked or [])[-40:])) if asked else ""
    prompt = (
        f"Составь ровно 10 хороших тестовых вопросов по теме «{topic}», чтобы по-настоящему проверить понимание, а не память. Пожелания: {comment or 'нет'}. "
        + request_lang_line(topic + " " + comment, lang) + "\n"
        + (MATERIAL_RULES + QUESTION_RULES + f"Материал студента:\n{ctx}\n" if ctx else "")
        + "Сложность: 3 лёгких (термины, суть), 4 средних (применение, небольшие расчёты или ситуации), 3 сложных (анализ, ловушки, сравнение). "
          "У каждого вопроса 4 варианта и ровно один правильный; неправильные варианты правдоподобны (типичные ошибки), без «все вышеперечисленные» и «ничего из этого». "
          "Вопросы разные по смыслу, формулировки короткие и однозначные. В explain — 1–2 предложения: почему верный ответ верен и в чём ловушка (не упоминай буквы вариантов). "
          "Формулы пиши в $...$." + avoid + "\n"
          'Верни ТОЛЬКО JSON: {"questions":[{"q":"...","options":["...","...","...","..."],"answer":"B","explain":"...","level":"easy"}]}. '
          "В options пиши только текст вариантов без букв; answer — буква правильного варианта (A–D)."
    )
    best: list[dict] = []
    for attempt in range(2):
        data = parse_json(ask_ai(prompt, json_mode=True, max_tokens=5000, temperature=0.7 if attempt == 0 else 0.5, timeout=110.0))
        qs = _clean_questions(data.get("questions") if isinstance(data, dict) else data)
        if ctx and qs:
            qs = _clean_questions(_verify_questions([dict(x) for x in qs], ctx, lang)) or qs
        if len(qs) > len(best):
            best = qs
        if len(best) >= 9:
            break
    if len(best) < 6:
        raise HTTPException(502, "ИИ вернул неполный набор вопросов. Попробуйте ещё раз — токены вернутся.")
    return best


@app.post("/api/ai-teacher")
def ai_teacher(topic: str = Form(...), comment: str = Form(""), doc_id: str = Form("default"),
               lang: str = Form("kz"), user: dict = Depends(current_user)):
    topic, comment, lang = topic.strip()[:200], comment.strip()[:500], _lang(lang)
    if len(topic) < 2:
        raise HTTPException(400, "Введите тему.")
    ctx = fit_context(get_doc_text(doc_id, user) or "")

    def run():
        f_expl = _TEACH_POOL.submit(_teacher_explain, topic, comment, ctx, lang)      # объяснение и вопросы пишутся параллельно — быстрее
        f_qs = _TEACH_POOL.submit(_gen_questions, topic, comment, ctx, lang, None)
        expl, qs = f_expl.result(), f_qs.result()
        vlang = detect_lang(topic + " " + comment) or lang                          # видео подбираем на языке запроса
        query = expl["video_query"] or f"{topic} {'объяснение' if vlang != 'en' else 'explained'}"
        return {"explanation": expl["explanation"], "video_query": query, "video": find_video(query, vlang), "game_questions": qs}

    data, left = paid(user, run)
    save_generation(user["id"], "teacher", topic, json.dumps(data, ensure_ascii=False), doc_id)
    return {"result": data, "tokens_left": left}


@app.post("/api/ai-teacher/more")
def ai_teacher_more(topic: str = Form(...), comment: str = Form(""), doc_id: str = Form("default"), lang: str = Form("kz"),
                    asked: str = Form(""), user: dict = Depends(current_user)):
    """Ещё 10 вопросов по теме за 5 токенов (уже заданные не повторяются)."""
    have = int(user.get("tokens", 0))
    if have < MORE_QUESTIONS_COST:
        raise HTTPException(402, f"Для ещё 10 вопросов нужно {MORE_QUESTIONS_COST} токенов, у вас {have}. Токены обновляются каждый день, а Premium снимает лимиты.")
    topic, comment, lang = topic.strip()[:200], comment.strip()[:500], _lang(lang)
    try:
        prev = [_str(a, 200) for a in json.loads(asked)][:80] if asked else []
    except Exception:
        prev = []
    ctx = fit_context(get_doc_text(doc_id, user) or "")
    qs, left, _ = paid_n(user, MORE_QUESTIONS_COST, lambda: _gen_questions(topic, comment, ctx, lang, prev))
    return {"questions": qs, "tokens_left": left}


# ==================== ПОСТУПЛЕНИЕ ЗА РУБЕЖ + МОТИВАЦИОННЫЕ ПИСЬМА ====================
ABROAD_COST = 30   # токенов за подбор стран
LETTER_COST = 30   # токенов за мотивационное письмо
ABROAD_NOTE = ("Суммы, сроки и требования — ориентиры и меняются каждый год. Перед подачей обязательно сверьтесь "
               "с официальным сайтом программы или университета.")

# Проверенная база по странам: id → данные. Цифры — округлённые ориентиры для граждан Казахстана (в год, USD).
ABROAD_KB = {
    "de": {"name": "Германия", "flag": "🇩🇪", "lang": "немецкий (часть магистратур — английский)", "tuition": "0–500 (госуниверситеты почти бесплатны; сбор за семестр ~150–350 €)",
           "living": "11 000–13 000", "work": "до 120 полных дней / 240 полудней в год", "post_study": "18 месяцев на поиск работы",
           "langreq": "TestDaF/DSH (B2–C1) или IELTS 6.0–6.5 для программ на английском", "best_for": ["engineering", "it", "science", "economics", "humanities"],
           "intakes": "зимний семестр (заявки обычно до 15 июля, часть вузов — до 31 мая или 15 января), летний семестр",
           "scholarships": [
               {"name": "DAAD", "covers": "ежемесячная стипендия ~934 € (магистр), страховка, перелёт", "level": "магистратура, PhD", "link": "https://www.daad.de"},
               {"name": "Deutschlandstipendium", "covers": "300 € в месяц от вуза и частных спонсоров", "level": "бакалавриат, магистратура", "link": "https://www.deutschlandstipendium.de"}],
           "note": "Бесплатное обучение, но нужен блокированный счёт (~11 900 €/год) для визы. Для бакалавриата часто нужен подготовительный год (Studienkolleg)."},
    "cz": {"name": "Чехия", "flag": "🇨🇿", "lang": "чешский (бесплатно) или английский (платно)", "tuition": "0 на чешском; 2 000–7 000 на английском",
           "living": "7 000–9 000", "work": "без ограничений по часам при визе студента", "post_study": "9 месяцев на поиск работы",
           "langreq": "чешский B2 для бесплатных программ; IELTS 5.5–6.5 для английских", "best_for": ["engineering", "it", "medicine", "arts", "science"],
           "intakes": "осень; заявки чаще всего с февраля по апрель–май",
           "scholarships": [{"name": "Czech Government Scholarship", "covers": "стипендия и проживание для магистров/PhD из ряда стран", "level": "магистратура, PhD", "link": "https://www.msmt.cz"}],
           "note": "Один из самых доступных вариантов в ЕС по стоимости жизни. Бесплатный год чешского при подготовке."},
    "pl": {"name": "Польша", "flag": "🇵🇱", "lang": "польский / английский", "tuition": "2 000–5 000 (платные программы); есть стипендии на полное покрытие",
           "living": "5 500–8 000", "work": "разрешена работа студентам", "post_study": "9 месяцев на поиск работы",
           "langreq": "IELTS 5.5–6.5 или польский B2", "best_for": ["engineering", "it", "medicine", "economics", "science"],
           "intakes": "осень (заявки до июня–июля), некоторые вузы — весной",
           "scholarships": [{"name": "NAWA (Polish National Agency for Academic Exchange)", "covers": "программы для студентов и преподавателей, стипендии", "level": "магистратура, PhD", "link": "https://nawa.gov.pl"}],
           "note": "Дешёвая жизнь, ЕС-диплом, быстрые сроки подачи документов."},
    "hu": {"name": "Венгрия", "flag": "🇭🇺", "lang": "английский / венгерский", "tuition": "0 по Stipendium Hungaricum; иначе 3 000–8 000",
           "living": "6 500–9 000", "work": "разрешена работа студентам", "post_study": "9 месяцев на поиск работы",
           "langreq": "IELTS 5.5–6.5 / TOEFL; для медицины — вступительные экзамены", "best_for": ["medicine", "engineering", "it", "economics", "science"],
           "intakes": "осень; конкурс Stipendium Hungaricum — обычно январь–февраль",
           "scholarships": [{"name": "Stipendium Hungaricum", "covers": "полное обучение, ежемесячная стипендия, жильё, медстраховка", "level": "бакалавриат, магистратура, PhD", "link": "https://stipendiumhungaricum.hu"}],
           "note": "Казахстан — страна-участница программы Stipendium Hungaricum, один из самых реалистичных грантов на полное покрытие."},
    "it": {"name": "Италия", "flag": "🇮🇹", "lang": "итальянский / английский", "tuition": "500–4 000 (сбор зависит от дохода семьи)",
           "living": "8 000–12 000", "work": "до 20 часов в неделю", "post_study": "12 месяцев на поиск работы",
           "langreq": "IELTS 5.5–6.5 или итальянский B2", "best_for": ["arts", "engineering", "humanities", "economics", "medicine"],
           "intakes": "осень; предварительная регистрация через Universitaly в начале года",
           "scholarships": [{"name": "Региональные стипендии DSU (диритто алло студио)", "covers": "освобождение от платы, питание, жильё по критериям дохода и оценок", "level": "все уровни", "link": "https://www.universitaly.it"},
                            {"name": "Italian Government Scholarships (MAECI)", "covers": "ежемесячная стипендия и страховка", "level": "магистратура, PhD", "link": "https://studyinitaly.esteri.it"}],
           "note": "Сильны дизайн, архитектура, инженерия. Сбор снижается при подаче справок о доходах."},
    "tr": {"name": "Турция", "flag": "🇹🇷", "lang": "турецкий / английский", "tuition": "0 по Türkiye Burslari; иначе 2 000–10 000",
           "living": "4 500–7 000", "work": "ограничено", "post_study": "по короткому разрешению",
           "langreq": "английский B2 / турецкий B2; год языковой подготовки при стипендии", "best_for": ["engineering", "medicine", "economics", "humanities", "it"],
           "intakes": "осень; заявки на Türkiye Burslari — обычно январь–февраль",
           "scholarships": [{"name": "Türkiye Burslari", "covers": "обучение, жильё, ежемесячная стипендия, перелёт, год турецкого языка", "level": "бакалавриат, магистратура, PhD", "link": "https://turkiyeburslari.gov.tr"}],
           "note": "Близкая культура и язык для казахстанцев, недорогая жизнь, сильная государственная стипендия."},
    "kr": {"name": "Южная Корея", "flag": "🇰🇷", "lang": "корейский / английский", "tuition": "0 по GKS; иначе 4 000–9 000",
           "living": "8 000–11 000", "work": "до 20–25 часов в неделю", "post_study": "до 2 лет по визе поиска работы",
           "langreq": "TOPIK 3–4 / IELTS 5.5–6.5; год языка при GKS", "best_for": ["it", "engineering", "arts", "economics", "science"],
           "intakes": "весна и осень; GKS — обычно февраль–март",
           "scholarships": [{"name": "Global Korea Scholarship (GKS)", "covers": "обучение, ~900 000 вон в месяц, перелёт, год корейского", "level": "бакалавриат, магистратура, PhD", "link": "https://www.studyinkorea.go.kr"}],
           "note": "Сильные IT и электроника. GKS — одна из лучших стипендий мира по уровню покрытия."},
    "jp": {"name": "Япония", "flag": "🇯🇵", "lang": "японский / английский", "tuition": "3 500–6 000 (гос. университеты); 0 по MEXT",
           "living": "9 000–12 000", "work": "до 28 часов в неделю", "post_study": "до 1 года по визе поиска работы",
           "langreq": "JLPT N2 / IELTS 6.0; EJU для бакалавриата", "best_for": ["engineering", "it", "science", "arts", "humanities"],
           "intakes": "апрель и октябрь; MEXT — через посольство, обычно апрель–май",
           "scholarships": [{"name": "MEXT Scholarship", "covers": "обучение, ~117 000–145 000 иен в месяц, перелёт", "level": "бакалавриат, магистратура, PhD", "link": "https://www.studyinjapan.go.jp"}],
           "note": "Высокий уровень науки и технологий, но высокая конкуренция за MEXT."},
    "cn": {"name": "Китай", "flag": "🇨🇳", "lang": "китайский / английский", "tuition": "0 по CSC; иначе 2 500–6 000",
           "living": "3 500–6 000", "work": "ограничено", "post_study": "ограничено",
           "langreq": "HSK 4–5 / IELTS 6.0", "best_for": ["engineering", "it", "medicine", "economics", "science"],
           "intakes": "осень; CSC — подача обычно с января по апрель",
           "scholarships": [{"name": "Chinese Government Scholarship (CSC)", "covers": "обучение, жильё, ежемесячная стипендия, страховка", "level": "бакалавриат, магистратура, PhD", "link": "https://www.campuschina.org"},
                            {"name": "Silk Road Scholarship", "covers": "полная или частичная стипендия для стран «Пояса и пути»", "level": "магистратура, PhD", "link": "https://www.campuschina.org"}],
           "note": "Недорогая жизнь, огромный выбор грантов, быстро растущие университеты."},
    "uk": {"name": "Великобритания", "flag": "🇬🇧", "lang": "английский", "tuition": "15 000–35 000+",
           "living": "14 000–20 000", "work": "до 20 часов в неделю", "post_study": "Graduate Route — 2 года (3 для PhD)",
           "langreq": "IELTS 6.0–7.5", "best_for": ["economics", "humanities", "science", "engineering", "law"],
           "intakes": "осень; UCAS до 31 января (Oxbridge — до 15 октября)",
           "scholarships": [{"name": "Chevening", "covers": "полное обучение, проживание, перелёт, виза", "level": "магистратура (1 год)", "link": "https://www.chevening.org"},
                            {"name": "Болашак (для граждан Казахстана)", "covers": "полное покрытие для граждан Казахстана", "level": "магистратура, PhD", "link": "https://bolashak.gov.kz"}],
           "note": "Магистратура длится 1 год, что экономит время и деньги. Chevening требует опыт работы от 2 лет."},
    "us": {"name": "США", "flag": "🇺🇸", "lang": "английский", "tuition": "20 000–60 000+ (с финансовой помощью дешевле)",
           "living": "12 000–20 000", "work": "на кампусе до 20 часов", "post_study": "OPT 1–3 года по специальности",
           "langreq": "TOEFL 80–100+ / IELTS 6.5–7.5, SAT/ACT (часть вузов — test-optional)", "best_for": ["it", "science", "economics", "engineering", "arts"],
           "intakes": "осень; Early Action в ноябре, Regular — до января",
           "scholarships": [{"name": "Need-based aid в топ-вузах (Harvard, MIT, Yale и др.)", "covers": "до 100% потребности семьи, без возврата", "level": "бакалавриат", "link": "https://www.commonapp.org"},
                            {"name": "Fulbright", "covers": "полное покрытие магистратуры/PhD", "level": "магистратура, PhD", "link": "https://foreign.fulbrightonline.org"},
                            {"name": "Болашак", "covers": "полное покрытие для граждан Казахстана", "level": "магистратура, PhD", "link": "https://bolashak.gov.kz"}],
           "note": "Самая дорогая страна, но у богатых вузов есть полная помощь иностранцам — ключевое: сильное эссе, внеучебка и оценки."},
    "ca": {"name": "Канада", "flag": "🇨🇦", "lang": "английский / французский", "tuition": "12 000–30 000",
           "living": "12 000–16 000", "work": "до 24 часов в неделю", "post_study": "PGWP до 3 лет",
           "langreq": "IELTS 6.0–7.0 / TOEFL 80–100", "best_for": ["it", "engineering", "science", "economics", "medicine"],
           "intakes": "осень, часть — зима/лето; заявки за 8–10 месяцев",
           "scholarships": [{"name": "Vanier CGS / университетские гранты", "covers": "для PhD ~50 000 CAD в год; для бакалавриата — вузовские награды", "level": "PhD, бакалавриат", "link": "https://vanier.gc.ca"},
                            {"name": "Lester B. Pearson (Toronto)", "covers": "полное покрытие бакалавриата", "level": "бакалавриат", "link": "https://future.utoronto.ca/pearson"}],
           "note": "Удобный путь к ПМЖ, но жильё в больших городах дорогое."},
    "au": {"name": "Австралия", "flag": "🇦🇺", "lang": "английский", "tuition": "18 000–40 000",
           "living": "16 000–22 000", "work": "до 48 часов в две недели", "post_study": "2–4 года по Temporary Graduate visa",
           "langreq": "IELTS 6.0–7.0", "best_for": ["engineering", "it", "medicine", "science", "economics"],
           "intakes": "февраль и июль",
           "scholarships": [{"name": "Australia Awards", "covers": "полное обучение и проживание", "level": "магистратура", "link": "https://www.dfat.gov.au/people-to-people/australia-awards"},
                            {"name": "Университетские гранты (например, ANU, Melbourne)", "covers": "частичное или полное покрытие по оценкам", "level": "все уровни", "link": "https://www.studyaustralia.gov.au"}],
           "note": "Высокая цена жизни, но хорошая заработная плата для подработки."},
    "ae": {"name": "ОАЭ", "flag": "🇦🇪", "lang": "английский", "tuition": "8 000–25 000",
           "living": "9 000–14 000", "work": "ограничено", "post_study": "золотая виза для лучших выпускников",
           "langreq": "IELTS 5.5–6.5", "best_for": ["economics", "engineering", "it", "arts"],
           "intakes": "осень и весна",
           "scholarships": [{"name": "Вузовские гранты (NYU Abu Dhabi, Khalifa University)", "covers": "NYUAD — полное покрытие по конкурсу", "level": "бакалавриат, магистратура", "link": "https://nyuad.nyu.edu"}],
           "note": "Недалеко от Казахстана, без налога на доход. NYU Abu Dhabi очень конкурентен, но даёт полное финансирование."},
    "my": {"name": "Малайзия", "flag": "🇲🇾", "lang": "английский", "tuition": "3 500–9 000",
           "living": "4 000–6 000", "work": "ограничено", "post_study": "ограничено",
           "langreq": "IELTS 5.5–6.5", "best_for": ["engineering", "it", "economics", "medicine"],
           "intakes": "февраль, июнь, сентябрь",
           "scholarships": [{"name": "Malaysia International Scholarship (MIS)", "covers": "обучение и ежемесячные выплаты", "level": "магистратура, PhD", "link": "https://biasiswa.mohe.gov.my"}],
           "note": "Недорогое обучение на английском с филиалами британских и австралийских вузов."},
    "kz_ab": {"name": "Двойные дипломы и обмен (из Казахстана)", "flag": "🎓", "lang": "английский", "tuition": "по условиям вашего вуза",
           "living": "зависит от страны", "work": "по условиям страны", "post_study": "по условиям страны",
           "langreq": "IELTS 6.0–6.5", "best_for": ["any"],
           "intakes": "осень и весна; заявка в вузе за 6–12 месяцев",
           "scholarships": [{"name": "Erasmus+ (обмен и Joint Master)", "covers": "ежемесячная стипендия, перелёт, обучение для участников", "level": "бакалавриат, магистратура", "link": "https://erasmus-plus.ec.europa.eu"},
                            {"name": "Болашак", "covers": "полное покрытие для граждан Казахстана", "level": "магистратура, PhD, стажировка", "link": "https://bolashak.gov.kz"}],
           "note": "Часто лучший первый шаг: семестр за рубежом без риска потерять место дома."},
}
ABROAD_KB_PROMPT = "\n".join(
    f"{k} | {v['name']} | язык: {v['lang']} | обучение $/год: {v['tuition']} | жизнь $/год: {v['living']} | сильные направления: {', '.join(v['best_for'])} | "
    f"стипендии: {'; '.join(s['name'] for s in v['scholarships'])}"
    for k, v in ABROAD_KB.items())


class AbroadIn(BaseModel):
    answers: list[dict] = Field(max_length=10)
    level: str = "bachelor"   # bachelor | master | phd | exchange
    duration: str = ""
    budget: str = ""
    lang: str = "ru"


@app.get("/api/abroad/info")
def abroad_info():
    """Публичный каталог стран и стипендий — стартовый экран раздела."""
    return {"countries": [{"id": k, **v} for k, v in ABROAD_KB.items()], "note": ABROAD_NOTE, "cost": ABROAD_COST, "letter_cost": LETTER_COST}


@app.post("/api/abroad")
def abroad(body: AbroadIn, user: dict = Depends(current_user)):
    rate_limit(f"abroad:{user['id']}", 10, 3600)
    ans = "\n".join(f"- {_str(a.get('q'), 200)} → {_str(a.get('a'), 300)}" for a in body.answers)
    level_names = {"bachelor": "бакалавриат", "master": "магистратура", "phd": "PhD", "exchange": "семестр по обмену"}
    prompt = (
        "Ты — консультант по поступлению в зарубежные вузы для студентов из Казахстана. Студент ответил на вопросы.\n"
        f"ОТВЕТЫ:\n{ans}\nУровень: {level_names.get(body.level, 'бакалавриат')}. Срок обучения: {_str(body.duration, 80) or 'не указан'}. "
        f"Бюджет: {_str(body.budget, 120) or 'не указан'}.\n\n"
        f"ПРОВЕРЕННАЯ БАЗА СТРАН: id | страна | ...\n{ABROAD_KB_PROMPT}\n\n"
        "ЗАДАЧА: выбери 5–6 стран из БАЗЫ (только их id), которые лучше всего подходят студенту, и отсортируй по убыванию. Для каждой:\n"
        "• match — число 0–100: насколько страна подходит по интересам, уровню, сроку, бюджету и языку.\n"
        "• why — 2–3 предложения: почему именно эта страна подходит ЭТОМУ студенту, с опорой на его ответы.\n"
        "• programs — 2–3 направления/программы, которые стоит рассмотреть (названия реальных направлений, без выдуманных вузов).\n"
        "• unis — 2–3 реальных известных университета страны. Не уверен в вузе — не включай.\n"
        "• steps — 3–4 коротких конкретных шага подготовки именно для этой страны.\n"
        "• risk — 1 предложение о главном риске или сложности.\n"
        "Числа стоимости и стипендии НЕ пиши, сервер подставит их сам. "
        "profile — краткий портрет студента (2 предложения). tips — 4 общих совета по подготовке документов.\n"
        + lang_line(body.lang) + "\n"
        'Верни ТОЛЬКО JSON: {"profile":"...","countries":[{"id":"de","match":85,"why":"...","programs":["..."],"unis":["..."],'
        '"steps":["..."],"risk":"..."}],"tips":["..."]}'
    )

    def run():
        d = parse_json(ask_ai(prompt, json_mode=True, max_tokens=4500, temperature=0.5))
        if not isinstance(d, dict) or not isinstance(d.get("countries"), list) or not d["countries"]:
            raise HTTPException(502, "ИИ вернул неверный формат. Попробуйте ещё раз.")
        return d

    d, left, _ = paid_n(user, ABROAD_COST, run)
    out, seen = [], set()
    for c in d["countries"][:8]:
        if not isinstance(c, dict) or c.get("id") not in ABROAD_KB or c["id"] in seen:
            continue
        seen.add(c["id"])
        kb = ABROAD_KB[c["id"]]
        try:
            match = max(0, min(100, int(c.get("match"))))
        except (TypeError, ValueError):
            match = 60
        out.append({"id": c["id"], "match": match, "why": _str(c.get("why"), 500), "programs": _list(c.get("programs"), 4, 120),
                    "unis": _list(c.get("unis"), 4, 90), "steps": _list(c.get("steps"), 5, 200), "risk": _str(c.get("risk"), 300), **kb})
    if not out:
        raise HTTPException(502, "ИИ не подобрал страны. Попробуйте ещё раз.")
    out.sort(key=lambda x: -x["match"])
    result = {"profile": _str(d.get("profile"), 500), "countries": out, "tips": _list(d.get("tips"), 6, 220), "note": ABROAD_NOTE, "tokens_left": left}
    gid = save_generation(user["id"], "abroad", "Поступление за рубеж: " + ", ".join(c["name"] for c in out[:3]), json.dumps(result, ensure_ascii=False))
    result["history_id"] = gid
    return result


class LetterIn(BaseModel):
    kind: str = "admission"        # admission | exchange | scholarship
    country: str = Field(default="", max_length=80)
    university: str = Field(max_length=120)
    program: str = Field(max_length=160)
    name: str = Field(default="", max_length=80)
    home_uni: str = Field(default="", max_length=120)
    studying: str = Field(default="", max_length=800)       # что изучает сейчас
    achievements: str = Field(default="", max_length=1200)  # проекты, результаты, исследования
    why_country: str = Field(default="", max_length=800)
    beyond: str = Field(default="", max_length=800)         # хобби, волонтёрство, спорт
    goals: str = Field(default="", max_length=800)          # как применит знания
    tone: str = "formal"          # formal | warm
    length: str = "medium"        # short | medium | long
    letter_lang: str = "en"       # en | ru | kz
    lang: str = "ru"


LETTER_KINDS = {
    "admission": "поступление на программу (бакалавриат/магистратура)",
    "exchange": "семестр по обмену из университета-партнёра",
    "scholarship": "конкурс на стипендию",
}
LETTER_LANGS = {"en": "английском", "ru": "русском", "kz": "казахском"}
LETTER_WORDS = {"short": "250–320", "medium": "350–450", "long": "480–600"}


@app.post("/api/letter")
def letter(body: LetterIn, user: dict = Depends(current_user)):
    rate_limit(f"letter:{user['id']}", 15, 3600)
    kind = body.kind if body.kind in LETTER_KINDS else "admission"
    ll = body.letter_lang if body.letter_lang in LETTER_LANGS else "en"
    words = LETTER_WORDS.get(body.length, LETTER_WORDS["medium"])
    uni, prog = _str(body.university, 120), _str(body.program, 160)
    if not uni or not prog:
        raise HTTPException(400, "Укажите университет и программу.")
    facts = "\n".join(f"{k}: {_str(v, 1200)}" for k, v in [
        ("Имя автора", body.name), ("Страна", body.country), ("Университет", uni), ("Программа", prog), ("Родной вуз / школа", body.home_uni),
        ("Что изучает сейчас", body.studying), ("Проекты, результаты, исследования", body.achievements),
        ("Почему эта страна / город", body.why_country), ("Интересы вне учёбы (волонтёрство, спорт, творчество)", body.beyond),
        ("Цели и как применит знания", body.goals)] if _str(v, 1200))
    tone = ("тёплый и личный, но по-прежнему академичный" if body.tone == "warm" else "строгий академический, уверенный, без пафоса")
    prompt = (
        f"Ты — опытный редактор мотивационных писем для абитуриентов и студентов из Казахстана. Напиши мотивационное письмо: {LETTER_KINDS[kind]}.\n"
        f"ФАКТЫ ОТ АВТОРА (используй ТОЛЬКО их, ничего не выдумывай):\n{facts}\n\n"
        "СТРУКТУРА (строго соблюдай):\n"
        f"1. Приветствие: «Dear Members of the Selection Committee,» (на {LETTER_LANGS[ll]} языке — эквивалент). Затем ровно 4 абзаца, затем прощание и подпись.\n"
        f"2. Абзац 1 — зацепка: сразу назови конкретный университет «{uni}» и программу «{prog}», кратко объясни, почему выбрать нужно именно автора, "
        "покажи энтузиазм и как программа поможет в учёбе, карьере или личностном росте. Пиши этот абзац в последнюю очередь, чтобы он точно отражал письмо.\n"
        "3. Абзац 2 — академический опыт: что изучает, в каких проектах участвовал, чего добился (продукт, приложение, исследование, конференции, тема курсовой/диплома). "
        "Обязательно свяжи это с конкретными направлениями или курсами программы (без выдуманных названий курсов и профессоров).\n"
        "4. Абзац 3 — Beyond academics: почему привлекла эта страна/город, как период обучения поможет личностно (знакомства, студенческая жизнь, культурный обмен), "
        "внеучебные интересы автора и как они связаны с культурой страны.\n"
        "5. Абзац 4 — итог: кратко повтори главные мысли ДРУГИМИ словами, подчеркни важность участия, как применит знания, как использует компетенции по возвращении домой.\n"
        "6. Закончи деловым прощанием («Yours sincerely,» / «Sincerely,» или эквивалент) и именем автора.\n\n"
        "СТИЛЬ:\n"
        f"• Тон: {tone}. Академичный формальный стиль, нейтральная лексика, никаких разговорных сокращений (do not, а не don't).\n"
        "• Голос живого человека: конкретные детали из фактов, разная длина предложений, одна-две личные детали. "
        "ЗАПРЕЩЕНЫ клише ИИ: «I am writing to express my profound interest», «passionate about», «in today's fast-paced world», «testament to», «tapestry», «delve», "
        "«unwavering», «I am confident that», «cutting-edge» и их аналоги на других языках. Без восклицательных знаков и пустого пафоса.\n"
        "• Если фактов для абзаца мало — пиши короче и честнее, не придумывай достижения, цифры, награды, названия. Не вставляй плейсхолдеры в квадратных скобках, "
        "кроме имени автора, если оно не указано: тогда подпиши «[Your Name]».\n"
        f"• Объём письма: {words} слов. Язык письма: {LETTER_LANGS[ll]}. Абзацы разделяй пустой строкой.\n"
        f"• Также верни: subject — тема письма (1 строка), tips — 3 коротких совета на {('русском')} языке, что автору стоит дополнить в письме своим голосом, "
        "чтобы оно не выглядело написанным ИИ.\n"
        'Верни ТОЛЬКО JSON: {"subject":"...","letter":"полный текст письма с приветствием и подписью","tips":["..."]}'
    )

    def run():
        d = parse_json(ask_ai(prompt, json_mode=True, max_tokens=3000, temperature=0.75))
        if not isinstance(d, dict) or len(_str(d.get("letter"), 6000)) < 200:
            raise HTTPException(502, "ИИ вернул неверный формат. Попробуйте ещё раз.")
        return d

    d, left, _ = paid_n(user, LETTER_COST, run)
    text = _str(d.get("letter"), 6000)
    result = {"subject": _str(d.get("subject"), 200), "letter": text, "tips": _list(d.get("tips"), 4, 220),
              "words": len(text.split()), "tokens_left": left}
    result["history_id"] = save_generation(user["id"], "letter", f"Письмо: {uni}", json.dumps(result, ensure_ascii=False))
    return result


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
