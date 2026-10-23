from __future__ import annotations
import os
import time
import random
import shutil
import urllib.request
import urllib.parse
import json
import asyncio
from datetime import datetime
from fastapi import FastAPI, File, Form, UploadFile, Request
from fastapi.responses import FileResponse, JSONResponse
from openai import OpenAI
from pypdf import PdfReader
from supabase import create_client, Client

app = FastAPI(title="StudyAI Platform")

ADMIN_SECRET = os.getenv("ADMIN_SECRET", "admin12345")

DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "")
ai_client = OpenAI(
    api_key=DEEPSEEK_API_KEY,
    base_url="https://api.deepseek.com"
)

SUPABASE_URL = os.getenv("SUPABASE_URL", "")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "")
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY) if SUPABASE_URL and SUPABASE_KEY else None

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")

active_documents = {}

def send_telegram_message(chat_id: str, text: str):
    if not TELEGRAM_BOT_TOKEN:
        return False
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    data = urllib.parse.urlencode({"chat_id": chat_id, "text": text, "parse_mode": "Markdown"}).encode("utf-8")
    try:
        req = urllib.request.Request(url, data=data)
        with urllib.request.urlopen(req) as response:
            return response.status == 200
    except Exception:
        return False

@app.get("/", response_class=FileResponse)
async def read_root():
    return "templates/index.html"

@app.get("/admin", response_class=FileResponse)
async def read_admin():
    return "templates/admin.html"

# --- АДМИН-ПАНЕЛЬ API ---
@app.get("/api/admin/users")
async def admin_get_users(admin_key: str):
    if admin_key != ADMIN_SECRET:
        return JSONResponse({"error": "Доступ запрещен! Неверный ключ администратора."}, status_code=403)
    try:
        if not supabase:
            return JSONResponse({"error": "База данных не подключена"}, status_code=500)
        res = supabase.table("users").select("*").order("id", desc=True).execute()
        return {"users": res.data or []}
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

@app.post("/api/admin/toggle-subscription")
async def admin_toggle_subscription(
    user_id: int = Form(...),
    status: bool = Form(...),
    admin_key: str = Form(...)
):
    if admin_key != ADMIN_SECRET:
        return JSONResponse({"error": "Доступ запрещен! Неверный ключ администратора."}, status_code=403)
    try:
        if not supabase:
            return JSONResponse({"error": "База данных не подключена"}, status_code=500)
        
        is_sub_bool = True if str(status).lower() in ["true", "1", "yes"] else False
        supabase.table("users").update({"is_subscribed": is_sub_bool}).eq("id", user_id).execute()
        return {"success": True, "user_id": user_id, "is_subscribed": is_sub_bool}
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

# --- API НАПОМИНАНИЙ И УВЕДОМЛЕНИЙ ---
@app.post("/api/telegram/reminder")
async def send_task_reminder(
    chat_id: str = Form(...),
    task_title: str = Form(...),
    task_date: str = Form(...)
):
    text = f"⏰ *Напоминание от StudyAI Hub*\n\nУ вас запланирована задача:\n📌 *{task_title}*\n📅 Дата: `{task_date}`\n\nНе забудьте выполнить её вовремя!"
    success = send_telegram_message(chat_id, text)
    if success:
        return {"success": True, "message": "Напоминание успешно отправлено в Telegram!"}
    return JSONResponse({"error": "Не удалось отправить сообщение. Проверьте токен бота."}, status_code=400)

# --- ЧАТ-БОТ API ---
@app.post("/api/chat")
async def chat_with_ai(
    user_id: int = Form(...),
    message: str = Form(...),
    doc_id: str = Form("default"),
    lang: str = Form("ru")
):
    try:
        if supabase:
            u_res = supabase.table("users").select("*").eq("id", user_id).execute()
            if not u_res.data:
                return JSONResponse({"error": "Пользователь не найден."}, status_code=401)
            user = u_res.data[0]
            tokens, is_sub = user.get("tokens", 0), user.get("is_subscribed", False)
            if not is_sub and tokens <= 0:
                return JSONResponse({"error": "PAYMENT_REQUIRED", "message": "Токены закончились!"}, status_code=402)
        else:
            tokens, is_sub = 3, True

        user_text = active_documents.get(doc_id, "")
        lang_instr = {
            "kz": "СТРОГО ҚАЗАҚ ТІЛІНДЕ жауап бер.",
            "en": "STRICTLY IN ENGLISH.",
            "ru": "Отвечай строго на русском языке."
        }.get(lang, "Отвечай строго на русском.")

        prompt = f"""{lang_instr}
Ты интеллектуальный ассистент-наставник StudyAI Hub. Отвечай кратко, грамотно и структурировано.
Контекст текущего проекта/материала:
{user_text[:4000] if user_text else 'Общий контекст учебы'}

Вопрос пользователя: {message}"""

        response = ai_client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "Ты экспертный ИИ-наставник."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.4
        )
        reply = response.choices[0].message.content
        new_tokens = tokens if is_sub else max(0, tokens - 1)
        if supabase:
            supabase.table("users").update({"tokens": new_tokens}).eq("id", user_id).execute()

        return {"reply": reply, "tokens_left": new_tokens, "is_subscribed": is_sub}
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

# --- АВТОРИЗАЦИЯ ТЕЛЕГРАМ ---
@app.post("/api/telegram/webhook")
async def telegram_webhook(request: Request):
    try:
        body = await request.json()
        message = body.get("message", {})
        chat = message.get("chat", {})
        chat_id = chat.get("id")
        text = message.get("text", "")
        
        if text.startswith("/start") and chat_id:
            code = str(random.randint(100000, 999999))
            str_chat_id = str(chat_id)
            if supabase:
                supabase.table("auth_codes").delete().eq("tg_id", str_chat_id).execute()
                supabase.table("auth_codes").insert({"tg_id": str_chat_id, "code": code}).execute()
            
            msg = f"👋 *StudyAI Hub — Авторизация*\n\nВаш 6-значный код для входа:\n`{code}`\n\nВведите его на сайте."
            send_telegram_message(str_chat_id, msg)
            
        return {"status": "ok"}
    except Exception as e:
        return {"status": "error", "message": str(e)}

@app.post("/api/auth/verify-code")
async def verify_code(code: str = Form(...)):
    code_clean = code.strip()
    try:
        if not supabase:
            return JSONResponse({"error": "База данных не подключена"}, status_code=500)

        res = supabase.table("auth_codes").select("*").eq("code", code_clean).execute()
        if not res.data or len(res.data) == 0:
            return JSONResponse({"error": "Неверный или устаревший код подтверждения!"}, status_code=400)
        
        record = res.data[0]
        tg_id = record["tg_id"]
        supabase.table("auth_codes").delete().eq("code", code_clean).execute()
        
        unique_username = f"tg_{tg_id}"
        user_res = supabase.table("users").select("*").eq("username", unique_username).execute()
        
        if user_res.data and len(user_res.data) > 0:
            user = user_res.data[0]
        else:
            new_user = supabase.table("users").insert({
                "username": unique_username,
                "password": "telegram_auth",
                "score": 0,
                "tokens": 3,
                "is_subscribed": False
            }).execute()
            user = new_user.data[0]

        return {
            "success": True,
            "user_id": user["id"],
            "username": unique_username,
            "telegram_id": str(tg_id),
            "display_name": f"Студент #{str(tg_id)[-4:]}",
            "score": user.get("score", 0),
            "tokens": user.get("tokens", 3),
            "is_subscribed": user.get("is_subscribed", False)
        }
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

@app.post("/upload")
async def upload_pdf(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        return JSONResponse({"error": "Загрузите PDF файл."}, status_code=400)
    temp_filename = f"temp_{file.filename}"
    try:
        with open(temp_filename, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        reader = PdfReader(temp_filename)
        text = "".join([page.extract_text() + "\n" for page in reader.pages if page.extract_text()])
        if os.path.exists(temp_filename): os.remove(temp_filename)
        if len(text.strip()) < 30: return JSONResponse({"error": "В PDF нет текста."}, status_code=400)
        doc_id = str(os.urandom(4).hex())
        active_documents[doc_id] = text
        return {"doc_id": doc_id, "message": "Файл успешно обработан!"}
    except Exception as e:
        if os.path.exists(temp_filename): os.remove(temp_filename)
        return JSONResponse({"error": str(e)}, status_code=500)

@app.post("/generate")
async def generate_content(
    doc_id: str = Form("default"), 
    action: str = Form(...), 
    user_id: int = Form(...), 
    lang: str = Form("ru"), 
    title: str = Form("Урок"),
    page: int = Form(1),
    custom_prompt: str = Form("")
):
    try:
        if supabase:
            u_res = supabase.table("users").select("*").eq("id", user_id).execute()
            if not u_res.data: 
                return JSONResponse({"error": "Пользователь не найден."}, status_code=401)
            user = u_res.data[0]
            tokens, is_sub = user.get("tokens", 0), user.get("is_subscribed", False)
            
            if not is_sub and tokens <= 0:
                return JSONResponse({
                    "error": "PAYMENT_REQUIRED", 
                    "message": "Подписка не активна и токены закончились!",
                    "is_subscribed": False,
                    "tokens_left": 0
                }, status_code=402)
        else:
            tokens, is_sub = 3, True
    except Exception:
        return JSONResponse({"error": "Ошибка проверки профиля"}, status_code=500)

    user_text = active_documents.get(doc_id, "")
    lang_instr = {
        "kz": "Жауапты СТРОГО МІНСІЗ ҚАЗАҚ ТІЛІНДЕ жаз.",
        "en": "Provide the response STRICTLY IN ENGLISH.",
        "ru": "Предоставь ответ СТРОГО НА РУССКОМ ЯЗЫКЕ."
    }.get(lang, "Отвечай строго на русском.")

    start_q = (page - 1) * 20 + 1
    end_q = page * 20

    if action == "planner":
        prompt_text = f"""{lang_instr}
Ты персональный ИИ-ассистент по тайм-менеджменту и планированию для студентов.
Контекст и задачи студента:
"{custom_prompt}"

Составь текстовый отчет по графику. А В САМОМ КОНЦЕ ОТВЕТА ОБЯЗАТЕЛЬНО добавь JSON-блок с перераспределенным списком задач для календаря:
```json
[
  {{"title": "Название задачи", "date": "YYYY-MM-DD", "desc": "Подробная информация и инструкция по выполнению задачи"}}
]
```"""
    elif action == "quiz":
        prompt_text = f"""{lang_instr}
Составь СТРОГО 20 УНИКАЛЬНЫХ тестовых вопросов (с номера {start_q} по номер {end_q}).
Верни ответ СТРОГО в формате валидного JSON-массива без какого-либо дополнительного текста или тегов markdown!
Формат каждого элемента JSON:
{{
  "id": 1,
  "question": "Текст вопроса...",
  "options": ["A) Вариант 1", "B) Вариант 2", "C) Вариант 3", "D) Вариант 4"],
  "answer": "A",
  "explanation": "Объяснение."
}}"""
    elif action == "flashcards":
        prompt_text = f"""{lang_instr}
Создай 10 ИНТЕРАКТИВНЫХ УЧЕБНЫХ КАРТОЧЕК. Оформляй СТРОГО по шаблону:
### Карточка X
**Вопрос:** [...]
**Ответ:** [...]
---"""
    elif action == "summary":
        prompt_text = f"{lang_instr}\nСделай качественный конспект с формулами ($...$) и главными терминами."
    else:
        prompt_text = f"{lang_instr}\nОбъясни материал простыми словами."

    try:
        content_text = (user_text or doc_id)[:7000] if action != "planner" else custom_prompt
        response = ai_client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "Ты экспертный ИИ-наставник и органайзер."},
                {"role": "user", "content": f"{prompt_text}\n\nКонтекст:\n{content_text}"}
            ],
            temperature=0.3
        )
        res_text = response.choices[0].message.content
        new_tokens = tokens if is_sub else max(0, tokens - 1)
        
        if supabase:
            supabase.table("users").update({"tokens": new_tokens}).eq("id", user_id).execute()
            supabase.table("saved_materials").insert({"user_id": user_id, "title": title, "content": res_text, "type": action}).execute()

        return {"result": res_text, "tokens_left": new_tokens, "is_subscribed": is_sub, "page": page}
    except Exception as e:
        return JSONResponse({"error": f"Ошибка DeepSeek API: {str(e)}"}, status_code=500)

@app.post("/api/user/add-score")
async def add_score(user_id: int = Form(...), points: int = Form(50)):
    try:
        if supabase:
            u_res = supabase.table("users").select("score").eq("id", user_id).execute()
            if u_res.data:
                current_score = u_res.data[0].get("score", 0)
                new_score = current_score + points
                supabase.table("users").update({"score": new_score}).eq("id", user_id).execute()
                return {"success": True, "new_score": new_score}
        return {"success": True, "new_score": points}
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

@app.post("/api/user/subscribe")
async def subscribe_user(user_id: int = Form(...)):
    try:
        if supabase:
            supabase.table("users").update({"is_subscribed": True}).eq("id", user_id).execute()
        return {"success": True}
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

@app.get("/api/user/{user_id}/saves")
async def get_user_saves(user_id: int):
    try:
        if supabase:
            saves = supabase.table("saved_materials").select("*").eq("user_id", user_id).order("id", desc=True).execute()
            user = supabase.table("users").select("*").eq("id", user_id).execute().data[0]
            return {
                "username": user.get("username"), 
                "score": user.get("score", 0), 
                "tokens": user.get("tokens"), 
                "is_subscribed": user.get("is_subscribed", False), 
                "saves": saves.data or []
            }
        return {"username": "Студент", "score": 100, "tokens": 3, "is_subscribed": True, "saves": []}
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
