from __future__ import annotations
import os
import random
import shutil
import urllib.request
import urllib.parse
import json
import asyncio
from datetime import datetime, timedelta
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

# --- ФОНОВЫЙ ПЛАНИРОВЩИК (2026 ГОД) ---
async def background_reminder_checker():
    while True:
        try:
            if supabase:
                today_str = datetime.now().strftime("%Y-%m-%d")
                tasks_res = supabase.table("user_tasks").select("*").eq("date", today_str).eq("notified", False).execute()
                if tasks_res.data:
                    for task in tasks_res.data:
                        tg_id = task.get("tg_id")
                        title = task.get("title", "Учебная задача")
                        if tg_id:
                            msg = f"🔔 *НАПОМИНАНИЕ НА СЕГОДНЯ!*\n\n📌 Задача: *{title}*\n📅 Дата: `{today_str}`\n\nПриступайте к выполнению через StudyAI Hub!"
                            if send_telegram_message(str(tg_id), msg):
                                supabase.table("user_tasks").update({"notified": True}).eq("id", task["id"]).execute()
        except Exception as e:
            print("Background reminder error:", e)
        await asyncio.sleep(3600)

@app.on_event("startup")
async def startup_event():
    asyncio.create_task(background_reminder_checker())

@app.get("/", response_class=FileResponse)
async def read_root():
    return "templates/index.html"

@app.get("/admin", response_class=FileResponse)
async def read_admin():
    return "templates/admin.html"

@app.post("/api/telegram/reminder")
async def send_task_reminder(
    chat_id: str = Form(...),
    task_title: str = Form(...),
    task_date: str = Form(...),
    user_id: int = Form(0)
):
    text = f"✅ *Напоминание успешно создано!*\n\n📌 Задача: *{task_title}*\n📅 Дата дедлайна: `{task_date}`\n\nМы напомним вам в этот день в Telegram!"
    success = send_telegram_message(chat_id, text)
    try:
        if supabase and user_id:
            supabase.table("user_tasks").insert({
                "user_id": user_id,
                "tg_id": str(chat_id),
                "title": task_title,
                "date": task_date,
                "notified": False
            }).execute()
    except Exception:
        pass

    if success:
        return {"success": True, "message": "Напоминание отправлено!"}
    return JSONResponse({"error": "Сообщение не отправлено."}, status_code=400)

# --- ИИ УЧИТЕЛЬ (ВИДЕО И ИГРА) ---
@app.post("/api/ai-teacher")
async def ai_teacher(
    user_id: int = Form(...),
    topic: str = Form(""),
    comment: str = Form(""),
    doc_id: str = Form("default"),
    lang: str = Form("kz")
):
    try:
        file_text = active_documents.get(doc_id, "")[:3000]
        lang_instr = {"kz": "Қазақ тілінде.", "ru": "На русском языке.", "en": "In English."}.get(lang, "На русском языке.")
        
        prompt = f"""{lang_instr}
Ты супер-интересный интерактивный ИИ-преподаватель. 
Тема/Задача: "{topic}"
Дополнительный комментарий: "{comment}"
Контекст текста файла: "{file_text}"

Твоя задача:
1. Объясни материал простым, увлекательным языком.
2. Дай рекомендацию по поиску полезного видео на YouTube или предложи наглядный пример.
3. Подготовь мини-тест (игру) строго из 3 интерактивных вопросов для закрепления.
Верни ответ СТРОГО в формате JSON (без лишнего текста вокруг):
{
  "explanation": "Объяснение и рекомендация видео...",
  "game_questions": [
    {
      "q": "Вопрос 1?",
      "options": ["А) Вариант", "В) Вариант", "С) Вариант"],
      "answer": "А",
      "hint": "Отличное пояснение"
    }
  ]
}"""

        response = ai_client.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "system", "content": "Ты крутой преподаватель с игровыми элементами."}, {"role": "user", "content": prompt}],
            temperature=0.4
        )
        return {"result": response.choices[0].message.content}
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

# --- БИЛЕТЫ (ИГРОВОЙ ФОРМАТ) ---
@app.post("/api/tickets")
async def generate_tickets(
    user_id: int = Form(...),
    theme: str = Form(""),
    comment: str = Form(""),
    count: int = Form(5),
    doc_id: str = Form("default"),
    lang: str = Form("kz")
):
    try:
        file_text = active_documents.get(doc_id, "")[:3000]
        lang_instr = {"kz": "Қазақ тілінде.", "ru": "На русском языке.", "en": "In English."}.get(lang, "На русском языке.")
        
        prompt = f"""{lang_instr}
Подготовь ровно {count} экзаменационных билетов по теме "{theme}".
Дополнительный комментарий: "{comment}"
Текст файла: "{file_text}"

В каждом билете должны быть вопросы и мини-задание в игровом стиле.
Верни ответ СТРОГО в виде JSON-массива:
[
  {
    "ticket_num": 1,
    "question": "Вопрос билета...",
    "game_task": "Мини-задание или вопрос-тест в игровом стиле",
    "options": ["A) ...", "B) ...", "C) ..."],
    "correct": "A"
  }
]"""

        response = ai_client.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "system", "content": "Ты строгий и справедливый профессор."}, {"role": "user", "content": prompt}],
            temperature=0.3
        )
        res_text = response.choices[0].message.content
        if supabase:
            supabase.table("saved_materials").insert({"user_id": user_id, "title": f"Билеты: {theme}", "content": res_text, "type": "tickets"}).execute()
        return {"result": res_text}
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

@app.post("/api/chat")
async def chat_with_ai(
    user_id: int = Form(...),
    message: str = Form(...),
    doc_id: str = Form("default"),
    lang: str = Form("kz")
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
Ты интеллектуальный ИИ-наставник StudyAI Hub. Отвечай кратко, качественно и структурировано.
Контекст текущего материала:
{user_text[:4000] if user_text else 'Общий учебный контекст'}

Вопрос: {message}"""

        response = ai_client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "Ты опытный эксперт-преподаватель."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.4
        )
        reply = response.choices[0].message.content
        new_tokens = tokens if is_sub else max(0, tokens - 1)
        if supabase:
            supabase.table("users").update({"tokens": new_tokens}).eq("id", user_id).execute()
            supabase.table("chat_history").insert({"user_id": user_id, "message": message, "reply": reply}).execute()

        return {"reply": reply, "tokens_left": new_tokens, "is_subscribed": is_sub}
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

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
            
            msg = f"👋 *StudyAI Hub — Авторизация*\n\nВаш код входа:\n`{code}`\n\nВведите его на сайте."
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
            return JSONResponse({"error": "Неверный или устаревший код!"}, status_code=400)
        
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
    temp_filename = f"temp_{file.filename}"
    try:
        with open(temp_filename, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        reader = PdfReader(temp_filename)
        text = "".join([page.extract_text() + "\n" for page in reader.pages if page.extract_text()])
        if os.path.exists(temp_filename): os.remove(temp_filename)
        doc_id = str(os.urandom(4).hex())
        active_documents[doc_id] = text
        return {"doc_id": doc_id, "message": "Файл успешно загружен!"}
    except Exception as e:
        if os.path.exists(temp_filename): os.remove(temp_filename)
        return JSONResponse({"error": str(e)}, status_code=500)

@app.post("/generate")
async def generate_content(
    doc_id: str = Form("default"), 
    action: str = Form(...), 
    user_id: int = Form(...), 
    lang: str = Form("kz"), 
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
                    "message": "Токены закончились!",
                    "is_subscribed": False,
                    "tokens_left": 0
                }, status_code=402)
        else:
            tokens, is_sub = 3, True
    except Exception:
        return JSONResponse({"error": "Ошибка проверки профиля"}, status_code=500)

    user_text = active_documents.get(doc_id, "")
    lang_instr = {
        "kz": "СТРОГО НА КАЗАХСКОМ ЯЗЫКЕ (Мінсіз қазақ тілінде).",
        "en": "STRICTLY IN ENGLISH.",
        "ru": "Строго на русском языке."
    }.get(lang, "Строго на русском языке.")

    start_q = (page - 1) * 20 + 1
    end_q = page * 20
    current_year = datetime.now().year # Защита 2026 года

    if action == "planner":
        tomorrow_date = (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d")
        prompt_text = f"""{lang_instr}
Ты ИИ-ассистент по тайм-менеджменту для студентов. Текущий год строго {current_year}.
Запрос пользователя: "{custom_prompt}"
Если пользователь говорит «завтра», укажи дату дедлайна: {tomorrow_date}.

Верни подробный отчет и В КОНЦЕ БЛОК JSON:
[
  {{"title": "Название задачи", "date": "{current_year}-MM-DD", "desc": "Инструкция и метод выполнения"}}
]"""
    elif action == "quiz":
        prompt_text = f"""{lang_instr}
Создай ровно 20 уникальных тестовых вопросов (номера с {start_q} по {end_q}).
Верни только валидный JSON-массив:
[
  {
    "id": 1,
    "question": "Вопрос...",
    "options": ["A) Вариант 1", "B) Вариант 2", "C) Вариант 3", "D) Вариант 4"],
    "answer": "A",
    "explanation": "Пояснение."
  }
]"""
    elif action == "flashcards":
        prompt_text = f"""{lang_instr}
Создай 10 интерактивных учебных карточек строго по шаблону:
### Карточка X
**Вопрос:** [...]
**Ответ:** [...]
---"""
    elif action == "summary":
        prompt_text = f"{lang_instr}\nСделай глубокий академический конспект с формулами ($...$) и основными терминами."
    else:
        prompt_text = f"{lang_instr}\nОбъясни этот материал простыми примерами."

    try:
        content_text = (user_text or doc_id)[:7000] if action != "planner" else custom_prompt
        response = ai_client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": f"Ты передовой ИИ-эксперт {current_year} года."},
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

@app.get("/api/user/{user_id}/saves")
async def get_user_saves(user_id: int):
    try:
        if supabase:
            saves = supabase.table("saved_materials").select("*").eq("user_id", user_id).order("id", desc=True).execute()
            chats = supabase.table("chat_history").select("*").eq("user_id", user_id).order("id", desc=True).execute()
            user = supabase.table("users").select("*").eq("id", user_id).execute().data[0]
            return {
                "username": user.get("username"), 
                "score": user.get("score", 0), 
                "tokens": user.get("tokens"), 
                "is_subscribed": user.get("is_subscribed", False),
                "saves": saves.data or [],
                "chats": chats.data or []
            }
        return {"username": "Студент", "score": 100, "tokens": 3, "is_subscribed": True, "saves": [], "chats": []}
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
