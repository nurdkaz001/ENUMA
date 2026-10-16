from __future__ import annotations
import os
import time
import random
import shutil
import urllib.request
import urllib.parse
import json
from fastapi import FastAPI, File, Form, UploadFile, Request
from fastapi.responses import FileResponse, JSONResponse
from google import genai
from pypdf import PdfReader
from supabase import create_client, Client

app = FastAPI(title="StudyAI Platform")

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
client = genai.Client(api_key=GEMINI_API_KEY)

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")
BOT_USERNAME = "HUB_preza_bot"  # Укажите юзернейм вашего бота без @

active_documents = {}

def send_telegram_message(chat_id: str, text: str):
    if not TELEGRAM_BOT_TOKEN:
        print(f"⚠️ Telegram Token не задан. Сообщение для {chat_id}: {text}")
        return False
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    data = urllib.parse.urlencode({"chat_id": chat_id, "text": text, "parse_mode": "Markdown"}).encode("utf-8")
    try:
        req = urllib.request.Request(url, data=data)
        with urllib.request.urlopen(req) as response:
            return response.status == 200
    except Exception as e:
        print("Telegram Send Error:", e)
        return False

@app.get("/", response_class=FileResponse)
async def read_root():
    return "templates/index.html"

# Эндпоинт, который принимает вебхук от Telegram бота (когда пользователь нажимает /start)
@app.post("/api/telegram/webhook")
async def telegram_webhook(request: Request):
    try:
        body = await request.json()
        message = body.get("message", {})
        chat = message.get("chat", {})
        chat_id = chat.get("id")
        text = message.get("text", "")
        
        if text.startswith("/start"):
            # Генерируем 6-значный код
            code = str(random.randint(100000, 999999))
            str_chat_id = str(chat_id)
            
            # Сохраняем в Supabase код, привязанный к chat_id
            supabase.table("auth_codes").delete().eq("tg_id", str_chat_id).execute()
            supabase.table("auth_codes").insert({"tg_id": str_chat_id, "code": code}).execute()
            
            # Отправляем ответ в Telegram
            msg = f"👋 *Добро пожаловать в StudyAI Hub!*\n\nВаш код подтверждения для входа:\n`{code}`\n\nВведите этот код на сайте."
            send_telegram_message(str_chat_id, msg)
            
        return {"status": "ok"}
    except Exception as e:
        return {"status": "error", "message": str(e)}

# Проверка введенного 6-значного кода
@app.post("/api/auth/verify-code")
async def verify_code(code: str = Form(...)):
    code_clean = code.strip()
    try:
        res = supabase.table("auth_codes").select("*").eq("code", code_clean).execute()
        
        if not res.data or len(res.data) == 0:
            return JSONResponse({"error": "Неверный или устаревший код!"}, status_code=400)
        
        record = res.data[0]
        tg_id = record["tg_id"]
        
        # Удаляем использованный код
        supabase.table("auth_codes").delete().eq("code", code_clean).execute()
        
        # Проверяем или создаем пользователя
        user_res = supabase.table("users").select("*").eq("username", f"tg_{tg_id}").execute()
        
        if user_res.data and len(user_res.data) > 0:
            user = user_res.data[0]
        else:
            new_user = supabase.table("users").insert({
                "username": f"tg_{tg_id}",
                "password": "telegram_auth",
                "score": 0,
                "tokens": 3,
                "is_subscribed": False
            }).execute()
            user = new_user.data[0]

        return {
            "success": True,
            "user_id": user["id"],
            "username": f"User #{tg_id[-4:]}",
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
async def generate_content(doc_id: str = Form(...), action: str = Form(...), user_id: int = Form(...), lang: str = Form("ru"), title: str = Form("Урок")):
    try:
        u_res = supabase.table("users").select("*").eq("id", user_id).execute()
        if not u_res.data: return JSONResponse({"error": "Пользователь не найден."}, status_code=401)
        user = u_res.data[0]
        tokens, is_sub = user.get("tokens", 0), user.get("is_subscribed", False)
        if not is_sub and tokens <= 0:
            return JSONResponse({"error": "PAYMENT_REQUIRED", "message": "Токены закончились! Оформите подписку за 500 ₸."}, status_code=402)
    except Exception:
        return JSONResponse({"error": "Ошибка лимитов"}, status_code=500)

    user_text = active_documents.get(doc_id)
    if not user_text and not doc_id.startswith("topic_"):
        return JSONResponse({"error": "Сессия устарела. Загрузите файл заново."}, status_code=400)

    lang_instr = {"kz": "Мінсіз қазақ тілінде жауап бер.", "en": "Answer strictly in English.", "ru": "Отвечай строго на русском."}.get(lang, "Отвечай на русском.")
    prompts = {
        "summary": f"{lang_instr} Сделай структурированный краткий конспект с формулами и терминами.",
        "quiz": f"{lang_instr} Составь 10 тестовых вопросов с вариантами (A, B, C, D) и ответами в конце.",
        "flashcards": f"{lang_instr} Создай 8 учебных карточек в формате:\n### Карточка X\n**Вопрос:** ...\n**Ответ:** ...\n---",
        "simple": f"{lang_instr} Объясни материал простыми словами."
    }
    
    try:
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=f"{prompts.get(action, 'Сделай выжимку')}\n\nТекст:\n{(user_text or doc_id)[:7000]}"
        )
        res_text = response.text
        new_tokens = tokens if is_sub else max(0, tokens - 1)
        supabase.table("users").update({"tokens": new_tokens}).eq("id", user_id).execute()
        supabase.table("saved_materials").insert({"user_id": user_id, "title": title, "content": res_text, "type": action}).execute()
        return {"result": res_text, "tokens_left": new_tokens, "is_subscribed": is_sub}
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

@app.post("/api/user/subscribe")
async def subscribe_user(user_id: int = Form(...)):
    try:
        supabase.table("users").update({"is_subscribed": True}).eq("id", user_id).execute()
        return {"success": True}
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

@app.get("/api/user/{user_id}/saves")
async def get_user_saves(user_id: int):
    try:
        saves = supabase.table("saved_materials").select("*").eq("user_id", user_id).order("id", desc=True).execute()
        user = supabase.table("users").select("*").eq("id", user_id).execute().data[0]
        return {"username": user.get("username"), "tokens": user.get("tokens"), "is_subscribed": user.get("is_subscribed"), "saves": saves.data or []}
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
