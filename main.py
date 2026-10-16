from __future__ import annotations
import os
import time
import random
import shutil
import urllib.request
import urllib.parse
from fastapi import FastAPI, File, Form, UploadFile
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

active_documents = {}

def send_telegram_code(phone: str, code: str):
    """Отправка OTP-кода (в продакшене через SMS или Telegram по привязке номера)"""
    print(f"[{phone}] Код авторизации: {code}")
    if not TELEGRAM_BOT_TOKEN:
        return True
    return True

@app.get("/", response_class=FileResponse)
async def read_root():
    return "templates/index.html"

# 1. Запрос кода по номеру телефона
@app.post("/api/auth/send-code")
async def send_code(phone: str = Form(...)):
    phone_clean = ''.join(filter(str.isdigit, phone))
    if len(phone_clean) < 10:
        return JSONResponse({"error": "Введите корректный номер телефона!"}, status_code=400)
    
    code = str(random.randint(1000, 9999))  # 4-значный удобный код
    
    try:
        # Удаляем старые коды для этого номера и записываем новый
        supabase.table("auth_codes").delete().eq("tg_id", phone_clean).execute()
        supabase.table("auth_codes").insert({"tg_id": phone_clean, "code": code}).execute()
        
        send_telegram_code(phone_clean, code)
        return {"success": True, "message": f"Код отправлен! (Для тестов код: {code})"}
    except Exception as e:
        return JSONResponse({"error": f"Ошибка отправки: {str(e)}"}, status_code=500)

# 2. Подтверждение кода по номеру телефона
@app.post("/api/auth/verify-code")
async def verify_code(phone: str = Form(...), code: str = Form(...)):
    phone_clean = ''.join(filter(str.isdigit, phone))
    try:
        res = supabase.table("auth_codes").select("*").eq("tg_id", phone_clean).eq("code", code.strip()).execute()
        
        if not res.data or len(res.data) == 0:
            return JSONResponse({"error": "Неверный код авторизации!"}, status_code=400)
        
        # Проверяем / создаем пользователя по номеру
        user_res = supabase.table("users").select("*").eq("username", phone_clean).execute()
        
        if user_res.data and len(user_res.data) > 0:
            user = user_res.data[0]
        else:
            new_user = supabase.table("users").insert({
                "username": phone_clean,
                "password": "phone_authenticated",
                "score": 0,
                "tokens": 3,
                "is_subscribed": False
            }).execute()
            user = new_user.data[0]

        return {
            "success": True,
            "user_id": user["id"],
            "username": user["username"],
            "score": user.get("score", 0),
            "tokens": user.get("tokens", 3),
            "is_subscribed": user.get("is_subscribed", False)
        }
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

@app.get("/api/leaderboard")
async def get_leaderboard():
    try:
        res = supabase.table("users").select("username, score").order("score", desc=True).limit(10).execute()
        return {"leaderboard": res.data if res.data else []}
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

@app.post("/upload")
async def upload_pdf(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        return JSONResponse({"error": "Пожалуйста, загрузите файл в формате PDF."}, status_code=400)
    
    temp_filename = f"temp_{file.filename}"
    try:
        with open(temp_filename, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        reader = PdfReader(temp_filename)
        text = ""
        for page in reader.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"
                
        if os.path.exists(temp_filename):
            os.remove(temp_filename)

        if len(text.strip()) < 30:
            return JSONResponse({"error": "В PDF-файле не найден текст."}, status_code=400)

        doc_id = str(os.urandom(4).hex())
        active_documents[doc_id] = text
        return {"doc_id": doc_id, "message": "Файл успешно обработан!"}
    except Exception as e:
        if os.path.exists(temp_filename):
            os.remove(temp_filename)
        return JSONResponse({"error": f"Ошибка обработки файла: {str(e)}"}, status_code=500)

@app.post("/generate")
async def generate_content(
    doc_id: str = Form(...), 
    action: str = Form(...), 
    user_id: int = Form(...), 
    lang: str = Form("ru"),
    title: str = Form("Урок")
):
    try:
        u_res = supabase.table("users").select("*").eq("id", user_id).execute()
        if not u_res.data:
            return JSONResponse({"error": "Пользователь не найден. Авторизуйтесь!"}, status_code=401)
        
        user = u_res.data[0]
        tokens = user.get("tokens", 0)
        is_subscribed = user.get("is_subscribed", False)

        if not is_subscribed and tokens <= 0:
            return JSONResponse({
                "error": "PAYMENT_REQUIRED",
                "message": "Бесплатные токены закончились! Оформите подписку за 500 ₸ для безлимитного доступа."
            }, status_code=402)
    except Exception as e:
        return JSONResponse({"error": "Ошибка проверки лимитов"}, status_code=500)

    user_text = active_documents.get(doc_id)
    if not user_text and not doc_id.startswith("topic_"):
        return JSONResponse({"error": "Сессия устарела. Загрузите материал заново."}, status_code=400)

    lang_instructions = {
        "kz": "Жауапты СТРОГО ҚАЗАҚ ТІЛІНДЕ дайында.",
        "en": "Provide the response STRICTLY IN ENGLISH.",
        "ru": "Предоставь ответ СТРОГО НА РУССКОМ ЯЗЫКЕ."
    }
    selected_lang = lang_instructions.get(lang, lang_instructions["ru"])

    prompts = {
        "summary": f"{selected_lang} Сделай структурированный краткий конспект на основе этого текста. Выдели главные тезисы, ключевые формулы и термины.",
        "quiz": f"{selected_lang} Составь 10 качественных тестовых вопросов с вариантами ответов (A, B, C, D) и укажи правильные ответы в самом конце.",
        "flashcards": f"{selected_lang} Создай 8 учебных карточек (Вопрос — Ответ). Строго оформляй каждой карточке в формате:\n### Карточка X\n**Вопрос:** ...\n**Ответ:** ...\n---",
        "simple": f"{selected_lang} Объясни сложные идеи из этого текста максимально простыми словами, используя понятные аналогии.",
        "plan": f"{selected_lang} Составь эффективный пошаговый план подготовки и изучения этого материала."
    }

    prompt_text = prompts.get(action, "Сделай подробную выжимку.")
    optimized_text = (user_text or doc_id)[:7000]

    try:
        response = client.models.generate_content(
            model='gemini-3.8-flash',
            contents=f"{prompt_text}\n\nТекст документа:\n{optimized_text}"
        )
        result_text = response.text
        
        new_tokens = tokens if is_subscribed else max(0, tokens - 1)
        new_score = user.get("score", 0) + 20

        supabase.table("users").update({
            "tokens": new_tokens,
            "score": new_score
        }).eq("id", user_id).execute()

        supabase.table("saved_materials").insert({
            "user_id": user_id,
            "title": title,
            "content": result_text,
            "type": action
        }).execute()

        return {
            "result": result_text,
            "new_score": new_score,
            "tokens_left": new_tokens,
            "is_subscribed": is_subscribed
        }
    except Exception as e:
        return JSONResponse({"error": f"Ошибка ИИ: {str(e)}"}, status_code=500)

@app.post("/api/user/subscribe")
async def subscribe_user(user_id: int = Form(...)):
    try:
        supabase.table("users").update({
            "is_subscribed": True
        }).eq("id", user_id).execute()
        return {"success": True, "message": "Подписка на 1 месяц успешно активирована!"}
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

@app.get("/api/user/{user_id}/saves")
async def get_user_saves(user_id: int):
    try:
        saves_res = supabase.table("saved_materials").select("*").eq("user_id", user_id).order("id", desc=True).execute()
        user_res = supabase.table("users").select("*").eq("id", user_id).execute()
        
        user_info = user_res.data[0] if user_res.data else {}
        saves = saves_res.data if saves_res.data else []

        return {
            "username": user_info.get("username", ""),
            "score": user_info.get("score", 0),
            "tokens": user_info.get("tokens", 0),
            "is_subscribed": user_info.get("is_subscribed", False),
            "saves": saves
        }
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
