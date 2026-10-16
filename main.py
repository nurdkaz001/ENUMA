from __future__ import annotations
import os
import time
import shutil
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

active_documents = {}

@app.get("/", response_class=FileResponse)
async def read_root():
    return "templates/index.html"

@app.post("/api/register")
async def register(username: str = Form(...), password: str = Form(...)):
    try:
        response = supabase.table("users").insert({
            "username": username,
            "password": password,
            "score": 0
        }).execute()
        user = response.data[0]
        return {"success": True, "user_id": user["id"], "username": user["username"], "score": user["score"]}
    except Exception as e:
        return JSONResponse({"error": "Такое имя пользователя уже занято или ошибка базы данных!"}, status_code=400)

@app.post("/api/login")
async def login(username: str = Form(...), password: str = Form(...)):
    try:
        response = supabase.table("users").select("*").eq("username", username).eq("password", password).execute()
        if response.data and len(response.data) > 0:
            user = response.data[0]
            return {"success": True, "user_id": user["id"], "username": user["username"], "score": user["score"]}
        return JSONResponse({"error": "Неверное имя пользователя или пароль."}, status_code=400)
    except Exception as e:
        return JSONResponse({"error": f"Ошибка авторизации: {str(e)}"}, status_code=500)

# Глобальный рейтинг студентов для личного кабинета
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
async def generate_content(doc_id: str = Form(...), action: str = Form(...), user_id: int = Form(None), title: str = Form("Урок")):
    user_text = active_documents.get(doc_id)
    if not user_text:
        return JSONResponse({"error": "Сессия устарела или файл не найден. Загрузите материал заново."}, status_code=400)

    prompts = {
        "summary": "Сделай структурированный краткий конспект на основе этого текста. Выдели главные тезисы, ключевые формулы и термины.",
        "quiz": "Составь 10 качественных тестовых вопросов с вариантами ответов (A, B, C, D) и укажи правильные ответы в самом конце.",
        "flashcards": "Создай 8 учебных карточек (Вопрос — Ответ) для заучивания ключевых понятий из текста.",
        "simple": "Объясни сложные идеи из этого текста максимально простыми словами, используя понятные аналогии.",
        "plan": "Составь эффективный пошаговый план подготовки и изучения этого материала."
    }

    prompt_text = prompts.get(action, "Сделай подробную выжимку из текста.")
    optimized_text = user_text[:7000]

    max_retries = 3
    result_text = ""
    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(
                model='gemini-1.5-flash',
                contents=f"{prompt_text}\n\nТекст документа:\n{optimized_text}"
            )
            result_text = response.text
            break
        except Exception as e:
            if "503" in str(e) and attempt < max_retries - 1:
                time.sleep(2)
                continue
            return JSONResponse({"error": f"Ошибка генерации ИИ: {str(e)}"}, status_code=500)

    if user_id and user_id > 0:
        try:
            supabase.table("saved_materials").insert({
                "user_id": user_id,
                "title": title,
                "content": result_text,
                "type": action
            }).execute()

            user_res = supabase.table("users").select("score").eq("id", user_id).execute()
            current_score = user_res.data[0]["score"] if user_res.data else 0
            new_score = current_score + 20

            supabase.table("users").update({"score": new_score}).eq("id", user_id).execute()
            return {"result": result_text, "new_score": new_score}
        except Exception as db_err:
            print("DB Save Error:", db_err)

    return {"result": result_text, "new_score": None}

@app.get("/api/user/{user_id}/saves")
async def get_user_saves(user_id: int):
    try:
        saves_res = supabase.table("saved_materials").select("*").eq("user_id", user_id).order("id", desc=True).execute()
        user_res = supabase.table("users").select("score, username").eq("id", user_id).execute()
        
        user_info = user_res.data[0] if user_res.data else {"username": "", "score": 0}
        saves = saves_res.data if saves_res.data else []

        return {
            "username": user_info["username"],
            "score": user_info["score"],
            "saves": saves
        }
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
