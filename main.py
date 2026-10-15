from __future__ import annotations
import os
import time
import shutil
from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from google import genai
from pypdf import PdfReader

# Инициализация приложения
app = FastAPI(title="StudyAI Platform")

# Настройка клиента Gemini
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
client = genai.Client(api_key=GEMINI_API_KEY)

# Временное хранилище документов
active_documents = {}

@app.get("/", response_class=FileResponse)
async def read_root():
    return "templates/index.html"

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
async def generate_content(doc_id: str = Form(...), action: str = Form(...)):
    user_text = active_documents.get(doc_id)
    if not user_text:
        return JSONResponse({"error": "Сессия устарела или файл не найден. Загрузите материал заново."}, status_code=400)

    prompts = {
        "summary": "Сделай структурированный краткий конспект на основе этого текста. Выдели главные тезисы, ключевые формулы и термины.",
        "quiz": "Составь 5 качественных тестовых вопросов с вариантами ответов (A, B, C, D) и укажи правильные ответы в самом конце.",
        "flashcards": "Создай 6 учебных карточек (Вопрос — Ответ) для заучивания ключевых понятий из текста.",
        "simple": "Объясни сложные идеи из этого текста максимально простыми словами, используя понятные аналогии.",
        "plan": "Составь эффективный пошаговый план подготовки и изучения этого материала."
    }

    prompt_text = prompts.get(action, "Сделай подробную выжимку из текста.")
    optimized_text = user_text[:7000]

    # Автоматический повтор запроса при высокой нагрузке (ошибка 503)
    max_retries = 3
    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(
                model='gemini-3.8-flash',
                contents=f"{prompt_text}\n\nТекст документа:\n{optimized_text}"
            )
            return {"result": response.text}
        except Exception as e:
            if "503" in str(e) and attempt < max_retries - 1:
                time.sleep(2)  # Пауза перед повторной попыткой
                continue
            return JSONResponse({"error": f"Ошибка генерации ИИ: {str(e)}"}, status_code=500)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
