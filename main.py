from __future__ import annotations
import os
import shutil
from fastapi import FastAPI, File, Form, UploadFile, Request
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from google import genai
from pypdf import PdfReader

# Инициализация FastAPI приложения
app = FastAPI(title="EdTech AI Platform for Students")

# Настройка ключа и клиента Gemini
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
client = genai.Client(api_key=GEMINI_API_KEY)

# Настройка шаблонов (папка templates в корне проекта)
templates = Jinja2Templates(directory="templates")

# Временное хранилище текстов обработанных PDF
active_documents = {}

@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    # Безопасный рендеринг через окружение шаблонов Jinja2 напрямую
    template = templates.get_template("index.html")
    html_content = template.render({"request": request})
    return HTMLResponse(content=html_content)

@app.post("/upload")
async def upload_pdf(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        return JSONResponse({"error": "Пожалуйста, загрузите файл в формате PDF."}, status_code=400)
    
    temp_filename = f"temp_{file.filename}"
    try:
        with open(temp_filename, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Чтение текста из PDF
        reader = PdfReader(temp_filename)
        text = ""
        for page in reader.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"
                
        if os.path.exists(temp_filename):
            os.remove(temp_filename)

        if len(text.strip()) < 30:
            return JSONResponse({"error": "В PDF-файле не найден текст. Возможно, это отсканированная картинка."}, status_code=400)

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
        return JSONResponse({"error": "Сессия устарела или файл не найден. Загрузите PDF заново."}, status_code=400)

    prompts = {
        "summary": "Сделай структурированный краткий конспект на основе этого текста. Выдели главные тезисы, ключевые формулы и термины.",
        "quiz": "Составь 5 качественных тестовых вопросов с вариантами ответов (A, B, C, D) и укажи правильные ответы в самом конце.",
        "flashcards": "Создай 6 учебных карточек (Вопрос — Ответ) для заучивания ключевых понятий из текста.",
        "simple": "Объясни сложные идеи из этого текста максимально простыми словами, используя понятные аналогии.",
        "plan": "Составь эффективный пошаговый план подготовки и изучения этого материала."
    }

    prompt_text = prompts.get(action, "Сделай подробную выжимку из текста.")
    optimized_text = user_text[:7000] # Ограничение для стабильной работы модели

    try:
        # Используем актуальную модель Gemini
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=f"{prompt_text}\n\nТекст документа:\n{optimized_text}"
        )
        return {"result": response.text}
    except Exception as e:
        return JSONResponse({"error": f"Ошибка генерации ИИ: {str(e)}"}, status_code=500)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
