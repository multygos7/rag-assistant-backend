"""
Flask-сервер для RAG-ассистента Райтек.
Отдаёт веб-интерфейс + API endpoint для вопросов.

Установка зависимостей:
    pip3 install flask supabase gigachat --break-system-packages

Запуск локально:
    python3 app.py
Потом открыть в браузере: http://localhost:5000
"""

from flask import Flask, request, jsonify, render_template
from supabase import create_client
from gigachat import GigaChat
from gigachat.models import Chat, Messages, MessagesRole

# ==================== НАСТРОЙКИ ====================
SUPABASE_URL = "https://jidtwjamnglkqoqizvjl.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImppZHR3amFtbmdsa3FvcWl6dmpsIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc2OTI4NDg3MSwiZXhwIjoyMDg0ODYwODcxfQ.3Yl9LIJ7BnRmAAYrzc0NiPEilphXQm8EdObmh2_y5BU"
GIGACHAT_CREDENTIALS = "MDE5ZjQ3MDgtODM2Zi03ZGU3LWJlNTMtZTQzMTI3MjE3NDVhOjkyMTU5ZDkyLWQzN2YtNDM4Ny05YTRhLWQxNjQxYmE4ZTI0ZQ=="

MIN_RANK = 0.01

# Словарь синонимов для типичных тем — расширяем запрос словами той же темы,
# чтобы разные формы слова (глагол/существительное) находили нужный документ
SYNONYM_GROUPS = [
    ["увольнение", "уволиться", "уволить", "увольняюсь"],
    ["отпуск", "отдохнуть", "отгул"],
    ["больничный", "заболел", "болею", "нетрудоспособность"],
    ["командировка", "командировку", "поехать"],
    ["сертификация", "сертификат", "аттестация", "экзамен"],
    ["дмс", "бенефиты", "страховка"],
    ["зарплата", "оплата труда", "выплата", "аванс"],
    ["трудоустройство", "прием на работу", "оформление"],
]


def expand_query(question):
    """Добавляем синонимы известных тем к запросу, если они там встречаются."""
    lower_q = question.lower()
    extra_words = []
    for group in SYNONYM_GROUPS:
        if any(word in lower_q for word in group):
            extra_words.extend(group)
    if extra_words:
        return question + " " + " ".join(set(extra_words))
    return question
# =====================================================

app = Flask(__name__)
supabase = create_client(SUPABASE_URL, SUPABASE_KEY)


def get_giga():
    """Создаём новое подключение к GigaChat на каждый запрос (токен живёт 30 минут)."""
    return GigaChat(credentials=GIGACHAT_CREDENTIALS, verify_ssl_certs=False)


def search_documents(question, match_count=3):
    result = supabase.rpc("search_kb_documents", {
        "query_text": question,
        "match_count": match_count
    }).execute()
    return result.data


def build_prompt(question, documents):
    context = "\n\n".join([
        f"### {doc['title']}\n{doc['content'][:2000]}"
        for doc in documents
    ])
    return f"""Ты — ИИ-ассистент компании Райтек. Отвечай на вопросы сотрудников ТОЛЬКО на основе предоставленного контекста из базы знаний.

Если в контексте нет ответа на вопрос — прямо скажи: "К сожалению, точного ответа в базе знаний не нашлось" и не выдумывай ответ.

Не отвечай на вопросы не связанные с работой компании (погода, курсы валют, рецепты и т.д.) — вежливо откажи, сославшись на то что ты отвечаешь только по базе знаний компании.

Контекст из базы знаний:
{context}

Вопрос пользователя: {question}

Дай чёткий, короткий ответ по существу."""


def answer_question(question):
    expanded = expand_query(question)
    documents = search_documents(expanded)

    if not documents or documents[0]["rank"] < MIN_RANK:
        return {
            "answer": "К сожалению, точного ответа в базе знаний не нашлось. Могу создать заявку по вашему вопросу для специалиста.",
            "sources": [],
            "off_topic": True
        }

    prompt = build_prompt(question, documents)
    giga = get_giga()
    chat = Chat(messages=[Messages(role=MessagesRole.USER, content=prompt)])
    response = giga.chat(chat)
    answer = response.choices[0].message.content

    sources = [doc["title"] for doc in documents if doc["rank"] >= MIN_RANK]

    # Логируем диалог
    try:
        supabase.table("chat_logs").insert({
            "user_name": "web_user",
            "question": question,
            "answer": answer,
            "is_off_topic": False,
        }).execute()
    except Exception:
        pass  # логирование не критично, не должно ронять ответ

    return {"answer": answer, "sources": sources, "off_topic": False}


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/ask", methods=["POST"])
def ask():
    data = request.get_json()
    question = (data or {}).get("question", "").strip()

    if not question:
        return jsonify({"error": "Пустой вопрос"}), 400

    try:
        result = answer_question(question)
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5001))
    app.run(host="0.0.0.0", port=port)
