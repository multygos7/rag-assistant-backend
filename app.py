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
import requests
import io
from pypdf import PdfReader
from docx import Document as DocxDocument

# ==================== НАСТРОЙКИ ====================
SUPABASE_URL = "https://jidtwjamnglkqoqizvjl.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImppZHR3amFtbmdsa3FvcWl6dmpsIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc2OTI4NDg3MSwiZXhwIjoyMDg0ODYwODcxfQ.3Yl9LIJ7BnRmAAYrzc0NiPEilphXQm8EdObmh2_y5BU"
GIGACHAT_CREDENTIALS = "MDE5ZjQ3MDgtODM2Zi03ZGU3LWJlNTMtZTQzMTI3MjE3NDVhOjkyMTU5ZDkyLWQzN2YtNDM4Ny05YTRhLWQxNjQxYmE4ZTI0ZQ=="
OPENROUTER_API_KEY = "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"
QWEN_MODEL = "qwen/qwen3-next-80b-a3b-instruct:free"

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

MAX_FILE_TEXT_CHARS = 12000  # ограничиваем объём текста из файла, чтобы не раздувать промпт


def extract_text_from_file(file_storage):
    """Извлекаем текст из загруженного файла: txt, pdf или docx."""
    filename = file_storage.filename.lower()
    content = file_storage.read()

    if filename.endswith(".txt"):
        text = content.decode("utf-8", errors="ignore")
    elif filename.endswith(".pdf"):
        reader = PdfReader(io.BytesIO(content))
        text = "\n".join(page.extract_text() or "" for page in reader.pages)
    elif filename.endswith(".docx"):
        doc = DocxDocument(io.BytesIO(content))
        text = "\n".join(p.text for p in doc.paragraphs)
    else:
        raise ValueError("Поддерживаются только файлы .txt, .pdf, .docx")

    text = text.strip()
    if len(text) > MAX_FILE_TEXT_CHARS:
        text = text[:MAX_FILE_TEXT_CHARS] + "\n[...текст обрезан, файл слишком большой...]"
    return text


def get_giga():
    """Создаём новое подключение к GigaChat на каждый запрос (токен живёт 30 минут)."""
    return GigaChat(credentials=GIGACHAT_CREDENTIALS, verify_ssl_certs=False)


def ask_qwen(prompt, retries=1):
    """Отправляем запрос в Qwen через OpenRouter. Если конкретно Qwen сейчас
    перегружен (429) — переключаемся на openrouter/free, который сам подберёт
    доступную бесплатную модель. Возвращаем текст + реальное название модели-ответчика."""
    import time

    def call(model_id):
        return requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers={"Authorization": f"Bearer {OPENROUTER_API_KEY}"},
            json={"model": model_id, "messages": [{"role": "user", "content": prompt}]},
            timeout=30
        )

    last_error = None
    for attempt in range(retries + 1):
        response = call(QWEN_MODEL)
        data = response.json()
        if "choices" in data:
            return data["choices"][0]["message"]["content"], data.get("model", QWEN_MODEL)
        last_error = data
        if response.status_code == 429 and attempt < retries:
            time.sleep(3)
            continue
        break

    # Qwen недоступен — пробуем автоматический подбор среди других бесплатных моделей
    response = call("openrouter/free")
    data = response.json()
    if "choices" in data:
        return data["choices"][0]["message"]["content"], data.get("model", "openrouter/free")

    raise Exception(f"OpenRouter error: {last_error}")


def ask_llm(giga, prompt, model_choice):
    """Единая обёртка — отправляет промпт в выбранную модель (gigachat или qwen).
    Возвращает (текст_ответа, реальное_название_модели_ответившей)."""
    if model_choice == "qwen":
        return ask_qwen(prompt)
    chat = Chat(messages=[Messages(role=MessagesRole.USER, content=prompt)])
    response = giga.chat(chat)
    return response.choices[0].message.content, "GigaChat"


def expand_query_with_llm(giga, question):
    """Просим GigaChat подсказать связанные термины к вопросу — чтобы находить
    документы даже если пользователь не использует те же слова что в базе знаний."""
    prompt = f"""Пользователь задал вопрос сотруднику компании: "{question}"

Твоя задача — предложить 5-8 ключевых слов и связанных терминов, которые могут встречаться в документах базы знаний компании и помогут найти ответ на этот вопрос. Учитывай синонимы, смежные термины, официальные названия процессов/модулей.

Ответь ТОЛЬКО списком слов через запятую, без пояснений. Например: отпуск, заявление, оформление, СБИС, кадры"""

    chat = Chat(messages=[Messages(role=MessagesRole.USER, content=prompt)])
    response = giga.chat(chat)
    keywords = response.choices[0].message.content.strip()
    return keywords


def rerank_with_llm(giga, question, candidates, top_n=3):
    """Просим LLM выбрать самые релевантные документы из широкого пула кандидатов
    по смыслу вопроса, а не по формальному совпадению слов."""
    candidates_list = "\n".join([
        f"{i+1}. {c['title']}: {c['content'][:200]}"
        for i, c in enumerate(candidates)
    ])

    prompt = f"""Вопрос сотрудника: "{question}"

Вот список документов из базы знаний (номер, название, начало текста):

{candidates_list}

Выбери номера документов (от 1 до {len(candidates)}), которые реально помогут ответить на этот вопрос по смыслу — даже если в них не совпадают точные слова из вопроса. Выбери не более {top_n} самых релевантных. Если ни один документ не отвечает на вопрос по существу — ответь "нет".

Ответь ТОЛЬКО номерами через запятую (например: 3, 7, 1) или словом "нет"."""

    chat = Chat(messages=[Messages(role=MessagesRole.USER, content=prompt)])
    response = giga.chat(chat)
    raw_answer = response.choices[0].message.content.strip()
    answer = raw_answer.lower()

    if "нет" in answer and len(answer) < 10:
        return [], raw_answer

    import re as _re
    numbers = [int(n) for n in _re.findall(r'\d+', answer)]
    selected = []
    for n in numbers:
        if 1 <= n <= len(candidates):
            selected.append(candidates[n - 1])
    return selected[:top_n], raw_answer


def search_documents(exact_query, fuzzy_query, exact_limit=60, fuzzy_limit=15):
    """Ищем релевантные документы: точный поиск глубже (не теряет специфичные темы),
    нечёткий поиск неглубоко (страхует только от опечаток)."""
    result = supabase.rpc("search_kb_documents", {
        "exact_query_text": exact_query,
        "fuzzy_query_text": fuzzy_query,
        "exact_limit": exact_limit,
        "fuzzy_limit": fuzzy_limit
    }).execute()
    return result.data


def build_prompt(question, documents, history=None, file_text=None):
    context = "\n\n".join([
        f"### {doc['title']}\n{doc['content'][:2000]}"
        for doc in documents
    ])

    history_block = ""
    if history:
        turns = "\n".join([f"{'Сотрудник' if h['role']=='user' else 'Ассистент'}: {h['content']}" for h in history[-6:]])
        history_block = f"\nПредыдущая часть разговора (для контекста, не повторяй это в ответе):\n{turns}\n"

    file_block = ""
    if file_text:
        file_block = f"\nПользователь также приложил файл со следующим содержимым — используй его при ответе если вопрос об этом файле:\n{file_text}\n"

    return f"""Ты — ИИ-ассистент компании Райтек. Отвечай на вопросы сотрудников на основе предоставленного контекста из базы знаний.

Документы ниже уже отобраны отдельным алгоритмом как релевантные вопросу — не сомневайся в их применимости и не пиши "точного ответа не нашлось", если в контексте есть релевантная информация, даже частичная. Отвечай уверенно на основе того что есть в контексте.

Если контекст совсем не по теме вопроса (пустой по смыслу для этого конкретного вопроса) — тогда прямо скажи: "К сожалению, точного ответа в базе знаний не нашлось".

Не отвечай на вопросы не связанные с работой компании (погода, курсы валют, рецепты и т.д.) — вежливо откажи, сославшись на то что ты отвечаешь только по базе знаний компании.

Если вопрос касается чисел, времени, дат, диапазонов ("можно ли в такое-то время", "хватит ли столько дней" и т.п.) — рассуждай пошагово вслух перед финальным ответом: выпиши точные цифры/границы из контекста, сравни их с тем что спрашивает пользователь, и только потом дай вывод. Не торопись с ответом "да/нет" не проверив числа.

Пиши обычным текстом без markdown-разметки — не используй звёздочки для выделения жирным (**слово**), не используй решётки для заголовков. Обычные предложения и списки через тире.
{history_block}{file_block}
Контекст из базы знаний:
{context}

Вопрос пользователя: {question}

Дай чёткий, содержательный ответ по существу на основе контекста."""


def answer_question(question, model_choice="gigachat", history=None, file_text=None):
    giga = get_giga()
    debug = {"question": question, "model_choice": model_choice}

    # Шаг 1: расширяем вопрос — синонимы + LLM-подсказки
    expanded = expand_query(question)
    llm_keywords = expand_query_with_llm(giga, question)
    search_query = expanded + " " + llm_keywords
    debug["synonym_expansion"] = expanded
    debug["llm_keywords"] = llm_keywords

    # Шаг 2: ищем — точный поиск по расширенному запросу, нечёткий по короткому исходному вопросу
    candidates = search_documents(exact_query=search_query, fuzzy_query=question)
    debug["candidates"] = [
        {"title": c["title"], "rank": round(c["rank"], 4), "match_type": c.get("match_type", "?")}
        for c in candidates
    ]

    # Если ничего не нашли в базе знаний И файл не приложен — честно сообщаем что не нашли
    if not candidates and not file_text:
        debug["result"] = "off_topic_no_candidates"
        return {
            "answer": "К сожалению, точного ответа в базе знаний не нашлось. Могу создать заявку по вашему вопросу для специалиста.",
            "sources": [], "off_topic": True, "debug": debug, "model_used": model_choice
        }

    documents = []
    reranker_raw = None
    if candidates:
        # Шаг 3: reranker — LLM выбирает реально релевантные документы из пула
        documents, reranker_raw = rerank_with_llm(giga, question, candidates, top_n=3)
    debug["reranker_selected"] = [d["title"] for d in documents]
    debug["reranker_raw_response"] = reranker_raw

    # Если reranker ничего не выбрал И файла нет — тоже "не нашлось"
    if not documents and not file_text:
        debug["result"] = "off_topic_reranker_empty"
        return {
            "answer": "К сожалению, точного ответа в базе знаний не нашлось. Могу создать заявку по вашему вопросу для специалиста.",
            "sources": [], "off_topic": True, "debug": debug, "model_used": model_choice
        }

    prompt = build_prompt(question, documents, history=history, file_text=file_text)
    debug["final_prompt_preview"] = prompt[:1500]

    # Шаг 4: финальный ответ — генерируем той моделью, которую выбрал пользователь
    answer, actual_model = ask_llm(giga, prompt, model_choice)
    answer = answer.replace("**", "").replace("##", "").replace("###", "")
    debug["result"] = "answered"
    debug["actual_model"] = actual_model

    sources = [{"title": doc["title"], "url": doc.get("source_url")} for doc in documents]

    try:
        supabase.table("chat_logs").insert({
            "user_name": "web_user",
            "question": question,
            "answer": answer,
            "is_off_topic": False,
            "model_used": actual_model,
        }).execute()
    except Exception:
        pass

    return {"answer": answer, "sources": sources, "off_topic": False, "debug": debug, "model_used": actual_model}


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/ask", methods=["POST"])
def ask():
    data = request.get_json()
    question = (data or {}).get("question", "").strip()
    model_choice = (data or {}).get("model", "gigachat")
    history = (data or {}).get("history", [])
    file_text = (data or {}).get("file_text")

    if not question:
        return jsonify({"error": "Пустой вопрос"}), 400

    try:
        result = answer_question(question, model_choice, history=history, file_text=file_text)
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "Файл не передан"}), 400

    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "Пустое имя файла"}), 400

    try:
        text = extract_text_from_file(file)
        if not text:
            return jsonify({"error": "Не удалось извлечь текст из файла"}), 400
        return jsonify({"filename": file.filename, "text": text, "chars": len(text)})
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"Ошибка обработки файла: {e}"}), 500


if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5001))
    app.run(host="0.0.0.0", port=port)
