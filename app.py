"""
Flask-сервер для RAG-ассистента Райтек.
Отдаёт веб-интерфейс + API endpoint для вопросов.

Установка зависимостей:
    pip3 install flask supabase gigachat --break-system-packages

Запуск локально:
    python3 app.py
Потом открыть в браузере: http://localhost:5000
"""

from flask import Flask, request, jsonify, render_template, session
from functools import wraps
from werkzeug.security import check_password_hash
from supabase import create_client
from gigachat import GigaChat
from gigachat.models import Chat, Messages, MessagesRole
import requests
import io
import re
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
app.secret_key = "raytec-rag-mvp-secret-key-change-in-prod-2026"
supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

MAX_FILE_TEXT_CHARS = 12000  # ограничиваем объём текста из файла, чтобы не раздувать промпт

# Список для обезличивания ФИО "на лету" в пользовательском вводе (перед отправкой во внешний LLM)
_PII_FIRST_NAMES = """
Александр Алексей Анатолий Андрей Антон Аркадий Артем Артём Борис Вадим Валентин Валерий
Василий Виктор Виталий Владимир Владислав Вячеслав Геннадий Георгий Григорий Даниил Денис
Дмитрий Евгений Егор Иван Игорь Илья Кирилл Константин Леонид Максим Михаил Никита
Николай Олег Павел Петр Пётр Роман Руслан Семен Семён Сергей Станислав Степан Тимофей
Тимур Федор Фёдор Юрий Ярослав Богдан Эдуард
Александра Алина Алла Анастасия Анна Валентина Валерия Вера Виктория Галина Дарья
Диана Евгения Екатерина Елена Жанна Зоя Инна Ирина Кристина Ксения Лариса Лидия
Любовь Людмила Маргарита Марина Мария Надежда Наталья Нина Оксана Ольга Полина
Светлана Татьяна Ульяна Юлия Яна Карина Юлиана
""".split()
_PII_NAME_STEMS = set()
for _n in _PII_FIRST_NAMES:
    _PII_NAME_STEMS.add(_n[:-1])
    if len(_n) > 4:
        _PII_NAME_STEMS.add(_n[:-2])


def _is_name_like(word):
    for stem in _PII_NAME_STEMS:
        if word.startswith(stem) and len(word) - len(stem) <= 2:
            return True
    return False


def anonymize_text(text):
    """Обезличиваем персональные данные в тексте перед отправкой во внешний LLM API:
    телефоны, email, ИНН/СНИЛС, и вероятные ФИО. Возвращает (обезличенный_текст, найдено_ли_что-то)."""
    if not text:
        return text, False

    found = False
    result = text

    phone_re = re.compile(r'(\+7|8)[\s\-]?\(?\d{3}\)?[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}')
    result, n = phone_re.subn('[номер телефона]', result)
    found = found or n > 0

    email_re = re.compile(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')
    result, n = email_re.subn('[email]', result)
    found = found or n > 0

    inn_re = re.compile(r'\b(ИНН|СНИЛС|ОГРН|КПП)[\s:]*\d{8,15}\b', re.IGNORECASE)
    result, n = inn_re.subn(lambda m: f'[{m.group(1)}]', result)
    found = found or n > 0

    initials_re = re.compile(r'\b[А-ЯЁ][а-яё]+(?:у|ой|ым|е)?\s[А-ЯЁ]\.\s?(?:[А-ЯЁ]\.)?')
    result, n = initials_re.subn('[Сотрудник]', result)
    found = found or n > 0

    name_pair_re = re.compile(r'\b([А-ЯЁ][а-яё]+)\s([А-ЯЁ][а-яё]+)\b')
    def _repl(m):
        nonlocal found
        if _is_name_like(m.group(1)) or _is_name_like(m.group(2)):
            found = True
            return '[Сотрудник]'
        return m.group(0)
    result = name_pair_re.sub(_repl, result)

    return result, found


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


def build_prompt(question, documents, history=None, file_text=None, combined_mode=False):
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

    if combined_mode:
        answer_rule = 'Если в контексте базы знаний есть ответ — используй его в первую очередь и указывай что это из базы знаний компании. Если в контексте нет ответа — можешь ответить на основе своих общих знаний, но явно обозначь это фразой вроде "Это не из базы знаний компании, а общая информация:".'
    else:
        answer_rule = 'Отвечай ТОЛЬКО на основе предоставленного контекста. Если в контексте нет ответа на вопрос — прямо скажи: "К сожалению, точного ответа в базе знаний не нашлось" и не выдумывай ответ.'

    return f"""Ты — ИИ-ассистент компании Райтек. Отвечай на вопросы сотрудников на основе предоставленного контекста из базы знаний.

Документы ниже уже отобраны отдельным алгоритмом как релевантные вопросу.

{answer_rule}

Не отвечай на вопросы не связанные с работой компании (погода, курсы валют, рецепты и т.д.) — вежливо откажи, сославшись на то что ты отвечаешь только по базе знаний компании.

Если вопрос касается чисел, времени, дат, диапазонов ("можно ли в такое-то время", "хватит ли столько дней" и т.п.) — рассуждай пошагово вслух перед финальным ответом: выпиши точные цифры/границы из контекста, сравни их с тем что спрашивает пользователь, и только потом дай вывод. Не торопись с ответом "да/нет" не проверив числа.

Пиши обычным текстом без markdown-разметки — не используй звёздочки для выделения жирным (**слово**), не используй решётки для заголовков. Обычные предложения и списки через тире.
{history_block}{file_block}
Контекст из базы знаний:
{context}

Вопрос пользователя: {question}

Дай чёткий, содержательный ответ по существу."""


def build_prompt_external(question, history=None, file_text=None):
    """Промпт для режима 'Внешняя LLM' — без базы знаний вообще, обычный ассистент общего назначения."""
    history_block = ""
    if history:
        turns = "\n".join([f"{'Сотрудник' if h['role']=='user' else 'Ассистент'}: {h['content']}" for h in history[-6:]])
        history_block = f"\nПредыдущая часть разговора:\n{turns}\n"

    file_block = ""
    if file_text:
        file_block = f"\nПриложенный файл:\n{file_text}\n"

    return f"""Ты — ИИ-ассистент. Отвечай на вопрос пользователя на основе своих общих знаний (эта беседа НЕ использует корпоративную базу знаний).

Пиши обычным текстом без markdown-разметки — не используй звёздочки для выделения жирным, не используй решётки для заголовков.
{history_block}{file_block}
Вопрос пользователя: {question}

Дай чёткий, содержательный ответ."""


def log_interaction(question, answer, mode, model_used, is_off_topic=False, used_general_knowledge=False, anonymization_applied=False, sources=None):
    """Логируем КАЖДОЕ взаимодействие (включая случаи 'не найдено') для точной статистики."""
    try:
        supabase.table("chat_logs").insert({
            "user_name": session.get("user", "web_user"),
            "question": question,
            "answer": answer,
            "is_off_topic": is_off_topic,
            "model_used": model_used,
            "mode": mode,
            "used_general_knowledge": used_general_knowledge,
            "anonymization_applied": anonymization_applied,
            "sources_json": [s["title"] for s in (sources or [])],
        }).execute()
    except Exception:
        pass


def answer_question(question, model_choice="gigachat", history=None, file_text=None, mode="kb_only"):
    giga = get_giga()
    debug = {"question": question, "model_choice": model_choice, "mode": mode}

    # Обезличивание "на лету" — вопрос и файл прогоняются через фильтр ПЕРЕД отправкой во внешний LLM API
    anon_question, question_had_pii = anonymize_text(question)
    anon_file_text, file_had_pii = anonymize_text(file_text) if file_text else (file_text, False)
    anonymization_applied = question_had_pii or file_had_pii
    debug["anonymization_applied"] = anonymization_applied
    if question_had_pii:
        debug["question_anonymized_preview"] = anon_question

    # ===== РЕЖИМ "ВНЕШНЯЯ LLM" — без базы знаний вообще =====
    if mode == "external":
        prompt = build_prompt_external(anon_question, history=history, file_text=anon_file_text)
        debug["final_prompt_preview"] = prompt[:1500]
        answer, actual_model = ask_llm(giga, prompt, model_choice)
        answer = answer.replace("**", "").replace("##", "").replace("###", "")
        debug["result"] = "answered_external"
        log_interaction(question, answer, mode, actual_model, is_off_topic=False,
                        used_general_knowledge=True, anonymization_applied=anonymization_applied)
        return {"answer": answer, "sources": [], "off_topic": False, "debug": debug, "model_used": actual_model}

    # ===== РЕЖИМЫ "СТРОГО ПО БАЗЕ" и "КОМБИНИРОВАННЫЙ" — ищем в базе знаний =====
    combined_mode = (mode == "combined")

    # Шаг 1: расширяем вопрос — синонимы + LLM-подсказки
    expanded = expand_query(anon_question)
    llm_keywords = expand_query_with_llm(giga, anon_question)
    search_query = expanded + " " + llm_keywords
    debug["synonym_expansion"] = expanded
    debug["llm_keywords"] = llm_keywords

    # Шаг 2: ищем — точный поиск по расширенному запросу, нечёткий по короткому исходному вопросу
    candidates = search_documents(exact_query=search_query, fuzzy_query=anon_question)
    debug["candidates"] = [
        {"title": c["title"], "rank": round(c["rank"], 4), "match_type": c.get("match_type", "?")}
        for c in candidates
    ]

    # В режиме "строго по базе" — если ничего не нашли и файла нет, честно сообщаем что не нашли.
    # В комбинированном режиме — продолжаем даже без совпадений, модель ответит из общих знаний.
    if not candidates and not anon_file_text and not combined_mode:
        debug["result"] = "off_topic_no_candidates"
        answer_text = "К сожалению, точного ответа в базе знаний не нашлось. Могу создать заявку по вашему вопросу для специалиста."
        log_interaction(question, answer_text, mode, model_choice, is_off_topic=True,
                        anonymization_applied=anonymization_applied)
        return {
            "answer": answer_text,
            "sources": [], "off_topic": True, "debug": debug, "model_used": model_choice
        }

    documents = []
    reranker_raw = None
    if candidates:
        # Шаг 3: reranker — LLM выбирает реально релевантные документы из пула
        documents, reranker_raw = rerank_with_llm(giga, anon_question, candidates, top_n=3)
    debug["reranker_selected"] = [d["title"] for d in documents]
    debug["reranker_raw_response"] = reranker_raw

    if not documents and not anon_file_text and not combined_mode:
        debug["result"] = "off_topic_reranker_empty"
        answer_text = "К сожалению, точного ответа в базе знаний не нашлось. Могу создать заявку по вашему вопросу для специалиста."
        log_interaction(question, answer_text, mode, model_choice, is_off_topic=True,
                        anonymization_applied=anonymization_applied)
        return {
            "answer": answer_text,
            "sources": [], "off_topic": True, "debug": debug, "model_used": model_choice
        }

    prompt = build_prompt(anon_question, documents, history=history, file_text=anon_file_text, combined_mode=combined_mode)
    debug["final_prompt_preview"] = prompt[:1500]

    # Шаг 4: финальный ответ — генерируем той моделью, которую выбрал пользователь
    answer, actual_model = ask_llm(giga, prompt, model_choice)
    answer = answer.replace("**", "").replace("##", "").replace("###", "")
    debug["result"] = "answered"
    debug["actual_model"] = actual_model

    sources = [{"title": doc["title"], "url": doc.get("source_url")} for doc in documents]
    used_general_knowledge = combined_mode and not documents

    log_interaction(question, answer, mode, actual_model, is_off_topic=False,
                    used_general_knowledge=used_general_knowledge, anonymization_applied=anonymization_applied,
                    sources=sources)

    return {"answer": answer, "sources": sources, "off_topic": False, "debug": debug, "model_used": actual_model}


@app.route("/")
def index():
    return render_template("index.html")


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user" not in session:
            return jsonify({"error": "Требуется авторизация"}), 401
        return f(*args, **kwargs)
    return decorated


def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user" not in session:
            return jsonify({"error": "Требуется авторизация"}), 401
        if session.get("role") != "admin":
            return jsonify({"error": "Требуются права администратора"}), 403
        return f(*args, **kwargs)
    return decorated


@app.route("/api/login", methods=["POST"])
def login():
    data = request.get_json() or {}
    username = (data.get("username") or "").strip()
    password = data.get("password") or ""

    if not username or not password:
        return jsonify({"error": "Введите логин и пароль"}), 400

    result = supabase.table("app_users").select("*").eq("username", username).execute()
    if not result.data:
        return jsonify({"error": "Неверный логин или пароль"}), 401

    user = result.data[0]
    if not check_password_hash(user["password_hash"], password):
        return jsonify({"error": "Неверный логин или пароль"}), 401

    session["user"] = user["username"]
    session["role"] = user["role"]
    session["display_name"] = user.get("display_name") or user["username"]

    return jsonify({"username": user["username"], "role": user["role"], "display_name": session["display_name"]})


@app.route("/api/logout", methods=["POST"])
def logout():
    session.clear()
    return jsonify({"ok": True})


@app.route("/api/session")
def get_session():
    if "user" not in session:
        return jsonify({"authenticated": False})
    return jsonify({
        "authenticated": True,
        "username": session["user"],
        "role": session["role"],
        "display_name": session.get("display_name", session["user"])
    })


@app.route("/api/ask", methods=["POST"])
@login_required
def ask():
    data = request.get_json()
    question = (data or {}).get("question", "").strip()
    model_choice = (data or {}).get("model", "gigachat")
    history = (data or {}).get("history", [])
    file_text = (data or {}).get("file_text")
    mode = (data or {}).get("mode", "kb_only")

    if not question:
        return jsonify({"error": "Пустой вопрос"}), 400

    try:
        result = answer_question(question, model_choice, history=history, file_text=file_text, mode=mode)
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/upload", methods=["POST"])
@login_required
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


@app.route("/admin")
def admin_page():
    return render_template("admin.html")


@app.route("/api/admin/stats")
@admin_required
def admin_stats():
    period_days = int(request.args.get("days", 30))
    from datetime import datetime, timedelta, timezone
    since = (datetime.now(timezone.utc) - timedelta(days=period_days)).isoformat()

    logs = supabase.table("chat_logs").select("*").gte("created_at", since).execute().data

    total = len(logs)
    off_topic_count = sum(1 for l in logs if l.get("is_off_topic"))
    answered_count = total - off_topic_count
    answered_pct = round(answered_count / total * 100, 1) if total else 0

    by_model = {}
    for l in logs:
        m = l.get("model_used") or "неизвестно"
        by_model[m] = by_model.get(m, 0) + 1

    by_mode = {}
    for l in logs:
        m = l.get("mode") or "kb_only"
        by_mode[m] = by_mode.get(m, 0) + 1

    anonymized_count = sum(1 for l in logs if l.get("anonymization_applied"))
    general_knowledge_count = sum(1 for l in logs if l.get("used_general_knowledge"))

    # Топ вопросов, на которые не нашли ответ — полезно для наполнения базы знаний
    off_topic_questions = [l["question"] for l in logs if l.get("is_off_topic")][:20]

    return jsonify({
        "period_days": period_days,
        "total_questions": total,
        "answered_count": answered_count,
        "answered_pct": answered_pct,
        "off_topic_count": off_topic_count,
        "anonymized_count": anonymized_count,
        "general_knowledge_count": general_knowledge_count,
        "by_model": by_model,
        "by_mode": by_mode,
        "off_topic_questions": off_topic_questions,
    })


@app.route("/api/admin/logs")
@admin_required
def admin_logs():
    limit = int(request.args.get("limit", 50))
    result = supabase.table("chat_logs").select("*").order("created_at", desc=True).limit(limit).execute()
    return jsonify({"logs": result.data})


@app.route("/api/admin/sources")
@admin_required
def admin_sources():
    result = supabase.table("kb_documents").select("source").execute()
    counts = {}
    for row in result.data:
        s = row.get("source") or "неизвестно"
        counts[s] = counts.get(s, 0) + 1
    return jsonify({"sources": counts, "total_chunks": len(result.data)})


@app.route("/api/admin/timeseries")
@admin_required
def admin_timeseries():
    period_days = int(request.args.get("days", 30))
    from datetime import datetime, timedelta, timezone
    since = (datetime.now(timezone.utc) - timedelta(days=period_days)).isoformat()

    logs = supabase.table("chat_logs").select("created_at, is_off_topic").gte("created_at", since).execute().data

    daily = {}
    for l in logs:
        day = l["created_at"][:10]  # YYYY-MM-DD
        if day not in daily:
            daily[day] = {"total": 0, "answered": 0}
        daily[day]["total"] += 1
        if not l.get("is_off_topic"):
            daily[day]["answered"] += 1

    sorted_days = sorted(daily.keys())
    return jsonify({
        "labels": sorted_days,
        "total": [daily[d]["total"] for d in sorted_days],
        "answered": [daily[d]["answered"] for d in sorted_days],
    })


@app.route("/api/admin/top_documents")
@admin_required
def admin_top_documents():
    period_days = int(request.args.get("days", 30))
    from datetime import datetime, timedelta, timezone
    since = (datetime.now(timezone.utc) - timedelta(days=period_days)).isoformat()

    logs = supabase.table("chat_logs").select("sources_json").gte("created_at", since).execute().data

    counts = {}
    for l in logs:
        for title in (l.get("sources_json") or []):
            counts[title] = counts.get(title, 0) + 1

    top = sorted(counts.items(), key=lambda x: -x[1])[:10]
    return jsonify({"top_documents": [{"title": t, "count": c} for t, c in top]})


if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5001))
    app.run(host="0.0.0.0", port=port)
