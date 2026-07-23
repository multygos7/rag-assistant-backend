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
import pdfplumber
from docx import Document as DocxDocument
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph as DocxParagraph
# sentence_transformers НЕ импортируем здесь на уровне модуля — сам импорт тянет
# PyTorch в память, даже если саму модель эмбеддингов не грузить. Импортируем
# только внутри блока ниже, когда семантический поиск явно включён.

# ==================== НАСТРОЙКИ ====================
SUPABASE_URL = "https://jidtwjamnglkqoqizvjl.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImppZHR3amFtbmdsa3FvcWl6dmpsIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc2OTI4NDg3MSwiZXhwIjoyMDg0ODYwODcxfQ.3Yl9LIJ7BnRmAAYrzc0NiPEilphXQm8EdObmh2_y5BU"
GIGACHAT_CREDENTIALS = "MDE5ZjQ3MDgtODM2Zi03ZGU3LWJlNTMtZTQzMTI3MjE3NDVhOjkyMTU5ZDkyLWQzN2YtNDM4Ny05YTRhLWQxNjQxYmE4ZTI0ZQ=="
OPENROUTER_API_KEY = "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"

# Модели доступные через OpenRouter (ключ -> model id на OpenRouter)
OPENROUTER_MODELS = {
    "qwen": "qwen/qwen3-next-80b-a3b-instruct:free",
    "llama": "meta-llama/llama-3.3-70b-instruct:free",
    "gpt_oss": "openai/gpt-oss-20b:free",
}

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


_TRANSLIT_MAP = {
    'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'e', 'ж': 'zh',
    'з': 'z', 'и': 'i', 'й': 'y', 'к': 'k', 'л': 'l', 'м': 'm', 'н': 'n', 'о': 'o',
    'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u', 'ф': 'f', 'х': 'h', 'ц': 'ts',
    'ч': 'ch', 'ш': 'sh', 'щ': 'sch', 'ъ': '', 'ы': 'y', 'ь': '', 'э': 'e', 'ю': 'yu', 'я': 'ya',
}


def transliterate_ru_to_latin(text):
    """Механическая транслитерация — иногда технические названия в базе знаний
    пишут русское слово латиницей (например 'Казна' -> 'Kazna' в '1С:ERP_KAZNA').
    Это не перевод и не сленг, а простая побуквенная замена."""
    result = []
    for ch in text.lower():
        result.append(_TRANSLIT_MAP.get(ch, ch))
    return "".join(result)


def add_transliteration(question):
    """Добавляем к запросу транслитерированные варианты каждого достаточно длинного
    русского слова — чтобы находить технические названия написанные латиницей."""
    words = re.findall(r'[а-яё]{4,}', question.lower())
    translits = [transliterate_ru_to_latin(w) for w in words]
    if translits:
        return question + " " + " ".join(translits)
    return question
# =====================================================

app = Flask(__name__)
app.secret_key = "raytec-rag-mvp-secret-key-change-in-prod-2026"
supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

# Модель для семантического поиска (эмбеддинги) — опционально, управляется переменной
# окружения. По умолчанию ВЫКЛЮЧЕНА: sentence-transformers тянет за собой PyTorch,
# который сам по себе занимает больше памяти чем есть на бесплатном тарифе Render (512МБ)
# и вызывает аварийную остановку процесса (OOM) — try/except тут не спасает, так как
# Render убивает процесс на уровне ОС, а не через Python-исключение.
# Локально (где памяти достаточно) включай через переменную окружения:
#   ENABLE_SEMANTIC_SEARCH=true python3 app.py
import os as _os
ENABLE_SEMANTIC_SEARCH = _os.environ.get("ENABLE_SEMANTIC_SEARCH", "false").lower() == "true"

EMBEDDING_MODEL = None
if ENABLE_SEMANTIC_SEARCH:
    print("Загружаю модель эмбеддингов...")
    try:
        from sentence_transformers import SentenceTransformer
        EMBEDDING_MODEL = SentenceTransformer("sergeyzh/rubert-tiny-lite")
        print("Модель эмбеддингов загружена.")
    except Exception as e:
        EMBEDDING_MODEL = None
        print(f"Не удалось загрузить модель эмбеддингов: {e}")
else:
    print("Семантический поиск отключён (ENABLE_SEMANTIC_SEARCH не установлена) — работает точный + нечёткий поиск.")

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


def extract_docx_text(doc):
    """Извлекаем текст из docx ПОЛНОСТЬЮ — включая таблицы, а не только обычные абзацы.
    Обычный doc.paragraphs не видит текст внутри таблиц, из-за чего терялось
    много контента в документах где вся информация оформлена таблицами."""
    parts = []
    for block in doc.iter_inner_content():
        if isinstance(block, DocxParagraph):
            if block.text.strip():
                parts.append(block.text)
        elif isinstance(block, DocxTable):
            for row in block.rows:
                for cell in row.cells:
                    cell_text = cell.text.strip()
                    if cell_text:
                        parts.append(cell_text)
    return "\n".join(parts)


def extract_text_from_file(file_storage):
    """Извлекаем текст из загруженного файла: txt, pdf или docx."""
    filename = file_storage.filename.lower()
    content = file_storage.read()

    if filename.endswith(".txt"):
        text = content.decode("utf-8", errors="ignore")
    elif filename.endswith(".pdf"):
        text_parts = []
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                if page_text:
                    text_parts.append(page_text)
                for table in page.extract_tables():
                    for row in table:
                        row_text = " | ".join(cell or "" for cell in row)
                        if row_text.strip(" |"):
                            text_parts.append(row_text)
        text = "\n".join(text_parts)
    elif filename.endswith(".docx"):
        doc = DocxDocument(io.BytesIO(content))
        text = extract_docx_text(doc)
    else:
        raise ValueError("Поддерживаются только файлы .txt, .pdf, .docx")

    text = text.strip()
    if len(text) > MAX_FILE_TEXT_CHARS:
        text = text[:MAX_FILE_TEXT_CHARS] + "\n[...текст обрезан, файл слишком большой...]"
    return text


def get_giga():
    """Создаём новое подключение к GigaChat на каждый запрос (токен живёт 30 минут)."""
    return GigaChat(credentials=GIGACHAT_CREDENTIALS, verify_ssl_certs=False)


def ask_openrouter(prompt, model_id, retries=1):
    """Отправляем запрос в любую модель через OpenRouter. Если конкретная модель
    сейчас перегружена (429) — переключаемся на openrouter/free, который сам подберёт
    доступную бесплатную модель. Возвращаем текст + реальное название модели-ответчика."""
    import time

    def call(mid):
        return requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers={"Authorization": f"Bearer {OPENROUTER_API_KEY}"},
            json={"model": mid, "messages": [{"role": "user", "content": prompt}]},
            timeout=30
        )

    last_error = None
    for attempt in range(retries + 1):
        response = call(model_id)
        data = response.json()
        if "choices" in data:
            return data["choices"][0]["message"]["content"], data.get("model", model_id)
        last_error = data
        if response.status_code == 429 and attempt < retries:
            time.sleep(3)
            continue
        break

    # Модель недоступна — пробуем автоматический подбор среди других бесплатных моделей
    response = call("openrouter/free")
    data = response.json()
    if "choices" in data:
        return data["choices"][0]["message"]["content"], data.get("model", "openrouter/free")

    raise Exception(f"OpenRouter error: {last_error}")


def ask_llm(giga, prompt, model_choice):
    """Единая обёртка — отправляет промпт в выбранную модель.
    Возвращает (текст_ответа, реальное_название_модели_ответившей)."""
    if model_choice in OPENROUTER_MODELS:
        return ask_openrouter(prompt, OPENROUTER_MODELS[model_choice])
    chat = Chat(messages=[Messages(role=MessagesRole.USER, content=prompt)])
    response = giga.chat(chat)
    return response.choices[0].message.content, "GigaChat"


def expand_query_with_llm(giga, question, model_choice="llama"):
    """Просим ВЫБРАННУЮ пользователем модель подсказать связанные термины к вопросу —
    чтобы находить документы даже если пользователь не использует те же слова что в базе знаний.
    Отдельно просим английские/латинские эквиваленты русского IT-сленга —
    например 'джун' это заимствование от 'junior', а не транслитерация,
    поэтому механическая транслитерация тут не сработает, нужно именно знание сленга."""
    prompt = f"""Пользователь задал вопрос сотруднику компании: "{question}"

Твоя задача — предложить 5-8 ключевых слов и связанных терминов, которые могут встречаться в документах базы знаний компании и помогут найти ответ на этот вопрос. Учитывай синонимы, смежные термины, официальные названия процессов/модулей.

Отдельно проверь: если в вопросе есть русский IT/деловой сленг заимствованный из английского (например "джун" = junior, "мидл" = middle, "сеньор" = senior, "таск" = task, "апрув" = approve, "дедлайн" = deadline, "апи" = API, "баг" = bug, "ревью" = review и подобные) — обязательно добавь их английские эквиваленты латиницей в список, так как в базе знаний термин может быть на английском.

Ответь ТОЛЬКО списком слов через запятую, без пояснений. Например: отпуск, заявление, оформление, СБИС, кадры, junior, HR"""

    try:
        keywords, _ = ask_llm(giga, prompt, model_choice)
        return keywords.strip()
    except Exception:
        # Если модель недоступна (например закончилась квота) — просто пропускаем
        # этот шаг, поиск продолжит работать по синонимам и транслитерации
        return ""


def rerank_with_llm(giga, question, candidates, top_n=3, model_choice="llama"):
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

    try:
        raw_answer, _ = ask_llm(giga, prompt, model_choice)
        raw_answer = raw_answer.strip()
    except Exception:
        # Если ни одна модель недоступна — не роняем весь ответ, а просто берём
        # топ кандидатов по обычному рангу поиска, без "умного" отбора по смыслу
        return candidates[:top_n], "reranker unavailable — fallback to raw top-N"

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


def semantic_search(question, match_count=30):
    """Семантический поиск — находит документы близкие по СМЫСЛУ через эмбеддинги,
    а не по совпадению слов. Понимает синонимы, разное написание и т.д. без подсказок."""
    if EMBEDDING_MODEL is None:
        return []
    try:
        query_embedding = EMBEDDING_MODEL.encode(question).tolist()
        result = supabase.rpc("semantic_search_kb", {
            "query_embedding": query_embedding,
            "match_count": match_count
        }).execute()
        for doc in result.data:
            doc["match_type"] = "semantic"
        return result.data
    except Exception:
        return []


def search_documents(exact_query, fuzzy_query, semantic_query=None, exact_limit=60, fuzzy_limit=15, semantic_limit=30):
    """Ищем релевантные документы тремя независимыми способами и объединяем:
    точный поиск (глубокий, не теряет специфичные темы), нечёткий (страхует от опечаток),
    семантический (понимает смысл и синонимы без ручных подсказок)."""
    result = supabase.rpc("search_kb_documents", {
        "exact_query_text": exact_query,
        "fuzzy_query_text": fuzzy_query,
        "exact_limit": exact_limit,
        "fuzzy_limit": fuzzy_limit
    }).execute()
    combined = list(result.data)

    if semantic_query:
        semantic_results = semantic_search(semantic_query, match_count=semantic_limit)
        seen_ids = {doc["id"] for doc in combined}
        for doc in semantic_results:
            if doc["id"] not in seen_ids:
                combined.append(doc)
                seen_ids.add(doc["id"])

    return combined


def find_correction(question):
    """Проверяем — не было ли раньше похожего вопроса, на который дали неверный
    ответ, и пользователь (или админ) указал правильное исправление. Если да —
    подсовываем это исправление модели, чтобы не повторять старую ошибку."""
    try:
        result = supabase.rpc("find_correction", {"query_text": question}).execute()
        if result.data:
            return result.data[0]["correction"]
    except Exception:
        pass
    return None


def build_prompt(question, documents, history=None, file_text=None, combined_mode=False, correction=None):
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

    correction_block = ""
    if correction:
        correction_block = f"""
ВАЖНО: на очень похожий вопрос ранее уже был дан неверный ответ, и вот проверенное исправление от сотрудника компании — обязательно учти его и не повторяй прошлую ошибку:
"{correction}"
"""

    if combined_mode:
        answer_rule = 'Если в контексте базы знаний есть ответ — используй его в первую очередь и указывай что это из базы знаний компании. Если в контексте нет ответа — можешь ответить на основе своих общих знаний, но явно обозначь это фразой вроде "Это не из базы знаний компании, а общая информация:".'
    else:
        answer_rule = 'Отвечай ТОЛЬКО на основе предоставленного контекста. Если в контексте нет ответа на вопрос — прямо скажи: "К сожалению, точного ответа в базе знаний не нашлось" и не выдумывай ответ.'

    return f"""Ты — ИИ-ассистент компании Райтек. Отвечай на вопросы сотрудников на основе предоставленного контекста из базы знаний.

Документы ниже уже отобраны отдельным алгоритмом как релевантные вопросу.

{answer_rule}

Не отвечай на вопросы не связанные с работой компании (погода, курсы валют, рецепты и т.д.) — вежливо откажи, сославшись на то что ты отвечаешь только по базе знаний компании.

Если вопрос касается чисел, времени, дат, диапазонов ("можно ли в такое-то время", "хватит ли столько дней" и т.п.) — рассуждай пошагово вслух перед финальным ответом: выпиши точные цифры/границы из контекста, сравни их с тем что спрашивает пользователь, и только потом дай вывод. Не торопись с ответом "да/нет" не проверив числа.

Пример правильного рассуждения про диапазон времени:
Вопрос: "Можно пообедать в 14:00, если перерыв разрешён с 12:00 до 15:00?"
Рассуждение: Границы диапазона — 12:00 и 15:00. Проверяемое время — 14:00. 14:00 больше 12:00 и меньше 15:00, значит оно ВНУТРИ диапазона.
Вывод: Да, можно, это время попадает в разрешённый интервал.

Пиши обычным текстом без markdown-разметки — не используй звёздочки для выделения жирным (**слово**), не используй решётки для заголовков. Обычные предложения и списки через тире.
{correction_block}{history_block}{file_block}
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


def answer_question(question, model_choice="llama", history=None, file_text=None, mode="kb_only"):
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

    # Шаг 1: расширяем вопрос — синонимы + LLM-подсказки + транслитерация (для тех.названий типа Kazna)
    expanded = expand_query(anon_question)
    expanded = add_transliteration(expanded)
    llm_keywords = expand_query_with_llm(giga, anon_question, model_choice)
    search_query = expanded + " " + llm_keywords
    debug["synonym_expansion"] = expanded
    debug["llm_keywords"] = llm_keywords

    # Шаг 2: ищем — точный поиск по расширенному запросу, нечёткий по короткому исходному вопросу
    candidates = search_documents(exact_query=search_query, fuzzy_query=anon_question, semantic_query=anon_question)
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
        documents, reranker_raw = rerank_with_llm(giga, anon_question, candidates, top_n=5, model_choice=model_choice)
    debug["reranker_selected"] = [d["title"] for d in documents]
    debug["reranker_raw_response"] = reranker_raw

    if not documents and not anon_file_text and not combined_mode:
        debug["result"] = "off_topic_reranker_empty"
        answer_text = "К сожалению, точного ответа в базе знаний не нашлось. Могу создать заявку по вашему вопросу для специалиста."

        # Берём топ несколько "сырых" кандидатов (которые reranker не выбрал как явно релевантные)
        # и предлагаем пользователю самому подтвердить — вдруг один из них всё же то что нужно
        possible_matches = []
        seen_titles = set()
        for c in candidates:
            base_title = re.sub(r'\s*\(часть \d+/\d+\)\s*$', '', c["title"]).strip()
            if base_title not in seen_titles:
                possible_matches.append({"id": c["id"], "title": base_title, "url": c.get("source_url")})
                seen_titles.add(base_title)
            if len(possible_matches) >= 3:
                break

        log_interaction(question, answer_text, mode, model_choice, is_off_topic=True,
                        anonymization_applied=anonymization_applied)
        return {
            "answer": answer_text,
            "sources": [], "off_topic": True, "debug": debug, "model_used": model_choice,
            "possible_matches": possible_matches
        }

    known_correction = find_correction(anon_question)
    debug["known_correction_found"] = bool(known_correction)
    prompt = build_prompt(anon_question, documents[:3], history=history, file_text=anon_file_text, combined_mode=combined_mode, correction=known_correction)
    debug["final_prompt_preview"] = prompt[:1500]

    # Шаг 4: финальный ответ — генерируем той моделью, которую выбрал пользователь
    answer, actual_model = ask_llm(giga, prompt, model_choice)
    answer = answer.replace("**", "").replace("##", "").replace("###", "")
    debug["actual_model"] = actual_model

    # Модель могла САМА решить что ответа нет, несмотря на найденные документы —
    # проверяем финальный текст, а не только факт наличия документов, чтобы
    # статистика и источники были честными
    answer_says_not_found = (
        "точного ответа" in answer.lower() and "не нашл" in answer.lower()
    ) or ("к сожалению" in answer.lower()[:50] and "не нашл" in answer.lower())

    if answer_says_not_found:
        debug["result"] = "answered_but_model_said_not_found"
        possible_matches = []
        seen_titles = set()
        for d in documents:
            base_title = re.sub(r'\s*\(часть \d+/\d+\)\s*$', '', d["title"]).strip()
            if base_title not in seen_titles:
                possible_matches.append({"id": d["id"], "title": base_title, "url": d.get("source_url")})
                seen_titles.add(base_title)
        log_interaction(question, answer, mode, actual_model, is_off_topic=True,
                        anonymization_applied=anonymization_applied)
        return {
            "answer": answer, "sources": [], "off_topic": True, "debug": debug, "model_used": actual_model,
            "possible_matches": possible_matches
        }

    debug["result"] = "answered"
    sources = [{"title": doc["title"], "url": doc.get("source_url")} for doc in documents[:3]]
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


DEFAULT_SUGGESTIONS = [
    "Как оформить отпуск?",
    "Правила по больничному листу",
    "Как создать заявку в Jira?",
    "Доступ к 1С:ДО",
]


@app.route("/api/popular_questions")
@login_required
def popular_questions():
    """Показываем реально популярные вопросы из логов (успешно отвеченные),
    а если истории ещё мало — дополняем стандартными примерами."""
    def normalize(text):
        text = text.strip()
        if not text:
            return text
        text = text[0].upper() + text[1:]
        if not text.endswith(("?", ".", "!")):
            text += "?"
        return text

    try:
        logs = supabase.table("chat_logs").select("question").eq("is_off_topic", False).order("created_at", desc=True).limit(300).execute().data
    except Exception:
        logs = []

    counts = {}
    for l in logs:
        q = (l.get("question") or "").strip()
        if 5 < len(q) < 80:  # игнорируем пустые/слишком длинные вопросы для подсказок
            key = q.lower()
            if key not in counts:
                counts[key] = {"text": normalize(q), "count": 0}
            counts[key]["count"] += 1

    top = sorted(counts.values(), key=lambda x: -x["count"])
    result = []
    seen_keys = set()
    for item in top:
        key = item["text"].lower()
        if key not in seen_keys:
            result.append(item["text"])
            seen_keys.add(key)
        if len(result) >= 4:
            break

    for fallback in DEFAULT_SUGGESTIONS:
        if len(result) >= 4:
            break
        if fallback.lower() not in seen_keys:
            result.append(fallback)
            seen_keys.add(fallback.lower())

    return jsonify({"suggestions": result[:4]})


@app.route("/api/ask", methods=["POST"])
@login_required
def ask():
    data = request.get_json()
    question = (data or {}).get("question", "").strip()
    model_choice = (data or {}).get("model", "llama")
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


@app.route("/api/feedback", methods=["POST"])
@login_required
def submit_feedback():
    data = request.get_json() or {}
    question = (data.get("question") or "").strip()
    answer = (data.get("answer") or "").strip()
    feedback = data.get("feedback")  # 'up' или 'down'
    correction = (data.get("correction") or "").strip() or None

    if not question or not answer or feedback not in ("up", "down"):
        return jsonify({"error": "Некорректные данные"}), 400

    try:
        supabase.table("answer_feedback").insert({
            "question": question,
            "answer": answer,
            "feedback": feedback,
            "correction": correction,
            "user_name": session.get("user", "web_user"),
        }).execute()
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/ask_about_document", methods=["POST"])
@login_required
def ask_about_document():
    """Пользователь сам подтвердил — 'да, это тот документ'. Отвечаем уверенно
    именно на основе него, минуя поиск и reranker."""
    data = request.get_json() or {}
    question = (data.get("question") or "").strip()
    document_id = data.get("document_id")
    model_choice = data.get("model", "llama")

    if not question or not document_id:
        return jsonify({"error": "Не хватает данных"}), 400

    try:
        result = supabase.table("kb_documents").select("id, title, content, source_url").eq("id", document_id).execute()
        if not result.data:
            return jsonify({"error": "Документ не найден"}), 404
        doc = result.data[0]

        anon_question, question_had_pii = anonymize_text(question)
        giga = get_giga()
        prompt = build_prompt(anon_question, [doc], combined_mode=True)
        answer, actual_model = ask_llm(giga, prompt, model_choice)
        answer = answer.replace("**", "").replace("##", "").replace("###", "")

        sources = [{"title": doc["title"], "url": doc.get("source_url")}]
        log_interaction(question, answer, "kb_only", actual_model, is_off_topic=False,
                        anonymization_applied=question_had_pii, sources=sources)

        return jsonify({"answer": answer, "sources": sources, "off_topic": False, "model_used": actual_model})
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


@app.route("/api/admin/feedback_stats")
@admin_required
def admin_feedback_stats():
    period_days = int(request.args.get("days", 30))
    from datetime import datetime, timedelta, timezone
    since = (datetime.now(timezone.utc) - timedelta(days=period_days)).isoformat()

    feedback = supabase.table("answer_feedback").select("*").gte("created_at", since).order("created_at", desc=True).execute().data

    total = len(feedback)
    up_count = sum(1 for f in feedback if f["feedback"] == "up")
    down_count = total - up_count
    up_pct = round(up_count / total * 100, 1) if total else 0

    with_correction = [f for f in feedback if f["feedback"] == "down" and f.get("correction")]
    without_correction = [f for f in feedback if f["feedback"] == "down" and not f.get("correction")]

    recent_down = [
        {
            "question": f["question"],
            "answer": f["answer"][:200],
            "correction": f.get("correction"),
            "user_name": f.get("user_name"),
            "created_at": f["created_at"],
        }
        for f in feedback if f["feedback"] == "down"
    ][:20]

    return jsonify({
        "total": total,
        "up_count": up_count,
        "down_count": down_count,
        "up_pct": up_pct,
        "corrections_active": len(with_correction),
        "downvotes_no_correction": len(without_correction),
        "recent_down": recent_down,
    })


if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5001))
    app.run(host="0.0.0.0", port=port)
