import glob
import os
import re
from flask import Flask, request, render_template_string
from docx import Document
from difflib import SequenceMatcher

app = Flask(__name__)

HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Поиск ответов</title>
</head>
<body>
    <h1>Поиск по вопросам (Приблизительный поиск)</h1>
    <form method="POST">
        <label for="question">Введите вопрос:</label><br>
        <input type="text" id="question" name="question" style="width:300px;" required>
        <button type="submit">Найти ответ</button>
    </form>
    {% if answer %}
        <hr>
        <h2>Результат:</h2>
        <p>{{ answer }}</p>
    {% endif %}
</body>
</html>
"""

def parse_docx_file(file_path):
    """
    Читает .docx-файл и извлекает пары (вопрос, ответ).
    Логика:
      1) Ищет строки, содержащие 'S:' — после них предполагается вопрос.
      2) В той же строке ищет разделение (':' или '=') для ответа.
         Если нет, ответ может быть на следующей строке (начинается с '=' или просто текст).
      3) Символ 'I:' может указывать на начало новой записи, но не обязателен.
    Возвращает список кортежей (вопрос, ответ).
    """
    doc = Document(file_path)
    lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    results = []
    i = 0
    while i < len(lines):
        line = lines[i]
        # Ищем "S:"
        if "S:" in line:
            parts = line.split("S:", 1)
            s_part = parts[-1].strip()
            sep_colon = s_part.find(":")
            sep_equal = s_part.find("=")

            question = ""
            answer = ""

            if sep_colon != -1:
                question = s_part[:sep_colon].strip()
                answer_candidate = s_part[sep_colon + 1:].strip()
                if answer_candidate:
                    answer = answer_candidate
                else:
                    # Если нет ответа после двоеточия, пытаемся взять следующую строку
                    if i + 1 < len(lines):
                        nxt = lines[i + 1]
                        if nxt.startswith("="):
                            answer = nxt.lstrip("= ").strip()
                            i += 1
                        elif (not nxt.startswith("I:") and not nxt.startswith("S:")):
                            answer = nxt
                            i += 1

            elif sep_equal != -1:
                question = s_part[:sep_equal].strip()
                answer_candidate = s_part[sep_equal + 1:].strip()
                if answer_candidate:
                    answer = answer_candidate
                else:
                    if i + 1 < len(lines):
                        nxt = lines[i + 1]
                        if nxt.startswith("="):
                            answer = nxt.lstrip("= ").strip()
                            i += 1
                        elif (not nxt.startswith("I:") and not nxt.startswith("S:")):
                            answer = nxt
                            i += 1
            else:
                # Нет ':' или '=', значит всё s_part — это вопрос
                question = s_part
                # Ответ может оказаться на следующей строке
                if i + 1 < len(lines):
                    nxt = lines[i + 1]
                    if nxt.startswith("="):
                        answer = nxt.lstrip("= ").strip()
                        i += 1
                    elif (not nxt.startswith("I:") and not nxt.startswith("S:")):
                        answer = nxt
                        i += 1

            if question:
                results.append((question, answer))
        i += 1

    return results

def load_qa_data(docx_folder="docs"):
    """
    Загружает (вопрос, ответ) из всех .docx-файлов,
    названия которых начинаются с 'OSP' и цифрой (1-12).
    """
    pattern = os.path.join(docx_folder, "OSP*.docx")
    all_qa = []
    for file_path in glob.glob(pattern):
        qa_pairs = parse_docx_file(file_path)
        all_qa.extend(qa_pairs)
    return all_qa

# Функция для "нормализации" строк, чтобы убрать лишние пробелы, знаки и т.д.
def normalize_text(text):
    # Убираем всякие ' и пробелы вокруг слов, приводим к нижнему регистру
    # Можно также убирать пунктуацию, если нужно.
    lower_text = text.lower()
    # Заменяем множественные пробелы на один и убираем неалфавитные символы, если это актуально
    # Но аккуратно, вдруг нам нужны цифры. Для примера удалим только одиночные апострофы:
    processed = re.sub(r"[']", "", lower_text)
    # Опционально можно убрать и другие знаки, но это зависит от задачи
    # processed = re.sub(r"[^a-zа-я0-9\s]", "", processed)
    # Убираем повторные пробелы:
    processed = re.sub(r"\s+", " ", processed).strip()
    return processed

# Список (вопрос, ответ), выгруженный из файлов
qa_list = load_qa_data()

@app.route("/", methods=["GET", "POST"])
def index():
    found_answer = None
    if request.method == "POST":
        user_question = request.form.get("question", "").strip()
        user_normalized = normalize_text(user_question)

        best_match_answer = None
        best_similarity = 0.0

        # Перебираем все вопросы, ищем самый похожий
        for (q, a) in qa_list:
            question_normalized = normalize_text(q)
            # Считаем схожесть (значение от 0 до 1)
            similarity = SequenceMatcher(None, user_normalized, question_normalized).ratio()

            # Если схожесть выше, обновляем
            if similarity > best_similarity:
                best_similarity = similarity
                best_match_answer = a

        # Допустим, что если similarity выше 0.5, мы считаем это приемлемым совпадением
        if best_similarity >= 0.5 and best_match_answer:
            found_answer = best_match_answer
        else:
            found_answer = "Ответ не найден (нет достаточно похожих вопросов)."

    return render_template_string(HTML_TEMPLATE, answer=found_answer)

if __name__ == "__main__":
    app.run(debug=True)