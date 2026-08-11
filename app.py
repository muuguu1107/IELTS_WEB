
from flask import Flask, render_template, abort
from docx import Document
import os
import html
import re

app = Flask(__name__)


# =========================================================
# FOLDER
# =========================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

LSSN_DIR = os.path.join(
    BASE_DIR,
    "Lsn"
)


# =========================================================
# WORD FORMAT -> HTML
# =========================================================

def format_runs(paragraph):

    result = []

    for run in paragraph.runs:

        text = run.text

        if not text:
            continue

        text = html.escape(text)

        # Font color
        if run.font.color and run.font.color.rgb:

            color = str(run.font.color.rgb)

            text = (
                f'<span style="color:#{color}">'
                f'{text}'
                f'</span>'
            )

        # Bold
        if run.bold:
            text = f"<strong>{text}</strong>"

        # Italic
        if run.italic:
            text = f"<em>{text}</em>"

        # Underline
        if run.underline:
            text = f"<u>{text}</u>"

        result.append(text)

    return "".join(result)


# =========================================================
# HEADING 1 ШАЛГАХ
# =========================================================

def is_heading_1(paragraph):

    if not paragraph.style:
        return False

    style_name = paragraph.style.name.strip().lower()

    return style_name in [
        "heading 1",
        "heading1",
        "гарчиг 1"
    ]


# =========================================================
# ШИНЭ ҮГСИЙН ГАРЧИГ ШАЛГАХ
# =========================================================

def is_new_words_title(text):

    text = text.strip().lower()

    return text in [
        "шинэ үг",
        "шинэ үгс",
        "new word",
        "new words",
        "vocabulary"
    ]


# =========================================================
# ШИНЭ ҮГИЙН МӨР ЗАДЛАХ
# =========================================================

def parse_word(text):

    text = text.strip()

    # Жишээ:
    # norm — хэвшил
    # remove — зайлуулах
    # stem - иш

    match = re.match(
        r"^(.+?)\s*[—–-]\s*(.+)$",
        text
    )

    if not match:
        return None

    word = match.group(1).strip()

    meaning = match.group(2).strip()

    if not word or not meaning:
        return None

    return {
        "word": word,
        "meaning": meaning
    }


# =========================================================
# НЭГ ХИЧЭЭЛИЙН АГУУЛГА
# =========================================================

def parse_content(paragraphs):

    content = []

    i = 0

    while i < len(paragraphs):

        paragraph = paragraphs[i]

        text = paragraph.text.strip()

        if not text:

            i += 1

            continue


        # =================================================
        # ШИНЭ ҮГС
        # =================================================

        if is_new_words_title(text):

            words = []

            i += 1

            while i < len(paragraphs):

                p = paragraphs[i]

                p_text = p.text.strip()

                if not p_text:

                    i += 1

                    continue


                # ⭐ хэсэг эхэлбэл шинэ үг дуусна

                if p_text.startswith("⭐"):

                    break


                # Дараагийн Heading 1

                if is_heading_1(p):

                    break


                word = parse_word(p_text)

                if word:

                    words.append(word)

                    i += 1

                else:

                    # Энэ нь шинэ үг биш бол
                    # дараагийн ердийн paragraph
                    break


            if words:

                content.append({
                    "type": "new_words",
                    "words": words
                })

            continue


        # =================================================
        # ⭐ ХЭЛЦ / БҮТЭЦ
        # =================================================

        if text.startswith("⭐"):

            content.append({
                "type": "phrase",
                "html": format_runs(paragraph)
            })

            i += 1

            continue


        # =================================================
        # HEADING 2
        # =================================================

        style = ""

        if paragraph.style:

            style = paragraph.style.name.lower()


        if "heading 2" in style:

            content.append({
                "type": "heading2",
                "html": format_runs(paragraph)
            })

            i += 1

            continue


        # =================================================
        # HEADING 3
        # =================================================

        if "heading 3" in style:

            content.append({
                "type": "heading3",
                "html": format_runs(paragraph)
            })

            i += 1

            continue


        # =================================================
        # BULLET
        # =================================================

        if "list bullet" in style.lower():

            content.append({
                "type": "bullet",
                "html": format_runs(paragraph)
            })

            i += 1

            continue


        # =================================================
        # ЕРДИЙН PARAGRAPH
        # =================================================

        content.append({
            "type": "paragraph",
            "html": format_runs(paragraph)
        })

        i += 1


    return content


# =========================================================
# WORD ФАЙЛ УНШИХ
# =========================================================

def read_word(filepath):

    document = Document(filepath)

    lessons = []

    current_lesson = None

    current_paragraphs = []


    for paragraph in document.paragraphs:

        text = paragraph.text.strip()


        # Хоосон мөр
        if not text:

            continue


        # =================================================
        # HEADING 1 = ХИЧЭЭЛ
        # =================================================

        if is_heading_1(paragraph):


            # ӨМНӨХ ХИЧЭЭЛИЙГ ХАДГАЛАХ

            if current_lesson is not None:

                current_lesson["content"] = \
                    parse_content(
                        current_paragraphs
                    )


            # ШИНЭ ХИЧЭЭЛ

            current_lesson = {

                "title": text,

                "content": []
            }


            lessons.append(
                current_lesson
            )


            current_paragraphs = []

            continue


        # =================================================
        # HEADING 1-ЭЭС ӨМНӨХ АГУУЛГЫГ АЛГАСНА
        # =================================================

        if current_lesson is None:

            continue


        # Тухайн хичээлийн paragraph

        current_paragraphs.append(
            paragraph
        )


    # =====================================================
    # СҮҮЛИЙН ХИЧЭЭЛ
    # =====================================================

    if current_lesson is not None:

        current_lesson["content"] = \
            parse_content(
                current_paragraphs
            )


    return lessons


# =========================================================
# WORD ФАЙЛУУДЫГ УНШИХ
# =========================================================

def load_books():

    books = []


    if not os.path.exists(LSSN_DIR):

        print()
        print("======================================")
        print("АЛДАА")
        print("Lsn folder олдсонгүй")
        print(LSSN_DIR)
        print("======================================")
        print()

        return books


    for filename in sorted(
        os.listdir(LSSN_DIR)
    ):


        # DOCX биш бол алгасна

        if not filename.lower().endswith(
            ".docx"
        ):

            continue


        filepath = os.path.join(
            LSSN_DIR,
            filename
        )


        try:

            lessons = read_word(
                filepath
            )


            books.append({

                "name":
                    os.path.splitext(
                        filename
                    )[0],

                "filename":
                    filename,

                "lessons":
                    lessons

            })


            print(
                f"[OK] {filename}"
            )

            print(
                f"     {len(lessons)} хичээл"
            )


        except Exception as e:

            print()
            print(
                f"[ERROR] {filename}"
            )

            print(e)

            print()


    return books


# =========================================================
# WORD ФАЙЛУУДЫГ НЭГ УДАА УНШИНА
# =========================================================

BOOKS = load_books()


# =========================================================
# HOME
# =========================================================

@app.route("/")
def index():

    return render_template(
        "index.html",
        books=BOOKS
    )


# =========================================================
# НЭГ WORD ФАЙЛ
# =========================================================

@app.route(
    "/book/<int:book_id>"
)
def book(book_id):

    if (
        book_id < 0
        or
        book_id >= len(BOOKS)
    ):

        abort(404)


    book_data = BOOKS[
        book_id
    ]


    return render_template(
        "book.html",

        book=book_data,

        book_id=book_id
    )


# =========================================================
# НЭГ ХИЧЭЭЛ
# =========================================================

@app.route(
    "/book/<int:book_id>/lesson/<int:lesson_id>"
)
def lesson(
    book_id,
    lesson_id
):

    if (
        book_id < 0
        or
        book_id >= len(BOOKS)
    ):

        abort(404)


    book_data = BOOKS[
        book_id
    ]


    if (
        lesson_id < 0
        or
        lesson_id >= len(
            book_data["lessons"]
        )
    ):

        abort(404)


    lesson_data = \
        book_data["lessons"][
            lesson_id
        ]


    return render_template(
        "lesson.html",

        book=book_data,

        lesson=lesson_data,

        book_id=book_id,

        lesson_id=lesson_id
    )


# =========================================================
# START
# =========================================================

if __name__ == "__main__":

    print()
    print(
        "=========================================="
    )

    print(
        "        IELTS LEARNING WEB"
    )

    print(
        "=========================================="
    )

    print()

    if not BOOKS:

        print(
            "Word файл олдсонгүй!"
        )

    else:

        print(
            "Уншигдсан файлууд:"
        )

        print()

        for book in BOOKS:

            print(
                f"📘 {book['name']}"
            )

            print(
                f"   Хичээл: "
                f"{len(book['lessons'])}"
            )

            print()

    print(
        "=========================================="
    )

    print(
        "http://127.0.0.1:5000"
    )

    print(
        "=========================================="
    )

    print()


    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )

