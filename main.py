import os
import win32com.client as win32

def get_text_between_words(doc, word1, word2, add_newline_phrases=None):
    # Функция для получения текста из документа между двумя указанными словами
    full_text = doc.Range().Text
    start_index = full_text.find(word1)
    end_index = full_text.find(word2, start_index)
    if start_index == -1 or end_index == -1:
        return ""

    text = full_text[start_index:end_index].strip()

    if add_newline_phrases:
        # Разделяем указанные записи на новые строки
        for phrase in add_newline_phrases:
            text = text.replace(phrase, "\n" + phrase)

    return text

def add_newlines_to_keywords(text):
    # Функция для добавления переносов строк перед ключевыми словами
    keywords = ["эффект", "терапи", "лечени", "диагноз", "осложненн", "степени тяжести",
                "степень тяжести", "госпитализаци", "исход", "Anamnesis vitae",
                "заболевани", "Травм", "ктомия", "осложнени", "препарат", "учет",
                "получает", "Аллергологическ"]
    for keyword in keywords:
        text = text.replace(keyword, "\n" + keyword)
    return text

def create_medical_files(input_file, output_folder=None):
    # Создаем COM объект для работы с Word
    word = win32.Dispatch("Word.Application")
    word.Visible = False

    # Если input_file указывает только название файла, добавляем путь к текущей папке проекта
    if not os.path.isabs(input_file):
        input_file = os.path.join(os.getcwd(), input_file)

    # Открываем входной файл
    doc = word.Documents.Open(input_file)

    # Получаем текст для файла "Жалобы_и_история_заболевания.docx" (от слова "Пациентка" до "Status")
    add_newline_phrases = ["общую слабость", "головокружение", "увеличение объема живота",
                           "желтушность кожных покровов",
                           "приема жирной пищи", "физических нагрузок", "утомляемость", "гепатоспленомегалия",
                           "гепатопрокторами ", "Гиперспленизм", "цирроз печени субкомпенсированный", "Болезнь Боткина",
                           "туберкулез", "кожно-венерические", "аппендэктомия", "геморроидэктомия", "гемотрансфузия", ]
    text_complaints_history_part1 = get_text_between_words(doc, "Пациентка", "Anamnesis", add_newline_phrases)
    text_complaints_history_part2 = get_text_between_words(doc, "Anamnesis", "Status", add_newline_phrases)
    text_complaints_history = text_complaints_history_part1 + "\n\n" + text_complaints_history_part2

    # Добавляем переносы строк перед ключевыми словами в файле "Жалобы_и_история_заболевания.docx"
    text_complaints_history = add_newlines_to_keywords(text_complaints_history)

    # Получаем текст для файла "Опр-ИБциррПечСубкомпенс.doc" (от слова "Anamnesis vitae:" до конца файла)
    start_index_anamnesis_vitae = text_complaints_history.find("Anamnesis vitae:")
    if start_index_anamnesis_vitae != -1:
        text_anamnesis_vitae = text_complaints_history[start_index_anamnesis_vitae:]
        text_complaints_history = text_complaints_history[:start_index_anamnesis_vitae]

        # Если output_folder не указан, используем папку проекта
        if output_folder is None:
            output_folder = os.getcwd()

        # Создаем файл "Опр-ИБциррПечСубкомпенс.doc" и записываем туда вырезанный текст
        output_filename_anamnesis_vitae = os.path.join(output_folder, "Опр-ИБциррПечСубкомпенс.doc")
        doc_anamnesis_vitae = word.Documents.Add()
        doc_anamnesis_vitae.Range().Text = text_anamnesis_vitae
        doc_anamnesis_vitae.SaveAs(output_filename_anamnesis_vitae)

    # Получаем текст для файла "Осмотр.docx" (от слова "Осмотр Терапевта" до слова "Дежурные")
    text_examination_part1 = get_text_between_words(doc, "Осмотр Терапевта", "Дежурные")
    text_examination_part2 = get_text_between_words(doc, "Status", "ОАК")
    text_examination = text_examination_part1 + "\n\n" + text_examination_part2

    # Получаем текст для файла "Лабораторные_исследования.docx" (от слова "ОАК" до слова "УЗИ")
    text_lab_results = get_text_between_words(doc, "ОАК", "УЗИ")

    # Получаем текст для файла "Инструментальные_методы.docx" (от слова "УЗИ" до слова "Осмотр Терапевта")
    text_instrumental_methods = get_text_between_words(doc, "УЗИ", "Осмотр Терапевта")

    # Если output_folder не указан, используем папку проекта
    if output_folder is None:
        output_folder = os.getcwd()

    # Создаем файл "Жалобы_и_история_заболевания.docx" и записываем туда текст
    output_filename_complaints = os.path.join(output_folder, "Жалобы_и_история_заболевания.docx")
    doc_complaints = word.Documents.Add()
    doc_complaints.Range().Text = text_complaints_history
    doc_complaints.SaveAs(output_filename_complaints)

    # Создаем файл "Осмотр.docx" и записываем туда текст
    output_filename_examination = os.path.join(output_folder, "Осмотр.docx")
    doc_examination = word.Documents.Add()
    doc_examination.Range().Text = text_examination
    doc_examination.SaveAs(output_filename_examination)

    # Создаем файл "Лабораторные_исследования.docx" и записываем туда текст
    output_filename_lab_results = os.path.join(output_folder, "Лабораторные_исследования.docx")
    doc_lab_results = word.Documents.Add()
    doc_lab_results.Range().Text = text_lab_results
    doc_lab_results.SaveAs(output_filename_lab_results)

    # Создаем файл "Инструментальные_методы.docx" и записываем туда текст
    output_filename_instrumental_methods = os.path.join(output_folder, "Инструментальные_методы.docx")
    doc_instrumental_methods = word.Documents.Add()
    doc_instrumental_methods.Range().Text = text_instrumental_methods
    doc_instrumental_methods.SaveAs(output_filename_instrumental_methods)

    # Закрываем документы и завершаем работу с Word
    doc.Close()
    doc_complaints.Close()
    doc_examination.Close()
    doc_lab_results.Close()
    doc_instrumental_methods.Close()
    if start_index_anamnesis_vitae != -1:
        doc_anamnesis_vitae.Close()
    word.Quit()

    print("Файлы успешно созданы.")

if __name__ == "__main__":
    input_file = "ИБциррПечСубкомпенс.doc"  # Укажите только название файла, если он уже лежит в папке проекта
    create_medical_files(input_file)