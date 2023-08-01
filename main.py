import os
import win32com.client as win32

def get_text_between_words(doc, word1, word2):
    # Получаем текст из документа между двумя указанными словами
    full_text = doc.Range().Text
    start_index = full_text.find(word1)
    end_index = full_text.find(word2, start_index)
    if start_index == -1 or end_index == -1:
        return ""
    return full_text[start_index:end_index]

def create_medical_files(input_file, output_folder=None):
    # Создаем COM объект для работы с Word
    word = win32.Dispatch("Word.Application")
    word.Visible = False

    # Если input_file указывает только название файла, добавляем путь к текущей папке проекта
    if not os.path.isabs(input_file):
        input_file = os.path.join(os.getcwd(), input_file)

    # Открываем входной файл
    doc = word.Documents.Open(input_file)

    # Получаем текст для файла "Жалобы_и_история_заболевания.docx" (до слова "Status")
    text_complaints_history = get_text_between_words(doc, "Anamnesis morbi", "Status")

    # Получаем текст для файла "Осмотр.docx" (от слова "Осмотр Терапевта" до слова "Дежурные")
    text_examination_part1 = get_text_between_words(doc, "Осмотр терапевта", "Дежурные")
    text_examination_part2 = get_text_between_words(doc, "Status", "ОАК")
    text_examination = text_examination_part1 + "\n\n" + text_examination_part2

    # Получаем текст для файла "Лабораторные_исследования.docx" (от слова "ОАК" до слова "УЗИ")
    text_lab_results = get_text_between_words(doc, "ОАК", "УЗИ")

    # Получаем текст для файла "Инструментальные_методы.docx" (от слова "УЗИ" до слова "Осмотр Терапевта")
    text_instrumental_methods = get_text_between_words(doc, "УЗИ", "Осмотр терапевта")

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
    word.Quit()

    print("Файлы успешно созданы.")

if __name__ == "__main__":
    input_file = "ИБциррПечСубкомпенс.doc"  # Укажите только название файла, если он уже лежит в папке проекта
    create_medical_files(input_file)
