from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import parse_xml


def get_run_formatting(run):
    """Получает форматирование run: жирный и курсив"""
    is_bold = run.bold if run.bold is not None else False
    is_italic = run.italic if run.italic is not None else False
    return is_bold, is_italic


def find_structures(file_path):
    """
    Поиск структур в документе Word, начинающихся с # и одного из символов: ", *, =, \
    и заканчивающихся тем же символом. & используется для объединения структур.
    """
    try:
        doc = Document(file_path)
    except Exception as e:
        print(f"Ошибка при открытии файла: {e}")
        return []

    special_chars = {'"', '*', '=', '\\'}
    all_structures = []  # Главный список для всех структур

    # Проходим по всем параграфам
    for paragraph in doc.paragraphs:
        # Собираем все runs в параграфе с их форматированием
        runs_info = []
        for run in paragraph.runs:
            text = run.text
            if text:
                is_bold, is_italic = get_run_formatting(run)
                runs_info.append({
                    'text': text,
                    'bold': is_bold,
                    'italic': is_italic,
                    'run': run
                })

        # Сканируем текст для поиска структур
        full_text = ''.join([r['text'] for r in runs_info])

        i = 0
        while i < len(full_text):
            if full_text[i] == '#':
                if i + 1 < len(full_text) and full_text[i + 1] in special_chars:
                    start_char = full_text[i + 1]
                    structure_parts = []  # Список для частей структуры
                    current_part = {
                        'full_text': '#',
                        'start_char': start_char,
                        'content': '',
                        'end_char': '',
                        'bold': False,
                        'italic': False
                    }

                    j = i + 2
                    depth = 0  # Глубина вложенности для &
                    current_part_text = ''
                    part_start_pos = i

                    # Собираем информацию о форматировании для текущей позиции
                    def get_formatting_at_pos(pos):
                        """Получает форматирование символа в позиции pos"""
                        char_pos = 0
                        for r in runs_info:
                            text_len = len(r['text'])
                            if char_pos <= pos < char_pos + text_len:
                                return r['bold'], r['italic']
                            char_pos += text_len
                        return False, False

                    while j < len(full_text):
                        char = full_text[j]
                        bold, italic = get_formatting_at_pos(j)

                        if char == '&':
                            # Сохраняем текущую часть
                            if current_part_text:
                                current_part['content'] = current_part_text
                                structure_parts.append(current_part.copy())
                                current_part_text = ''

                            # Начинаем новую часть
                            current_part = {
                                'full_text': '#',
                                'start_char': start_char,
                                'content': '',
                                'end_char': '',
                                'bold': False,
                                'italic': False
                            }
                            depth += 1
                            j += 1

                        elif char == start_char and depth == 0:
                            # Нашли закрывающий символ
                            if current_part_text:
                                current_part['content'] = current_part_text
                            current_part['end_char'] = start_char
                            structure_parts.append(current_part.copy())

                            # Формируем структуру для вывода
                            struct_full_text = f"#{start_char}"
                            for part in structure_parts:
                                struct_full_text += part['content']
                                if part != structure_parts[-1]:
                                    struct_full_text += '&'
                            struct_full_text += start_char

                            # Собираем информацию о форматировании каждой части
                            parts_info = []
                            for idx, part in enumerate(structure_parts):
                                # Получаем форматирование для начала части
                                part_start_pos_in_doc = 0
                                # Находим позицию начала этой части в runs
                                part_info = {
                                    'key': f"#{part['start_char']}{part['content']}{part['end_char']}" if idx == 0 else f"{part['content']}{part['end_char']}",
                                    'is_bold': False,
                                    'is_italic': False
                                }

                                # Проверяем, есть ли содержимое
                                if part['content'] and len(part['content']) > 0:
                                    # Ищем первую букву содержимого
                                    first_letter_pos = None
                                    for letter_idx, letter in enumerate(part['content']):
                                        if letter.isalpha():
                                            first_letter_pos = letter_idx
                                            break

                                    if first_letter_pos is not None:
                                        # Находим позицию первой буквы в общем тексте
                                        char_pos = 0
                                        if idx == 0:
                                            # Для первой части: # + start_char + позиция первой буквы
                                            pos_in_struct = 2 + first_letter_pos
                                        else:
                                            # Для последующих частей: позиция после & + позиция первой буквы
                                            pos_in_struct = first_letter_pos

                                        # Находим эту позицию в runs
                                        global_pos = i
                                        if idx == 0:
                                            global_pos = i + 2 + first_letter_pos
                                        else:
                                            # Нужно найти позицию после всех предыдущих частей и &
                                            global_pos = i + 2
                                            for prev_part in structure_parts[:idx]:
                                                global_pos += len(prev_part['content']) + 1  # +1 для &
                                            global_pos += first_letter_pos

                                        bold, italic = get_formatting_at_pos(global_pos)
                                        part_info['is_bold'] = bold
                                        part_info['is_italic'] = italic

                                parts_info.append(part_info)

                            # Формируем результат для этой структуры
                            structure_result = [struct_full_text]
                            for part_info in parts_info:
                                structure_result.append(
                                    [part_info['key'], part_info['is_bold'], part_info['is_italic']])

                            all_structures.append(structure_result)

                            i = j
                            break

                        else:
                            current_part_text += char
                            j += 1
                    else:
                        i += 1
                else:
                    i += 1
            else:
                i += 1

    # Выводим результаты
    if all_structures:
        print("Найденные структуры:")
        for struct in all_structures:
            print(struct)
    else:
        print("Структуры не найдены.")

    return all_structures


if __name__ == "__main__":
    # Здесь укажи путь к твоему файлу
    file_path = "test_search.docx"  # <-- ЗДЕСЬ УКАЖИ ПУТЬ

    result = find_structures(file_path)