from docx import Document


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

    special_chars = {'"', '*', '=', '\\', '”', '“'}  # Добавил разные виды кавычек
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

        # Функция для получения форматирования по позиции
        def get_formatting_at_pos(pos):
            """Получает форматирование символа в позиции pos"""
            char_pos = 0
            for r in runs_info:
                text_len = len(r['text'])
                if char_pos <= pos < char_pos + text_len:
                    return r['bold'], r['italic']
                char_pos += text_len
            return False, False

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
                    current_part_text = ''

                    while j < len(full_text):
                        char = full_text[j]
                        bold, italic = get_formatting_at_pos(j)

                        # Проверяем, является ли символ закрывающим для текущей части
                        if char == start_char:
                            # Нашли закрывающий символ
                            if current_part_text:
                                current_part['content'] = current_part_text
                            current_part['end_char'] = start_char
                            structure_parts.append(current_part.copy())

                            # Формируем структуру для вывода
                            struct_full_text = f"#{start_char}"
                            for idx, part in enumerate(structure_parts):
                                struct_full_text += part['content']
                                if idx < len(structure_parts) - 1:
                                    struct_full_text += '&'
                            struct_full_text += start_char

                            # Собираем информацию о форматировании каждой части
                            parts_info = []
                            for idx, part in enumerate(structure_parts):
                                # Определяем ключ для этой части
                                if idx == 0:
                                    key = f"#{part['start_char']}{part['content']}{part['end_char']}"
                                else:
                                    key = f"&{part['content']}{part['end_char']}"

                                part_info = {
                                    'key': key,
                                    'is_bold': False,
                                    'is_italic': False
                                }

                                # Проверяем, есть ли содержимое
                                if part['content'] and len(part['content']) > 0:
                                    # Ищем первую букву или значимый символ в содержимом
                                    first_significant_pos = None
                                    for letter_idx, letter in enumerate(part['content']):
                                        if letter.isalpha() or letter.isdigit():
                                            first_significant_pos = letter_idx
                                            break

                                    if first_significant_pos is not None:
                                        # Находим позицию первого значимого символа в общем тексте
                                        global_pos = i + 2  # Позиция после # и start_char

                                        # Добавляем содержимое предыдущих частей
                                        for prev_part in structure_parts[:idx]:
                                            global_pos += len(prev_part['content'])
                                            if prev_part != structure_parts[idx - 1] or idx > 0:
                                                global_pos += 1  # +1 для &

                                        # Добавляем позицию первого символа в текущей части
                                        global_pos += first_significant_pos

                                        # Получаем форматирование
                                        if global_pos < len(full_text):
                                            bold, italic = get_formatting_at_pos(global_pos)
                                            part_info['is_bold'] = bold
                                            part_info['is_italic'] = italic

                                parts_info.append(part_info)

                            # Формируем результат для этой структуры
                            structure_result = [struct_full_text]
                            for part_info in parts_info:
                                structure_result.append(
                                    [part_info['key'], part_info['is_bold'], part_info['is_italic']])

                            # Проверяем, не пустая ли структура (игнорируем если нет содержимого)
                            has_content = False
                            for part in structure_parts:
                                if part['content']:
                                    has_content = True
                                    break

                            if has_content:
                                all_structures.append(structure_result)

                            i = j
                            break

                        elif char == '&':
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
                            j += 1

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
            print()  # Пустая строка для разделения структур
    else:
        print("Структуры не найдены.")

    return all_structures


if __name__ == "__main__":
    # Здесь укажи путь к твоему файлу
    file_path = "test_search.docx"  # <-- ЗДЕСЬ УКАЖИ ПУТЬ

    result = find_structures(file_path)