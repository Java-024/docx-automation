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

    special_chars = {'"', '*', '=', '\\', '”', '“', '"', "'"}  # Все возможные кавычки
    all_structures = []

    # Проходим по всем параграфам
    for paragraph in doc.paragraphs:
        # Собираем все runs с форматированием
        runs_info = []
        for run in paragraph.runs:
            text = run.text
            if text:
                is_bold, is_italic = get_run_formatting(run)
                runs_info.append({
                    'text': text,
                    'bold': is_bold,
                    'italic': is_italic
                })

        # Полный текст параграфа
        full_text = ''.join([r['text'] for r in runs_info])

        # Функция для получения форматирования по позиции
        def get_formatting_at_pos(pos):
            char_pos = 0
            for r in runs_info:
                text_len = len(r['text'])
                if char_pos <= pos < char_pos + text_len:
                    return r['bold'], r['italic']
                char_pos += text_len
            return False, False

        # Ищем структуры
        i = 0
        while i < len(full_text):
            if full_text[i] == '#':
                # Проверяем следующий символ
                if i + 1 < len(full_text) and full_text[i + 1] in special_chars:
                    start_char = full_text[i + 1]
                    end_pos = -1

                    # Ищем конец структуры
                    depth = 0
                    j = i + 2
                    structure_text = ''
                    parts = []  # Список частей структуры
                    current_part = ''
                    in_part = True
                    part_starts = []  # Позиции начала каждой части

                    while j < len(full_text):
                        char = full_text[j]

                        if char == '&' and depth == 0:
                            # Сохраняем текущую часть
                            if current_part:
                                parts.append(current_part)
                                current_part = ''
                            j += 1
                            continue

                        if char == start_char and depth == 0:
                            # Нашли закрывающий символ
                            if current_part:
                                parts.append(current_part)
                            end_pos = j

                            # Формируем полную структуру
                            full_structure = f"#{start_char}"
                            for idx, part in enumerate(parts):
                                full_structure += part
                                if idx < len(parts) - 1:
                                    full_structure += '&'
                            full_structure += start_char

                            # Собираем информацию о форматировании для каждой части
                            parts_info = []

                            # Позиция начала первой части (после # и start_char)
                            current_pos = i + 2

                            for idx, part in enumerate(parts):
                                # Определяем ключ части
                                if idx == 0:
                                    key = f"#{start_char}{part}{start_char}"
                                else:
                                    key = f"&{part}{start_char}"

                                part_info = [key, False, False]

                                # Ищем первый значимый символ в части
                                first_char_pos = -1
                                for char_idx, char in enumerate(part):
                                    if char.isalpha() or char.isdigit():
                                        first_char_pos = char_idx
                                        break

                                if first_char_pos != -1:
                                    # Вычисляем глобальную позицию первого символа
                                    global_pos = current_pos + first_char_pos
                                    bold, italic = get_formatting_at_pos(global_pos)
                                    part_info[1] = bold
                                    part_info[2] = italic

                                parts_info.append(part_info)

                                # Обновляем позицию для следующей части
                                current_pos += len(part) + 1  # +1 для & или закрывающего символа

                            # Добавляем структуру в результат
                            result_struct = [full_structure]
                            result_struct.extend(parts_info)
                            all_structures.append(result_struct)

                            i = end_pos
                            break
                        else:
                            current_part += char
                            j += 1

                    if end_pos == -1:
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
            print()
    else:
        print("Структуры не найдены.")

    return all_structures


if __name__ == "__main__":
    # Здесь укажи путь к твоему файлу
    file_path = "test_search.docx"  # <-- ЗДЕСЬ УКАЖИ ПУТЬ

    result = find_structures(file_path)