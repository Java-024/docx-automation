import docx

# Загрузка документа
doc = docx.Document('document.docx')

# Чтение параграфов
full_text = []
for para in doc.paragraphs:
    full_text.append(para.text)

# Объединение текста
print('\n'.join(full_text))
