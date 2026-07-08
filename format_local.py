import sys
from bot import apply_gost_styles
from docx import Document

try:
    input_path = sys.argv[1]
    output_path = input_path.replace('.docx', '_ГОСТ.docx')

    doc = Document(input_path)
    apply_gost_styles(doc, "ALL")
    doc.save(output_path)
    print(f"Файл успешно отформатирован и сохранен как:\n{output_path}")
except Exception as e:
    print(f"Ошибка: {e}")
