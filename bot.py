import os
import io
import re
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler

TOKEN = os.getenv("TELEGRAM_TOKEN")

WAITING_FOR_PAGES = 1

def iter_block_items(parent):
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("Something's not right")
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def parse_pages(pages_str: str) -> set:
    if pages_str.strip().lower() == "все":
        return "ALL"
    pages = set()
    parts = pages_str.replace(" ", "").split(",")
    for part in parts:
        if "-" in part:
            try:
                start, end = part.split("-")
                if start.isdigit() and end.isdigit():
                    pages.update(range(int(start), int(end) + 1))
            except ValueError:
                pass
        elif part.isdigit():
            pages.add(int(part))
    return pages

def create_element(name):
    return OxmlElement(name)

def create_attribute(element, name, value):
    element.set(qn(name), value)

def add_page_number(run):
    fldChar1 = create_element('w:fldChar')
    create_attribute(fldChar1, 'w:fldCharType', 'begin')
    instrText = create_element('w:instrText')
    create_attribute(instrText, 'xml:space', 'preserve')
    instrText.text = "PAGE"
    fldChar2 = create_element('w:fldChar')
    create_attribute(fldChar2, 'w:fldCharType', 'separate')
    fldChar3 = create_element('w:fldChar')
    create_attribute(fldChar3, 'w:fldCharType', 'end')
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)

def apply_gost_styles(doc: Document, target_pages="ALL"):
    if target_pages == "ALL":
        for section in doc.sections:
            section.top_margin = Cm(1.0)
            section.bottom_margin = Cm(1.0)
            section.left_margin = Cm(2.5)
            section.right_margin = Cm(1.0)
            
            header = section.header
            if not header.paragraphs:
                hp = header.add_paragraph()
            else:
                hp = header.paragraphs[0]
                hp.clear()
            hp.text = "НАЗВАНИЕ КОЛЛЕДЖА, СПЕЦИАЛЬНОСТЬ, ОБОЗНАЧЕНИЕ ДОКУМЕНТА"
            hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for run in hp.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(14)
                run.font.all_caps = True
                
            footer = section.footer
            if not footer.paragraphs:
                fp = footer.add_paragraph()
            else:
                fp = footer.paragraphs[0]
                fp.clear()
            fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run1 = fp.add_run("НАИМЕНОВАНИЕ ДОКУМЕНТА    ГРУППА    ")
            run1.font.name = 'Times New Roman'
            run1.font.size = Pt(14)
            run1.font.all_caps = True
            
            run2 = fp.add_run()
            run2.font.name = 'Times New Roman'
            run2.font.size = Pt(14)
            add_page_number(run2)

    page_num = 1
    for block in iter_block_items(doc):
        if target_pages == "ALL" or page_num in target_pages:
            if isinstance(block, Paragraph):
                text = block.text.strip()
                
                has_image = any(run._element.xpath('.//w:drawing') or run._element.xpath('.//w:pict') for run in block.runs)
                
                orig_align = block.alignment
                if orig_align is None and block.style and block.style.paragraph_format:
                    orig_align = block.style.paragraph_format.alignment
                    
                if has_image:
                    block.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    block.paragraph_format.first_line_indent = Cm(0)
                    block.paragraph_format.line_spacing = 1.0
                    block.paragraph_format.keep_with_next = True
                elif text.startswith("Рисунок"):
                    block.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    block.paragraph_format.first_line_indent = Cm(0)
                    block.paragraph_format.space_after = Pt(6)
                    block.paragraph_format.line_spacing = 1.0
                elif text.startswith("Таблица"):
                    block.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    block.paragraph_format.first_line_indent = Cm(0)
                    block.paragraph_format.space_after = Pt(0)
                    block.paragraph_format.line_spacing = 1.0
                elif orig_align in (WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT):
                    # Preserve Title pages and center headings, remove indent
                    block.paragraph_format.first_line_indent = Cm(0)
                    block.paragraph_format.line_spacing = 1.5
                else:
                    # Normal GOST text
                    block.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    block.paragraph_format.line_spacing = 1.5
                    block.paragraph_format.first_line_indent = Cm(1.25)
                
                for run in block.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(14)
                    
            elif isinstance(block, Table):
                block.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for row_idx, row in enumerate(block.rows):
                    is_header = (row_idx == 0)
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            if is_header:
                                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            else:
                                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            para.paragraph_format.line_spacing = 1.0
                            para.paragraph_format.first_line_indent = Cm(0)
                            for run in para.runs:
                                run.font.name = 'Times New Roman'
                                run.font.size = Pt(12)
                                if is_header:
                                    run.font.bold = True
        
        for elem in block._element.iter():
            if elem.tag.endswith('lastRenderedPageBreak'):
                page_num += 1
            elif elem.tag.endswith('br') and elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type') == 'page':
                page_num += 1

    # Изменение размеров всех картинок (чтобы не выходили за поля и занимали не более 1/2 страницы)
    if target_pages == "ALL":
        max_w = Cm(17.0)
        max_h = Cm(13.0)
        for shape in doc.inline_shapes:
            if shape.width and shape.height:
                if shape.width > max_w or shape.height > max_h:
                    ratio_w = max_w / shape.width
                    ratio_h = max_h / shape.height
                    ratio = min(ratio_w, ratio_h)
                    if ratio < 1.0:
                        shape.width = int(shape.width * ratio)
                        shape.height = int(shape.height * ratio)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет! Я бот для форматирования документов по ГОСТ. 📄\n\n"
        "Вы можете:\n"
        "1. Отправить мне текст обычным сообщением, и я соберу из него Word-документ.\n"
        "2. Отправить файл `.docx`."
    )
    return ConversationHandler.END

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    doc = Document()
    for paragraph in text.split('\n'):
        if paragraph.strip():
            doc.add_paragraph(paragraph.strip())
            
    apply_gost_styles(doc, "ALL")
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    await update.message.reply_document(
        document=buffer,
        filename="ГОСТ_Документ.docx",
        caption="Ваш текст успешно структурирован и оформлен по ГОСТу! ✅"
    )
    return ConversationHandler.END

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.message.document.file_name.endswith('.docx'):
        await update.message.reply_text("Пожалуйста, отправьте файл в формате .docx (Word).")
        return ConversationHandler.END
        
    doc_file = await update.message.document.get_file()
    context.user_data['file_bytes'] = await doc_file.download_as_bytearray()
    context.user_data['file_name'] = update.message.document.file_name
    
    reply_markup = ReplyKeyboardMarkup([["Все"]], resize_keyboard=True, one_time_keyboard=True)
    await update.message.reply_text(
        "К каким страницам применить форматирование ГОСТ (шрифты, абзацы)?\n\n"
        "Напишите номера страниц через запятую или дефис (например: `1, 3, 5-10`).\n"
        "Либо нажмите «Все», чтобы отформатировать весь документ целиком.\n\n"
        "*(Если вы укажете конкретные страницы, бот изменит только текст на них. Поля документа и колонтитулы останутся нетронутыми!)*.",
        parse_mode='Markdown',
        reply_markup=reply_markup
    )
    return WAITING_FOR_PAGES

async def process_document_pages(update: Update, context: ContextTypes.DEFAULT_TYPE):
    pages_str = update.message.text
    target_pages = parse_pages(pages_str)
    
    if not target_pages:
        await update.message.reply_text("Не удалось распознать номера страниц. Попробуйте еще раз (например: 1, 3-5) или напишите 'Все'.")
        return WAITING_FOR_PAGES
        
    status_message = await update.message.reply_text("⏳ Обрабатываю документ...", reply_markup=ReplyKeyboardRemove())
    
    try:
        file_bytes = context.user_data['file_bytes']
        doc = Document(io.BytesIO(file_bytes))
        
        apply_gost_styles(doc, target_pages)
        
        output_buffer = io.BytesIO()
        doc.save(output_buffer)
        output_buffer.seek(0)
        
        caption_text = "Ваш документ отформатирован! ✅"
        if target_pages != "ALL":
            caption_text += "\nШрифты и абзацы изменены только на выбранных страницах."
            
        await update.message.reply_document(
            document=output_buffer,
            filename=f"ГОСТ_{context.user_data['file_name']}",
            caption=caption_text
        )
        
        await status_message.delete()
        
    except Exception as e:
        await update.message.reply_text(f"❌ Произошла ошибка при обработке файла.")
        print(f"Error: {e}")
        
    context.user_data.clear()
    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Действие отменено.", reply_markup=ReplyKeyboardRemove())
    context.user_data.clear()
    return ConversationHandler.END

def main():
    app = Application.builder().token(TOKEN).build()
    
    conv_handler = ConversationHandler(
        entry_points=[MessageHandler(filters.Document.ALL, handle_document)],
        states={
            WAITING_FOR_PAGES: [MessageHandler(filters.TEXT & ~filters.COMMAND, process_document_pages)],
        },
        fallbacks=[CommandHandler("cancel", cancel)]
    )
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(conv_handler)
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    
    print("Бот запущен. Нажмите Ctrl+C для остановки.")
    app.run_polling()

if __name__ == "__main__":
    main()
