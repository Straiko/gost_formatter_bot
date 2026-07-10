import os
import sys
import time
import traceback

# Catch absolutely everything at startup
try:
    import io
    import re
    import json
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

    import telebot
    from telebot.types import ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove
    
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except Exception:
        pass

    TOKEN = os.getenv("TELEGRAM_TOKEN")
    if not TOKEN:
        TOKEN = os.getenv("BOT_TOKEN")
    
    if not TOKEN:
        print("\n" + "="*50, flush=True)
        print("ОШИБКА: В настройках Bothost не найдена переменная TELEGRAM_TOKEN!", flush=True)
        print("Пожалуйста, добавьте её в разделе 'Переменные окружения'.", flush=True)
        print("="*50 + "\n", flush=True)
        while True:
            time.sleep(60)

    bot = telebot.TeleBot(TOKEN)
    user_data = {}

except Exception as e:
    print("\n" + "="*50, flush=True)
    print("КРИТИЧЕСКАЯ ОШИБКА ПРИ ЗАПУСКЕ БОТА:", flush=True)
    traceback.print_exc()
    print("="*50 + "\n", flush=True)
    while True:
        time.sleep(60)


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
    try:
        if doc.part.numbering_part is not None:
            numbering = doc.part.numbering_part.numbering_definitions._numbering
            for abstract in numbering.findall(qn('w:abstractNum')):
                for lvl in abstract.findall(qn('w:lvl')):
                    lvl_pPr = lvl.find(qn('w:pPr'))
                    if lvl_pPr is None:
                        lvl_pPr = OxmlElement('w:pPr')
                        lvl.append(lvl_pPr)
                    ind = lvl_pPr.find(qn('w:ind'))
                    if ind is None:
                        ind = OxmlElement('w:ind')
                        lvl_pPr.append(ind)
                    ind.set(qn('w:left'), '0')
                    ind.set(qn('w:hanging'), '0')
                    
                    numFmt = lvl.find(qn('w:numFmt'))
                    if numFmt is not None and numFmt.get(qn('w:val')) == 'bullet':
                        lvlText = lvl.find(qn('w:lvlText'))
                        if lvlText is None:
                            lvlText = OxmlElement('w:lvlText')
                            lvl.insert(0, lvlText)
                        lvlText.set(qn('w:val'), '–')
    except Exception as e:
        print(f"Error modifying numbering definitions: {e}")

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
                block.paragraph_format.space_before = Pt(0)
                block.paragraph_format.space_after = Pt(0)
                
                has_image = any(run._element.xpath('.//w:drawing') or run._element.xpath('.//w:pict') for run in block.runs)
                
                orig_align = block.alignment
                if orig_align is None and block.style and block.style.paragraph_format:
                    orig_align = block.style.paragraph_format.alignment
                    
                is_figure_caption = False
                is_text_heading = False
                is_normal_text = False
                
                is_list = False
                if block.style and 'List' in block.style.name:
                    is_list = True
                elif block._element.pPr is not None and block._element.pPr.numPr is not None:
                    is_list = True
                
                is_heading_style = block.style and 'Heading' in block.style.name
                
                is_fig_by_style = block.style and 'рисунок' in block.style.name.lower()
                if is_fig_by_style and len(text) > 150:
                    is_fig_by_style = False
                is_fig_by_text = text.strip().startswith('Рисунок ') and ' - ' in text
                
                is_table_by_text = text.strip().lower().startswith('таблица')
                
                if has_image:
                    block.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    block.paragraph_format.first_line_indent = Cm(0)
                    block.paragraph_format.line_spacing = 1.0
                    block.paragraph_format.keep_with_next = True
                elif is_fig_by_style or is_fig_by_text or ("Рисунок" in text and any('SEQ' in (n.text or '') for r in block.runs for n in r._element.xpath('.//w:instrText'))):
                    is_figure_caption = True
                    block.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    block.paragraph_format.first_line_indent = Cm(0)
                    block.paragraph_format.space_after = Pt(6)
                    block.paragraph_format.line_spacing = 1.0
                elif is_table_by_text or ("таблица" in text.lower() and any('SEQ' in (n.text or '') for r in block.runs for n in r._element.xpath('.//w:instrText'))):
                    block.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    block.paragraph_format.first_line_indent = Cm(0)
                    block.paragraph_format.left_indent = Cm(0)
                    block.paragraph_format.space_after = Pt(0)
                    block.paragraph_format.line_spacing = 1.0
                elif is_heading_style or orig_align in (WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT):
                    if is_heading_style or orig_align == WD_ALIGN_PARAGRAPH.CENTER:
                        is_text_heading = True
                        block.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    block.paragraph_format.first_line_indent = Cm(0)
                    block.paragraph_format.line_spacing = 1.5
                else:
                    is_normal_text = True
                    block.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    block.paragraph_format.line_spacing = 1.5
                    if not is_list:
                        block.paragraph_format.left_indent = Cm(0)
                        block.paragraph_format.first_line_indent = Cm(1.25)
                    else:
                        pPr_list = block._p.get_or_add_pPr()
                        for old_ind in pPr_list.findall(qn('w:ind')):
                            pPr_list.remove(old_ind)
                        ind_el = OxmlElement('w:ind')
                        ind_el.set(qn('w:left'), '0')
                        ind_el.set(qn('w:firstLine'), '709')
                        pPr_list.append(ind_el)
                
                for run in block.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(14)
                    if is_figure_caption:
                        run.font.bold = False
                        run.font.italic = False
                    elif is_text_heading:
                        run.font.bold = True
                        run.font.italic = False
                    elif is_normal_text:
                        run.font.bold = False
                        run.font.italic = False
                    
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

@bot.message_handler(commands=['start'])
def send_welcome(message):
    bot.reply_to(message, "Привет! Я бот для форматирования документов по ГОСТ. 📄\n\nВы можете:\n1. Отправить мне текст обычным сообщением, и я соберу из него Word-документ.\n2. Отправить файл `.docx`.", reply_markup=ReplyKeyboardRemove())

@bot.message_handler(commands=['cancel'])
def cancel(message):
    if message.chat.id in user_data:
        del user_data[message.chat.id]
    bot.reply_to(message, "Действие отменено.", reply_markup=ReplyKeyboardRemove())

@bot.message_handler(content_types=['text'])
def handle_text(message):
    if message.chat.id in user_data and user_data[message.chat.id].get('state'):
        return
    text = message.text
    doc = Document()
    for paragraph in text.split('\n'):
        if paragraph.strip():
            doc.add_paragraph(paragraph.strip())
            
    apply_gost_styles(doc, "ALL")
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    bot.send_document(
        chat_id=message.chat.id,
        document=( "ГОСТ_Документ.docx", buffer ),
        caption="Ваш текст успешно структурирован и оформлен по ГОСТу! ✅"
    )

@bot.message_handler(content_types=['document'])
def handle_document(message):
    if not message.document.file_name.endswith('.docx'):
        bot.reply_to(message, "Пожалуйста, отправьте файл в формате .docx (Word).")
        return
        
    file_info = bot.get_file(message.document.file_id)
    downloaded_file = bot.download_file(file_info.file_path)
    
    user_data[message.chat.id] = {
        'file_bytes': downloaded_file,
        'file_name': message.document.file_name,
        'state': 'WAITING_FOR_PAGES'
    }
    
    markup = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    markup.add(KeyboardButton("Все"))
    bot.reply_to(
        message,
        "К каким страницам применить форматирование ГОСТ (шрифты, абзацы)?\n\n"
        "Напишите номера страниц через запятую или дефис (например: `1, 3, 5-10`).\n"
        "Либо нажмите «Все», чтобы отформатировать весь документ целиком.\n\n"
        "*(Если вы укажете конкретные страницы, бот изменит только текст на них. Поля документа и колонтитулы останутся нетронутыми!)*.",
        parse_mode='Markdown',
        reply_markup=markup
    )
    bot.register_next_step_handler(message, process_document_pages)

def process_document_pages(message):
    if message.text == '/cancel':
        return cancel(message)
    pages_str = message.text
    target_pages = parse_pages(pages_str)
    
    if not target_pages:
        bot.reply_to(message, "Не удалось распознать номера страниц. Попробуйте еще раз (например: 1, 3-5) или напишите 'Все'.")
        bot.register_next_step_handler(message, process_document_pages)
        return
        
    user_data[message.chat.id]['target_pages'] = target_pages
    
    markup = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    markup.add(KeyboardButton("0"), KeyboardButton("1"), KeyboardButton("2"), KeyboardButton("3"))
    bot.reply_to(
        message,
        "Хотите ли вы, чтобы ИИ проанализировал текст и автоматически сгенерировал списки?\n"
        "Укажите максимальное количество списков на главу (например: 0 — не нужно, 1, 2...)",
        reply_markup=markup
    )
    bot.register_next_step_handler(message, process_ai_lists)

def process_ai_lists(message):
    if message.text == '/cancel':
        return cancel(message)
    try:
        lists_per_chapter = int(message.text)
    except ValueError:
        bot.reply_to(message, "Пожалуйста, введите число (0, 1, 2...)")
        bot.register_next_step_handler(message, process_ai_lists)
        return

    status_message = bot.reply_to(message, "⏳ Обрабатываю документ (это может занять около минуты)...", reply_markup=ReplyKeyboardRemove())
    
    try:
        file_bytes = user_data[message.chat.id]['file_bytes']
        doc = Document(io.BytesIO(file_bytes))
        target_pages = user_data[message.chat.id].get('target_pages', 'ALL')
        
        if lists_per_chapter > 0:
            from openai import OpenAI
            client = OpenAI(
                api_key=os.getenv("PROXYAPI_KEY"),
                base_url="https://api.proxyapi.ru/openai/v1"
            )
            
            text_blocks = [p.text.strip() for p in doc.paragraphs if len(p.text.strip()) > 15]
            chunk_size = 60
            for i in range(0, len(text_blocks), chunk_size):
                chunk = "\\n\\n".join(text_blocks[i:i+chunk_size])
                prompt = f"""You are an expert editor formatting Russian text according to GOST. 
Here is a text chunk. Your task is to find places where formatting should be fixed. 
This includes:
1) Lists: Sequences of paragraphs that are clearly list items but missing bullet points. Add dashes ("- ") or numbers ("1) "). 
2) Headings: You MUST identify ALL headings and wrap them in <HEADING> and </HEADING> tags.

CRITICAL RULES:
- "original_paragraphs" MUST contain EVERY SINGLE PARAGRAPH from the original text that you are replacing.
- "new_paragraphs" MUST contain only the transformed versions of the exact text in "original_paragraphs". DO NOT hallucinate.
- The introductory sentence before a list MUST end with a COLON (:).
- For dashed lists ("- "): The text of every item MUST start with a LOWERCASE letter. Every item MUST end with a semicolon (";"), EXCEPT the very last item which MUST end with a period (".").

Return ONLY a JSON object with a single key "replacements", which is an array of objects.
"original_paragraphs": ["Exact text 1", "Exact text 2"],
"new_paragraphs": ["Transformed text 1", "Transformed text 2"]

If no paragraphs need conversion, return {{"replacements": []}}.
Text chunk:
{chunk}"""
                try:
                    response = client.chat.completions.create(
                        model="gpt-4o-mini",
                        messages=[{"role": "user", "content": prompt}],
                        response_format={"type": "json_object"}
                    )
                    data = json.loads(response.choices[0].message.content)
                    replacements = data.get("replacements", [])
                    
                    for rep in replacements:
                        orig_paras = rep.get("original_paragraphs", [])
                        if "original_paragraph" in rep and not orig_paras:
                            orig_paras = [rep["original_paragraph"]]
                            
                        new_paras = rep.get("new_paragraphs", [])
                        if not orig_paras or not new_paras: continue
                        
                        orig_paras = [p.strip() for p in orig_paras if p.strip()]

                        def normalize(t):
                            return re.sub(r'\W+', '', t.lower())

                        all_paras = doc.paragraphs
                        for idx in range(len(all_paras) - len(orig_paras) + 1):
                            match = True
                            for j in range(len(orig_paras)):
                                actual_norm = normalize(all_paras[idx+j].text)
                                orig_norm   = normalize(orig_paras[j])
                                key = orig_norm[:60]
                                if not key:
                                    match = False; break
                                if actual_norm != orig_norm and not actual_norm.startswith(key):
                                    match = False
                                    break
                            
                            if match:
                                elements_to_delete = []
                                for j in range(1, len(orig_paras)):
                                    elements_to_delete.append(all_paras[idx+j]._element)
                                
                                for p_idx, item in enumerate(new_paras):
                                    item = item.strip()
                                    is_heading = False
                                    
                                    if item.startswith("<HEADING>") and item.endswith("</HEADING>"):
                                        item = item[9:-10].strip()
                                        is_heading = True
                                        
                                    if p_idx == 0:
                                        all_paras[idx].text = item
                                        current_p = all_paras[idx]
                                    else:
                                        list_type = None
                                        if item.startswith("-"):
                                            item = item.lstrip("- ").strip()
                                            list_type = 'bullet'
                                        elif re.match(r'^\d+[\)\.]', item):
                                            item = re.sub(r'^\d+[\)\.]\s*', '', item).strip()
                                            list_type = 'number'

                                        if list_type:
                                            target_styles = ['List Bullet' if list_type == 'bullet' else 'List Number']
                                            new_p = None
                                            for s_name in target_styles:
                                                try:
                                                    new_p = doc.add_paragraph(item, style=s_name)
                                                    break
                                                except KeyError:
                                                    pass
                                            if not new_p:
                                                for s in doc.styles:
                                                    if s.type == 1:
                                                        name = s.name.lower()
                                                        if list_type == 'bullet' and ('bullet' in name or 'маркиров' in name):
                                                            new_p = doc.add_paragraph(item, style=s.name)
                                                            break
                                                        if list_type == 'number' and ('number' in name or 'нумеров' in name):
                                                            new_p = doc.add_paragraph(item, style=s.name)
                                                            break
                                            if not new_p:
                                                try:
                                                    new_p = doc.add_paragraph(item, style='List Paragraph')
                                                except KeyError:
                                                    new_p = doc.add_paragraph(item)
                                                pPr = new_p._p.get_or_add_pPr()
                                                numPr = OxmlElement('w:numPr')
                                                ilvl = OxmlElement('w:ilvl')
                                                ilvl.set(qn('w:val'), '0')
                                                numId = OxmlElement('w:numId')
                                                numId.set(qn('w:val'), '1' if list_type == 'bullet' else '2')
                                                numPr.append(ilvl)
                                                numPr.append(numId)
                                                pPr.append(numPr)

                                            pPr_f = new_p._p.get_or_add_pPr()
                                            for old_ind in pPr_f.findall(qn('w:ind')):
                                                pPr_f.remove(old_ind)
                                            ind = OxmlElement('w:ind')
                                            ind.set(qn('w:left'), '709')
                                            ind.set(qn('w:firstLine'), '0')
                                            pPr_f.append(ind)
                                        else:
                                            new_p = doc.add_paragraph(item)
                                            
                                        current_p._element.addnext(new_p._element)
                                        current_p = new_p

                                    if is_heading:
                                        from docx.enum.text import WD_ALIGN_PARAGRAPH
                                        current_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                        for r in current_p.runs:
                                            r.font.bold = True
                                
                                for elem in elements_to_delete:
                                    parent = elem.getparent()
                                    if parent is not None:
                                        parent.remove(elem)
                                    
                                break
                except Exception as e:
                    print(f"ProxyAPI error: {e}")
        
        apply_gost_styles(doc, target_pages)
        
        output_buffer = io.BytesIO()
        doc.save(output_buffer)
        output_buffer.seek(0)
        
        caption_text = "Ваш документ отформатирован! ✅"
        if target_pages != "ALL":
            caption_text += "\nШрифты и абзацы изменены только на выбранных страницах."
            
        bot.send_document(
            chat_id=message.chat.id,
            document=(f"ГОСТ_{user_data[message.chat.id]['file_name']}", output_buffer),
            caption=caption_text
        )
        
        bot.delete_message(chat_id=message.chat.id, message_id=status_message.message_id)
        
    except Exception as e:
        print(f"Error: {e}")
        traceback.print_exc()
        try:
            bot.reply_to(message, "❌ Произошла ошибка при обработке файла.")
        except:
            pass
        
    if message.chat.id in user_data:
        del user_data[message.chat.id]

if __name__ == "__main__":
    try:
        print("Бот запущен. Нажмите Ctrl+C для остановки.", flush=True)
        bot.infinity_polling(timeout=10, long_polling_timeout=5)
    except Exception as e:
        print("\n" + "="*50, flush=True)
        print("КРИТИЧЕСКАЯ ОШИБКА ПРИ РАБОТЕ БОТА:", flush=True)
        traceback.print_exc()
        print("="*50 + "\n", flush=True)
        while True:
            time.sleep(60)
