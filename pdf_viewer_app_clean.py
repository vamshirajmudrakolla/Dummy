import inspect
import time
from ctypes import windll, wintypes
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import os
import re
from collections import defaultdict
import difflib
from idlelib.tooltip import Hovertip
import threading
import sys
import pathlib
from tkinterdnd2 import DND_FILES, TkinterDnD
import fitz
import klembord
import subprocess
import copy
import tempfile
import traceback
try:
    import win32com.client
    import pythoncom
    on_windows = 1
except:
    on_windows = 0
try:
    import pyautogui
    PYAUTOGUI_AVAILABLE = True
except ImportError:
    PYAUTOGUI_AVAILABLE = False
TEMP_PDF_DIR = os.path.join(os.path.dirname(__file__), 'temp_pdfs')
os.makedirs(TEMP_PDF_DIR, exist_ok=True)
try:
    windll.user32.SetThreadDpiAwarenessContext(wintypes.HANDLE(-2))
except AttributeError:
    pass

def convert_clipboard_to_pdf(output_filename='clipboard_content.pdf'):
    try:
        klembord.init()
    except RuntimeError:
        return None
    html_content = None
    plain_text_content = None
    try:
        plain_text_content, html_content = klembord.get_with_rich_text()
    except Exception as e:
        plain_text_content = klembord.get_text()
    content_to_use = ''
    if html_content:
        content_to_use = html_content
        if content_to_use.lower().find('<html') != -1:
            content_to_use = content_to_use[content_to_use.lower().find('<html'):]
        elif content_to_use.lower().find('<head') != -1:
            content_to_use = content_to_use[content_to_use.lower().find('<head'):]

        def replace_style_content(match):
            style_content = match.group(1)
            new_style_content = re.sub('(?i)background:', 'background-color:', style_content)
            return f'style="{new_style_content}"'
        content_to_use = re.sub('style=["\\\'](.*?)["\\\']', replace_style_content, content_to_use, flags=re.DOTALL | re.IGNORECASE)
    elif plain_text_content:
        content_to_use = f"\n        <html>\n        <head>\n            <style>\n                body {{\n                    font-family: monospace; /* Often preferred for plain text */\n                    white-space: pre-wrap; /* Preserves whitespace and wraps long lines */\n                    word-wrap: break-word; /* Breaks long words if they don't fit */\n                }}\n            </style>\n        </head>\n        <body>\n            <div>{plain_text_content}</div>\n        </body>\n        </html>\n        "
    else:
        return None
    try:
        pathlib.Path(output_filename).parent.mkdir(parents=True, exist_ok=True)
        story = fitz.Story(html=content_to_use)
        writer = fitz.DocumentWriter(output_filename)
        mediabox = fitz.paper_rect('a4')
        where = mediabox + (36, 36, -36, -36)
        more = True
        page_number = 0
        while more:
            page_number += 1
            dev = writer.begin_page(mediabox)
            more, filled = story.place(where)
            story.draw(dev)
            writer.end_page()
        writer.close()
        return output_filename
    except Exception as e:
        return None

def convert_word_to_pdf_no_markup(input_file_path, output_pdf_path=None):
    if not on_windows:
        return
    input_file_path = input_file_path.replace('/', '\\')
    if output_pdf_path:
        output_pdf_path = output_pdf_path.replace('/', '\\')
    if not os.path.exists(input_file_path):
        return None
    if output_pdf_path is None:
        base_name = os.path.splitext(os.path.basename(input_file_path))[0]
        output_pdf_path = os.path.join(TEMP_PDF_DIR, f'{base_name}_temp_{os.urandom(4).hex()}.pdf')
    os.makedirs(TEMP_PDF_DIR, exist_ok=True)
    wdFormatPDF = 17
    wdRevisionsViewFinal = 0
    word_app = None
    doc = None
    try:
        pythoncom.CoInitialize()
        word_app = win32com.client.DispatchEx('Word.Application')
        word_app.Visible = False
        word_app.DisplayAlerts = False
        doc = word_app.Documents.Open(str(input_file_path))
        if hasattr(word_app.Options, 'WarnBeforeSavingPrintingSendingMarkup'):
            word_app.Options.WarnBeforeSavingPrintingSendingMarkup = False
        if doc.ActiveWindow:
            doc.ActiveWindow.View.RevisionsView = wdRevisionsViewFinal
        if hasattr(doc, 'ShowRevisions'):
            doc.ShowRevisions = False
        if hasattr(word_app.Options, 'PrintRevisions'):
            word_app.Options.PrintRevisions = False
        if hasattr(word_app.Options, 'PrintComments'):
            word_app.Options.PrintComments = False
        if hasattr(word_app.Options, 'PrintHiddenText'):
            word_app.Options.PrintHiddenText = False
        if hasattr(word_app.Options, 'PrintDrawingObjects'):
            word_app.Options.PrintDrawingObjects = True
        doc.SaveAs(str(output_pdf_path), FileFormat=wdFormatPDF)
        doc.Close(SaveChanges=False)
        return output_pdf_path
    except Exception as e:
        try:
            excepinfo = pythoncom.GetErrorInfo()
            if excepinfo:
                pass
        except Exception:
            pass
        return None
    finally:
        if word_app:
            try:
                word_app.Quit(SaveChanges=0)
            except Exception as e:
                pass
        pythoncom.CoUninitialize()

def extract_words_with_styles(pdf_document):
    all_words_data = []
    LINE_TOLERANCE_Y = 3
    for page_num, page in enumerate(pdf_document):
        if app.ignore_ligatures.get():
            words_data = page.get_text('words', flags=0)
        else:
            words_data = page.get_text('words')
        top_left_in_block = dict()
        grouped_lines = []
        for word_info in words_data:
            x0, y0, x1, y1, word_text, block_no, _, _ = word_info[:8]
            word_center_y = (y0 + y1) / 2
            added_to_existing_line = False
            if block_no not in top_left_in_block:
                top_left_in_block[block_no] = (x0, y0)
            elif y0 < top_left_in_block[block_no][1] or (y0 == top_left_in_block[block_no][1] and x0 < top_left_in_block[block_no][0]):
                top_left_in_block[block_no] = (x0, y0)
            for line_group in grouped_lines:
                if abs(line_group['y_center'] - word_center_y) < LINE_TOLERANCE_Y and line_group['block_no'] == block_no:
                    line_group['words'].append(word_info)
                    line_group['y_center'] = sum(((w[1] + w[3]) / 2 for w in line_group['words'])) / len(line_group['words'])
                    added_to_existing_line = True
                    break
            if not added_to_existing_line:
                grouped_lines.append({'y_center': word_center_y, 'words': [word_info], 'block_no': block_no})
        grouped_lines.sort(key=lambda lg: (top_left_in_block[lg['block_no']][1], top_left_in_block[lg['block_no']][0], lg['y_center']))
        for line_group in grouped_lines:
            line_group['words'].sort(key=lambda w: w[0])
            for word_info in line_group['words']:
                x0, y0, x1, y1, word_text, _, _, _ = word_info[:8]
                current_font_family = ''
                current_font_size = 12
                current_font_color = '#000000'
                current_font_weight = 'normal'
                current_font_style = 'normal'
                all_words_data.append({'text': word_text, 'x0': x0, 'y0': y0, 'x1': x1, 'y1': y1, 'page_num': page_num, 'font_family': current_font_family, 'font_size': current_font_size, 'font_color': current_font_color, 'font_weight': current_font_weight, 'font_style': current_font_style, 'unique_id': None, 'highlight_color': None})
    return all_words_data

def helper_case_quotes(words_data1, words_data2, case_insensitive, ignore_quotes):
    a_compare = [word_info['text'] for word_info in words_data1]
    b_compare = [word_info['text'] for word_info in words_data2]
    if case_insensitive:
        a_compare = [word.lower() for word in a_compare]
        b_compare = [word.lower() for word in b_compare]
    if ignore_quotes:
        a_compare = [word.replace('â€˜', "'").replace('â€™', "'").replace('Ê¼', "'").replace('â€œ', '"').replace('â€\x9d', '"') for word in a_compare]
        b_compare = [word.replace('â€˜', "'").replace('â€™', "'").replace('Ê¼', "'").replace('â€œ', '"').replace('â€\x9d', '"') for word in b_compare]
    return (a_compare, b_compare)

def align_words_with_difflib(words_data1, words_data2, case_insensitive, ignore_quotes):
    import time
    "\n    Aligns two sequences of words using difflib.SequenceMatcher\n    and assigns common IDs or marks as unique.\n    Modifies words_data1 and words_data2 in place by setting 'unique_id'\n    and 'highlight_color'.\n    Args:\n        words_data1 (list): List of dictionaries for words in document 1.\n        words_data2 (list): List of dictionaries for words in document 2.\n        case_insensitive (bool): If True, comparisons ignore case.\n        ignore_quotes (bool): If True, various quote types are normalized to standard quotes.\n    "
    a_compare, b_compare = helper_case_quotes(words_data1, words_data2, case_insensitive, ignore_quotes)
    s = difflib.SequenceMatcher(None, a_compare, b_compare)
    common_word_id_counter = 0
    idx1_current = 0
    idx2_current = 0
    for tag, i1, i2, j1, j2 in s.get_opcodes():
        if tag == 'equal':
            for k in range(i2 - i1):
                common_id = f'common-word-{common_word_id_counter}'
                words_data1[idx1_current + k]['unique_id'] = common_id
                words_data2[idx2_current + k]['unique_id'] = common_id
                words_data1[idx1_current + k]['highlight_color'] = None
                words_data2[idx2_current + k]['highlight_color'] = None
                common_word_id_counter += 1
            idx1_current += i2 - i1
            idx2_current += j2 - j1
        elif tag == 'delete':
            for k in range(i2 - i1):
                words_data1[idx1_current + k]['unique_id'] = None
                words_data1[idx1_current + k]['highlight_color'] = 'red'
            idx1_current += i2 - i1
        elif tag == 'insert':
            for k in range(j2 - j1):
                words_data2[idx2_current + k]['unique_id'] = None
                words_data2[idx2_current + k]['highlight_color'] = 'green'
            idx2_current += j2 - j1
        elif tag == 'replace':
            for k in range(i2 - i1):
                words_data1[idx1_current + k]['unique_id'] = None
                words_data1[idx1_current + k]['highlight_color'] = 'red'
            for k in range(j2 - j1):
                words_data2[idx2_current + k]['unique_id'] = None
                words_data2[idx2_current + k]['highlight_color'] = 'green'
            idx1_current += i2 - i1
            idx2_current += j2 - j1
    return (words_data1, words_data2)

def apply_annotations_to_pdf_pages(pdf_document, words_data):
    if not pdf_document or pdf_document.is_closed:
        return
    words_by_page = defaultdict(list)
    for word in words_data:
        if word['highlight_color']:
            words_by_page[word['page_num']].append(word)
    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)
        annotations_to_delete = [annot for annot in page.annots() if annot.type[0] == fitz.PDF_ANNOT_HIGHLIGHT and annot.info.get('title') == 'PDFComparer']
        for annot in annotations_to_delete:
            try:
                page.delete_annot(annot)
            except Exception as e:
                pass
        page_words = words_by_page[page_num]
        if not page_words:
            continue
        highlights_by_color = defaultdict(list)
        for word in page_words:
            rect = fitz.Rect(word['x0'], word['y0'], word['x1'], word['y1'])
            highlights_by_color[word['highlight_color']].append(rect)
        total_annotations_added = 0
        for color, rects_to_merge in highlights_by_color.items():
            if not rects_to_merge:
                continue
            merged_rects = []
            if rects_to_merge:
                current_merged_rect = rects_to_merge[0]
                for i in range(1, len(rects_to_merge)):
                    next_rect = rects_to_merge[i]
                    y_tolerance = 10
                    x_tolerance = 10
                    if abs(current_merged_rect.y0 - next_rect.y0) < y_tolerance and abs(current_merged_rect.y1 - next_rect.y1) < y_tolerance and (next_rect.x0 <= current_merged_rect.x1 + x_tolerance):
                        current_merged_rect = current_merged_rect | next_rect
                    else:
                        merged_rects.append(current_merged_rect)
                        current_merged_rect = next_rect
                merged_rects.append(current_merged_rect)
            highlight_color_rgb_float = (0.0, 0.0, 0.0)
            if color == 'red':
                highlight_color_rgb_float = (1.0, 0.0, 0.0)
            elif color == 'green':
                highlight_color_rgb_float = (0.0, 1.0, 0.0)
            elif color == 'blue':
                highlight_color_rgb_float = (0.0, 0.5, 1.0)
            for merged_rect in merged_rects:
                try:
                    annot = page.add_highlight_annot(merged_rect)
                    annot.set_colors(stroke=highlight_color_rgb_float)
                    annot.set_opacity(0.3)
                    annot.set_info(title='PDFComparer')
                    annot.update()
                    total_annotations_added += 1
                except Exception as e:
                    pass

class GitSequenceMatcher:

    def __init__(self, a, b, temp_dir=None):
        self.a = a
        self.b = b
        self.temp_file_a = None
        self.temp_file_b = None
        self.temp_dir = temp_dir

    def _create_temp_files(self):
        with tempfile.NamedTemporaryFile(mode='w+', delete=False, encoding='utf-8', dir=self.temp_dir) as f_a:
            self.temp_file_a = f_a.name
            for item in self.a:
                f_a.write(repr(item) + '\n')
        with tempfile.NamedTemporaryFile(mode='w+', delete=False, encoding='utf-8', dir=self.temp_dir) as f_b:
            self.temp_file_b = f_b.name
            for item in self.b:
                f_b.write(repr(item) + '\n')

    def _cleanup_temp_files(self):
        if self.temp_file_a and os.path.exists(self.temp_file_a):
            os.remove(self.temp_file_a)
        if self.temp_file_b and os.path.exists(self.temp_file_b):
            os.remove(self.temp_file_b)

    def get_opcodes(self):
        self._create_temp_files()
        process = None
        start_time = time.time()
        try:
            command = ['git', '--no-pager', 'diff', '--diff-algorithm=histogram', '--color=always', '--color-moved', '--unified=99999999', self.temp_file_a, self.temp_file_b]
            process = subprocess.run(command, capture_output=True, text=True, encoding='utf-8', errors='replace')
            diff_output = process.stdout
            if process.returncode == 0 and (not diff_output.strip()):
                with open(self.temp_file_a, 'r', encoding='utf-8', errors='replace') as f:
                    num_lines = sum((1 for _ in f))
                return [('equal', 0, num_lines, 0, num_lines, False)]
            COLOR_RED_FG = '\\x1b\\[31m'
            COLOR_GREEN_FG = '\\x1b\\[32m'
            COLOR_BOLD_MAGENTA_FG = '\\x1b\\[1;35m'
            COLOR_BLUE_FG = '\\x1b\\[1;34m'
            COLOR_BOLD_CYAN_FG = '\\x1b\\[1;36m'
            COLOR_BOLD_YELLOW_FG = '\\x1b\\[1;33m'
            COLOR_RED_BG = '\\x1b\\[41m'
            current_a_idx = 0
            current_b_idx = 0
            lines = diff_output.splitlines()
            in_hunk = False
            granular_changes = []
            for line_num, line in enumerate(lines):
                line_without_ansi = re.sub('\\x1b\\[[0-9;]*m', '', line)
                if line.startswith('\x1b[1mdiff --git'):
                    in_hunk = True
                    continue
                if not in_hunk:
                    continue
                if line_without_ansi.strip().startswith('index ') or line_without_ansi.strip().startswith('--- a/') or line_without_ansi.strip().startswith('+++ b/'):
                    continue
                if line_without_ansi.strip().startswith('@@'):
                    match = re.match('@@ -(\\d+)(?:,(\\d+))? \\+(\\d+)(?:,(\\d+))? @@', line_without_ansi.strip())
                    if match:
                        current_a_idx = int(match.group(1)) - 1
                        current_b_idx = int(match.group(3)) - 1
                    continue
                tag = None
                content_to_match = ''
                if re.search(f'^{COLOR_BOLD_MAGENTA_FG}-', line) or re.search(f'^{COLOR_BLUE_FG}-', line):
                    tag = 'moved_delete'
                    content_to_match = line_without_ansi[1:].strip()
                elif re.search(f'^{COLOR_BOLD_CYAN_FG}\\+', line) or re.search(f'^{COLOR_BOLD_YELLOW_FG}\\+', line) or line.strip().endswith(f'{COLOR_RED_BG}'):
                    tag = 'moved_insert'
                    content_to_match = line_without_ansi[1:].strip()
                elif re.search(f'^{COLOR_RED_FG}-', line):
                    tag = 'delete'
                    content_to_match = line_without_ansi[1:].strip()
                elif re.search(f'^{COLOR_GREEN_FG}\\+', line):
                    tag = 'insert'
                    content_to_match = line_without_ansi[1:].strip()
                elif line_without_ansi.startswith(' '):
                    tag = 'equal'
                    content_to_match = line_without_ansi[1:].strip()
                elif not line_without_ansi.strip():
                    continue
                else:
                    continue
                if tag and content_to_match:
                    if tag == 'delete' or tag == 'moved_delete':
                        granular_changes.append((tag, content_to_match, current_a_idx, current_a_idx + 1, current_b_idx, current_b_idx))
                        current_a_idx += 1
                    elif tag == 'insert' or tag == 'moved_insert':
                        granular_changes.append((tag, content_to_match, current_a_idx, current_a_idx, current_b_idx, current_b_idx + 1))
                        current_b_idx += 1
                    elif tag == 'equal':
                        granular_changes.append((tag, content_to_match, current_a_idx, current_a_idx + 1, current_b_idx, current_b_idx + 1))
                        current_a_idx += 1
                        current_b_idx += 1
            moved_candidates = {}
            for idx, (g_tag, g_content, g_a1, g_a2, g_b1, g_b2) in enumerate(granular_changes):
                if g_tag in ['moved_delete', 'moved_insert']:
                    if g_content not in moved_candidates:
                        moved_candidates[g_content] = []
                    moved_candidates[g_content].append((g_a1, g_b1, g_tag, idx))
            is_moved_flags = {}
            for content, candidates in moved_candidates.items():
                deletes = [c for c in candidates if c[2] == 'moved_delete']
                inserts = [c for c in candidates if c[2] == 'moved_insert']
                matched_deletes = set()
                matched_inserts = set()
                for d_a1, d_b1, d_tag, d_idx in deletes:
                    if d_idx in matched_deletes:
                        continue
                    for i_a1, i_b1, i_tag, i_idx in inserts:
                        if i_idx in matched_inserts:
                            continue
                        is_moved_flags[d_idx] = True
                        is_moved_flags[i_idx] = True
                        matched_deletes.add(d_idx)
                        matched_inserts.add(i_idx)
                        break
            final_opcodes_pre_replace = []
            current_tag = None
            current_i1, current_i2, current_j1, current_j2 = (-1, -1, -1, -1)
            current_is_moved_flag = False
            for idx, (g_tag, g_content, g_a1, g_a2, g_b1, g_b2) in enumerate(granular_changes):
                actual_tag = g_tag
                if actual_tag in ['moved_delete', 'moved_insert']:
                    actual_tag = 'delete' if g_tag == 'moved_delete' else 'insert'
                is_moved_for_this_item = is_moved_flags.get(idx, False)
                if current_tag is None:
                    current_tag = actual_tag
                    current_i1, current_i2 = (g_a1, g_a2)
                    current_j1, current_j2 = (g_b1, g_b2)
                    current_is_moved_flag = is_moved_for_this_item
                    continue
                can_extend = False
                if actual_tag == current_tag and is_moved_for_this_item == current_is_moved_flag:
                    if actual_tag == 'equal':
                        if g_a1 == current_i2 and g_b1 == current_j2:
                            can_extend = True
                    elif actual_tag == 'delete':
                        if g_a1 == current_i2:
                            can_extend = True
                    elif actual_tag == 'insert':
                        if g_b1 == current_j2:
                            can_extend = True
                if can_extend:
                    current_i2 = g_a2
                    current_j2 = g_b2
                else:
                    final_opcodes_pre_replace.append((current_tag, current_i1, current_i2, current_j1, current_j2, current_is_moved_flag))
                    current_tag = actual_tag
                    current_i1, current_i2 = (g_a1, g_a2)
                    current_j1, current_j2 = (g_b1, g_b2)
                    current_is_moved_flag = is_moved_for_this_item
            if current_tag is not None:
                final_opcodes_pre_replace.append((current_tag, current_i1, current_i2, current_j1, current_j2, current_is_moved_flag))
            consolidated_opcodes = []
            i = 0
            while i < len(final_opcodes_pre_replace):
                current_op = final_opcodes_pre_replace[i]
                tag, i1, i2, j1, j2, is_moved = current_op
                if (tag == 'delete' and (not is_moved)) and i + 1 < len(final_opcodes_pre_replace):
                    next_op = final_opcodes_pre_replace[i + 1]
                    next_tag, next_i1, next_i2, next_j1, next_j2, next_is_moved = next_op
                    if (next_tag == 'insert' and (not next_is_moved)) and i2 == next_i1 and (j2 == next_j1):
                        consolidated_opcodes.append(('replace', i1, i2, j1, next_j2, False))
                        i += 2
                        continue
                consolidated_opcodes.append(current_op)
                i += 1
            opcodes = sorted(consolidated_opcodes, key=lambda x: (x[1], x[3]))
        except Exception as e:
            traceback.print_exc()
            if process:
                pass
            return []
        finally:
            self._cleanup_temp_files()
        return opcodes

def align_words_with_git_diff(words_data1, words_data2, case_insensitive, ignore_quotes):
    a_compare, b_compare = helper_case_quotes(words_data1, words_data2, case_insensitive, ignore_quotes)
    s = GitSequenceMatcher(a_compare, b_compare, temp_dir='.')
    common_word_id_counter = 0
    idx1_current = 0
    idx2_current = 0
    for tag, i1, i2, j1, j2, is_moved in s.get_opcodes():
        if tag == 'equal':
            for k in range(i2 - i1):
                common_id = f'common-word-{common_word_id_counter}'
                words_data1[idx1_current + k]['unique_id'] = common_id
                words_data2[idx2_current + k]['unique_id'] = common_id
                words_data1[idx1_current + k]['highlight_color'] = None
                words_data2[idx2_current + k]['highlight_color'] = None
                common_word_id_counter += 1
            idx1_current += i2 - i1
            idx2_current += j2 - j1
        elif tag == 'delete' and (not is_moved):
            for k in range(i2 - i1):
                words_data1[idx1_current + k]['unique_id'] = None
                words_data1[idx1_current + k]['highlight_color'] = 'red'
            idx1_current += i2 - i1
        elif tag == 'insert' and (not is_moved):
            for k in range(j2 - j1):
                words_data2[idx2_current + k]['unique_id'] = None
                words_data2[idx2_current + k]['highlight_color'] = 'green'
            idx2_current += j2 - j1
        elif tag == 'replace':
            for k in range(i2 - i1):
                words_data1[idx1_current + k]['unique_id'] = None
                words_data1[idx1_current + k]['highlight_color'] = 'red'
            for k in range(j2 - j1):
                words_data2[idx2_current + k]['unique_id'] = None
                words_data2[idx2_current + k]['highlight_color'] = 'green'
            idx1_current += i2 - i1
            idx2_current += j2 - j1
        elif tag == 'insert' and is_moved:
            for k in range(j2 - j1):
                words_data2[idx2_current + k]['unique_id'] = None
                words_data2[idx2_current + k]['highlight_color'] = 'blue'
            idx2_current += j2 - j1
        elif tag == 'delete' and is_moved:
            for k in range(i2 - i1):
                words_data1[idx1_current + k]['unique_id'] = None
                words_data1[idx1_current + k]['highlight_color'] = 'blue'
            idx1_current += i2 - i1
    return (words_data1, words_data2)

def is_git_diff_available():
    try:
        subprocess.run(['git', 'diff'], check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        return False
if is_git_diff_available():
    align_words = align_words_with_git_diff
else:
    align_words = align_words_with_difflib

class PDFViewerPane:
    PAGE_PADDING = 10
    BUFFER_PAGES = 3

    def __init__(self, master, parent_app, pane_id):
        self.sorted = None
        self.master = master
        self.parent_app = parent_app
        self.pane_id = pane_id
        self.pdf_document = None
        self.words_data = []
        self.zoom_level = 1.0
        self.rendered_page_cache = {}
        self.page_layout_info = {}
        self.total_document_height = 0
        self.max_document_width = 0
        self.last_mouse_x = 0
        self.last_mouse_y = 0
        self.file_name = None
        self.pan_start_x = 0
        self.pan_start_y = 0
        self.canvas_start_x_offset = 0
        self.canvas_start_y_offset = 0
        self.panning = False
        self.render_job_id = None
        self.resize_job_id = None
        self.ignore_scroll_events_counter = 0
        self.temp_pdf_path = None
        self.loading_message_id = None
        self.setup_ui()

    def setup_ui(self):
        self.canvas_frame = ttk.Frame(self.master, relief=tk.SUNKEN, borderwidth=1)
        self.canvas_frame.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        self.v_scrollbar = ttk.Scrollbar(self.canvas_frame, orient=tk.VERTICAL)
        self.v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.h_scrollbar = ttk.Scrollbar(self.canvas_frame, orient=tk.HORIZONTAL)
        self.h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas = tk.Canvas(self.canvas_frame, bg='gray', yscrollcommand=self.v_scrollbar.set, xscrollcommand=self.h_scrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.v_scrollbar.config(command=self.on_vertical_scroll)
        self.h_scrollbar.config(command=self.on_horizontal_scroll)
        self.canvas.bind('<Configure>', self.on_canvas_configure)
        self.canvas.bind('<MouseWheel>', self.on_mousewheel)
        self.canvas.bind('<Button-4>', self.on_mousewheel)
        self.canvas.bind('<Button-5>', self.on_mousewheel)
        self.canvas.bind('<ButtonPress-1>', self.start_pan)
        self.canvas.bind('<B1-Motion>', self.do_pan)
        self.canvas.bind('<ButtonRelease-1>', self.stop_pan)
        self.canvas.bind('<Up>', self.on_key_scroll)
        self.canvas.bind('<Down>', self.on_key_scroll)
        self.canvas.bind('<Left>', self.on_key_scroll)
        self.canvas.bind('<Right>', self.on_key_scroll)
        self.canvas.bind('<Prior>', self.on_key_scroll)
        self.canvas.bind('<Next>', self.on_key_scroll)
        self.canvas.bind('<Home>', self.on_key_scroll)
        self.canvas.bind('<End>', self.on_key_scroll)
        self.canvas.bind('<<UserCanvasScrolled>>', lambda event, pane=self: self.parent_app.on_pane_scrolled(event, pane))
        self.canvas.drop_target_register(DND_FILES)
        self.canvas.dnd_bind('<<Drop>>', self.on_drop)
        self.canvas.bind('<Button-3>', self.on_right_click)
        self.context_menu = tk.Menu(self.master, tearoff=0)
        self.canvas.bind('<Double-Button-1>', self._toggle_pan_mode)
        self.canvas.bind('<Motion>', self._on_pan_move)
        self._pan_mode_active = False
        self._cursor_start_pos = None
        self._after_id = None

    def _toggle_pan_mode(self, event):
        if self._pan_mode_active:
            self._deactivate_pan_mode()
        else:
            self._activate_pan_mode()

    def _activate_pan_mode(self):
        if not PYAUTOGUI_AVAILABLE:
            return
        self._pan_mode_active = True
        self.canvas.config(cursor='hand2')
        self._cursor_start_pos = pyautogui.position()
        canvas_x = self.canvas.winfo_pointerx() - self.canvas.winfo_rootx()
        canvas_y = self.canvas.winfo_pointery() - self.canvas.winfo_rooty()
        self.canvas.scan_mark(canvas_x, canvas_y)
        self._snap_back_timer()

    def _deactivate_pan_mode(self):
        self._pan_mode_active = False
        self.canvas.config(cursor='')
        if self._after_id:
            self.master.after_cancel(self._after_id)
            self._after_id = None

    def _on_pan_move(self, event):
        if self._pan_mode_active:
            self.canvas.scan_dragto(self._cursor_start_pos.x - self.canvas.winfo_rootx(), event.y, gain=3)
            self.schedule_render_visible_pages()
            if self.ignore_scroll_events_counter == 0:
                self.canvas.event_generate('<<UserCanvasScrolled>>')
            if self._after_id:
                self.master.after_cancel(self._after_id)
                self._after_id = None
                self._after_id = self.master.after(40, self._snap_back_timer)

    def _snap_back_timer(self):
        if not self._pan_mode_active:
            return
        pyautogui.moveTo(self._cursor_start_pos.x, self._cursor_start_pos.y)
        canvas_x = self._cursor_start_pos.x - self.canvas.winfo_rootx()
        canvas_y = self._cursor_start_pos.y - self.canvas.winfo_rooty()
        self.canvas.scan_mark(canvas_x, canvas_y)
        self._after_id = self.master.after(400, self._snap_back_timer)

    def on_right_click(self, event):
        self.context_menu.delete(0, tk.END)
        if self.pdf_document and (not self.pdf_document.is_closed):
            self.context_menu.add_command(label='Save PDF with Annotations...', command=self.save_pdf_with_annotations)
            self.context_menu.add_command(label='Toggle light/dark mode', command=self.toggle_light_dark_mode)
            self.context_menu.add_separator()
        self.context_menu.add_command(label='Paste from Clipboard', command=self.paste_from_clipboard_action)
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def toggle_light_dark_mode(self):
        if not self.pdf_document or self.pdf_document.is_closed:
            return
        current_mode = None
        for page in self.pdf_document:
            if current_mode:
                break
            for annot in page.annots():
                if annot.type[0] == fitz.PDF_ANNOT_HIGHLIGHT and annot.info.get('title') == 'PDFComparer':
                    current_mode = annot.blendmode
                    break
        if current_mode is None:
            return
        new_mode = 'Exclusion' if current_mode == 'Multiply' else 'Multiply'
        for page in self.pdf_document:
            for annot in page.annots():
                if annot.type[0] == fitz.PDF_ANNOT_HIGHLIGHT and annot.info.get('title') == 'PDFComparer':
                    annot.set_blendmode(new_mode)
                    if new_mode == 'Exclusion':
                        annot.set_opacity(1)
                    else:
                        annot.set_opacity(0.3)
                    annot.update()
        self._clear_all_rendered_pages()
        self.calculate_document_layout()
        self.render_visible_pages()
        self.canvas.focus_set()

    def paste_from_clipboard_action(self):
        temp_output_filename = os.path.join(TEMP_PDF_DIR, f'clipboard_temp_{os.urandom(8).hex()}.pdf')
        self.display_loading_message('Pasting from clipboard...')
        load_thread = threading.Thread(target=self._paste_from_clipboard_threaded, args=(temp_output_filename,))
        load_thread.daemon = True
        load_thread.start()

    def _paste_from_clipboard_threaded(self, temp_output_filename):
        converted_file_path = convert_clipboard_to_pdf(temp_output_filename)
        self.master.after(1, self._on_paste_from_clipboard_complete_gui_update, converted_file_path, temp_output_filename)

    def _on_paste_from_clipboard_complete_gui_update(self, converted_file_path, original_temp_filename):
        self.hide_loading_message()
        if converted_file_path:
            self.parent_app._initiate_load_process(converted_file_path, 0 if self.pane_id == 'left' else 1, 'Clipboard Content')
            self.temp_pdf_path = converted_file_path
        else:
            messagebox.showerror('Clipboard Error', 'Could not convert clipboard content to PDF. It might be empty or contain unsupported content.')
            self._clear_all_rendered_pages()

    def save_pdf_with_annotations(self):
        if not self.pdf_document or self.pdf_document.is_closed:
            messagebox.showinfo('Save PDF', 'No PDF document is currently open in this pane to save.')
            return
        initial_file = self.file_name if self.file_name else 'document'
        base_name, ext = os.path.splitext(initial_file)
        if ext.lower() != '.pdf':
            initial_file = base_name + '.pdf'
        if '_diff' not in base_name.lower():
            initial_file = f'{base_name}_diff{ext}'
        file_path = filedialog.asksaveasfilename(defaultextension='.pdf', filetypes=[('PDF files', '*.pdf'), ('All files', '*.*')], initialfile=initial_file)
        if file_path:
            try:
                self.pdf_document.save(file_path)
                messagebox.showinfo('Save PDF', f'PDF saved successfully to:\n{file_path}')
            except Exception as e:
                messagebox.showerror('Save PDF Error', f'Failed to save PDF: {e}')

    def on_drop(self, event):
        file_path = event.data
        if file_path.startswith('{') and file_path.endswith('}'):
            file_path = file_path[1:-1]
        self.parent_app.open_pdf_from_drop(file_path, self.pane_id)

    def display_loading_message(self, message='Loading...'):
        self.hide_loading_message()
        self.canvas.delete('all')
        canvas_center_x = self.canvas.winfo_width() / 2
        canvas_center_y = self.canvas.winfo_height() / 2
        self.loading_message_id = self.canvas.create_text(canvas_center_x, canvas_center_y, text=message, fill='black', font=('Helvetica', 24, 'bold'), tags='loading_message')
        self.canvas.config(scrollregion=(0, 0, 0, 0))

    def hide_loading_message(self):
        if self.loading_message_id:
            self.canvas.delete(self.loading_message_id)
            self.loading_message_id = None

    def load_pdf_internal(self, file_path):
        temp_pdf_path_used = None
        pdf_document_obj = None
        words_data_obj = []
        try:
            original_file_extension = os.path.splitext(file_path)[1].lower()
            if original_file_extension in ['.doc', '.docx', '.rtf', '.txt']:
                converted_pdf_path = convert_word_to_pdf_no_markup(file_path)
                if converted_pdf_path:
                    file_path = converted_pdf_path
                    temp_pdf_path_used = converted_pdf_path
                else:
                    return (None, [], None, 'Conversion Failed')
            pdf_document_obj = fitz.open(file_path)
            words_data_obj = extract_words_with_styles(pdf_document_obj)
            if file_path.find('clipboard_temp_') != -1:
                temp_pdf_path_used = file_path
            return (pdf_document_obj, words_data_obj, temp_pdf_path_used, None)
        except Exception as e:
            raise
            if temp_pdf_path_used and os.path.exists(temp_pdf_path_used):
                try:
                    os.remove(temp_pdf_path_used)
                except Exception as cleanup_e:
                    pass
            if pdf_document_obj:
                pdf_document_obj.close()
            return (None, [], None, f'Could not open PDF: {e}')

    def get_current_view_coords(self):
        return (self.canvas.canvasx(0), self.canvas.canvasy(0))

    def get_current_view_height_in_content_coords(self):
        return self.canvas.winfo_height()

    def calculate_document_layout(self):
        self.page_layout_info.clear()
        y_offset = 0
        max_width = 0
        if not self.pdf_document or self.pdf_document.is_closed or self.pdf_document.page_count == 0:
            self.total_document_height = 0
            self.max_document_width = 0
            self.canvas.config(scrollregion=(0, 0, 0, 0))
            return
        for i in range(self.pdf_document.page_count):
            page = self.pdf_document.load_page(i)
            base_width = int(page.mediabox.width)
            base_height = int(page.mediabox.height)
            self.page_layout_info[i] = {'base_width': base_width, 'base_height': base_height, 'y_start_offset': y_offset}
            y_offset += base_height + self.PAGE_PADDING
            max_width = max(max_width, base_width)
        self.total_document_height = y_offset
        self.max_document_width = max_width
        self.canvas.config(scrollregion=(0, 0, self.max_document_width * self.zoom_level, self.total_document_height * self.zoom_level))

    def schedule_render_visible_pages(self, event=None):
        if self.render_job_id:
            self.master.after_cancel(self.render_job_id)
        if self.pdf_document and (not self.pdf_document.is_closed):
            self.render_job_id = self.master.after(50, self.render_visible_pages)

    def schedule_fit_to_width(self, event=None):
        if self.resize_job_id:
            self.master.after_cancel(self.resize_job_id)
        self.resize_job_id = self.master.after(150, self.fit_to_width)

    def _clear_all_rendered_pages(self):
        self.canvas.delete('all')
        self.rendered_page_cache.clear()
        self.hide_loading_message()

    def render_visible_pages(self):
        if not self.pdf_document or self.pdf_document.is_closed or self.pdf_document.page_count == 0:
            return
        if self.canvas.winfo_width() == 0 or self.canvas.winfo_height() == 0:
            return
        current_view_x_content_coord = self.canvas.canvasx(0)
        current_view_y_content_coord = self.canvas.canvasy(0)
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        visible_y_start = current_view_y_content_coord
        visible_y_end = current_view_y_content_coord + canvas_height
        pages_to_render_now = set()
        for page_num in range(self.pdf_document.page_count):
            page_info = self.page_layout_info.get(page_num)
            if not page_info:
                continue
            scaled_y_start = page_info['y_start_offset'] * self.zoom_level
            scaled_height = page_info['base_height'] * self.zoom_level
            scaled_height_with_padding = scaled_height + self.PAGE_PADDING * self.zoom_level
            buffer_height_px = self.BUFFER_PAGES * scaled_height_with_padding
            page_top_buffered = scaled_y_start - buffer_height_px
            page_bottom_buffered = scaled_y_start + scaled_height_with_padding + buffer_height_px
            if page_bottom_buffered >= visible_y_start and page_top_buffered <= visible_y_end:
                pages_to_render_now.add(page_num)
        pages_currently_cached = set(self.rendered_page_cache.keys())
        pages_to_remove = pages_currently_cached - pages_to_render_now
        for page_num in pages_to_remove:
            if page_num in self.rendered_page_cache:
                data = self.rendered_page_cache[page_num]
                self.canvas.delete(data['canvas_id'])
                del self.rendered_page_cache[page_num]
        for page_num in pages_to_render_now:
            if page_num not in self.rendered_page_cache:
                try:
                    if self.pdf_document.is_closed:
                        continue
                    page = self.pdf_document.load_page(page_num)
                    matrix = fitz.Matrix(self.zoom_level, self.zoom_level)
                    pix = page.get_pixmap(matrix=matrix)
                    img = Image.frombytes('RGB', [pix.width, pix.height], pix.samples)
                    tk_img = ImageTk.PhotoImage(img)
                    content_width_at_zoom = self.max_document_width * self.zoom_level
                    page_width_at_zoom = page.rect.width * self.zoom_level
                    page_x_offset_on_canvas = (content_width_at_zoom - page_width_at_zoom) / 2
                    y_pos_on_canvas = self.page_layout_info[page_num]['y_start_offset'] * self.zoom_level
                    canvas_id = self.canvas.create_image(page_x_offset_on_canvas, y_pos_on_canvas, anchor=tk.NW, image=tk_img)
                    self.rendered_page_cache[page_num] = {'image': tk_img, 'canvas_id': canvas_id}
                except Exception as e:
                    break

    def fit_to_width(self):
        if not self.pdf_document or self.pdf_document.is_closed or self.pdf_document.page_count == 0:
            return
        if self.canvas.winfo_width() == 0:
            self.master.after(100, self.fit_to_width)
            return
        page_rect = self.pdf_document.load_page(0).mediabox
        page_width = page_rect.width
        canvas_width = self.canvas.winfo_width() - self.v_scrollbar.winfo_width()
        if canvas_width <= 0:
            canvas_width = 1
        new_zoom = canvas_width / page_width
        self.set_zoom(new_zoom)

    def set_zoom(self, new_zoom_level, mouse_x_canvas_pixel=None, mouse_y_canvas_pixel=None, from_sync=False):
        old_zoom = self.zoom_level
        if abs(new_zoom_level - old_zoom) < 0.001:
            return
        if not self.pdf_document or self.pdf_document.is_closed:
            return
        if self.canvas.winfo_width() == 0 or self.canvas.winfo_height() == 0:
            self.master.after(100, lambda: self.set_zoom(new_zoom_level, mouse_x_canvas_pixel, mouse_y_canvas_pixel, from_sync))
            return
        if mouse_x_canvas_pixel is None:
            mouse_x_canvas_pixel = int(self.canvas.winfo_width() / 2)
        if mouse_y_canvas_pixel is None:
            mouse_y_canvas_pixel = int(self.canvas.winfo_height() / 2)
        mouse_x_content_coord_old_zoom = self.canvas.canvasx(mouse_x_canvas_pixel)
        mouse_y_content_coord_old_zoom = self.canvas.canvasy(mouse_y_canvas_pixel)
        mouse_x_doc_coord = mouse_x_content_coord_old_zoom / old_zoom if old_zoom != 0 else 0
        mouse_y_doc_coord = mouse_y_content_coord_old_zoom / old_zoom if old_zoom != 0 else 0
        self._clear_all_rendered_pages()
        self.zoom_level = new_zoom_level
        self.calculate_document_layout()
        new_mouse_x_content_coord = mouse_x_doc_coord * self.zoom_level
        new_mouse_y_content_coord = mouse_y_doc_coord * self.zoom_level
        new_x_scroll_pixels = new_mouse_x_content_coord - mouse_x_canvas_pixel
        new_y_scroll_pixels = new_mouse_y_content_coord - mouse_y_canvas_pixel
        self._apply_scroll(new_x_scroll_pixels, new_y_scroll_pixels)
        self.render_visible_pages()
        self.canvas.focus_set()
        if not from_sync and self.parent_app.sync_zoom_enabled.get():
            self.parent_app.sync_zoom(self, new_zoom_level, mouse_x_canvas_pixel, mouse_y_canvas_pixel)
        self.parent_app.update_zoom_label(self.pane_id, self.zoom_level)

    def set_zoom_from_scale_widget(self, val):
        self.set_zoom(float(val), from_sync=False)

    def on_vertical_scroll(self, *args):
        self.canvas.yview(*args)
        self.schedule_render_visible_pages()
        if self.ignore_scroll_events_counter == 0:
            self.canvas.event_generate('<<UserCanvasScrolled>>')

    def on_horizontal_scroll(self, *args):
        self.canvas.xview(*args)
        self.schedule_render_visible_pages()
        if self.ignore_scroll_events_counter == 0:
            self.canvas.event_generate('<<UserCanvasScrolled>>')

    def on_mousewheel(self, event):
        if not self.pdf_document or self.pdf_document.is_closed:
            return 'break'
        scroll_delta = 0
        if event.delta:
            scroll_delta = -int(event.delta / 120)
        elif event.num == 4:
            scroll_delta = -1
        elif event.num == 5:
            scroll_delta = 1
        if event.state & 4:
            old_zoom = self.zoom_level
            zoom_factor = 1.1 if scroll_delta < 0 else 1 / 1.1
            new_zoom_level = self.zoom_level * zoom_factor
            min_zoom = 0.25
            max_zoom = 4.0
            new_zoom_level = max(min_zoom, min(max_zoom, new_zoom_level))
            if abs(new_zoom_level - old_zoom) > 0.001:
                self.set_zoom(new_zoom_level, event.x, event.y, from_sync=False)
        elif event.state == 9 or event.state & 1:
            self.canvas.xview_scroll(scroll_delta, 'units')
            if self.ignore_scroll_events_counter == 0:
                self.canvas.event_generate('<<UserCanvasScrolled>>')
        else:
            self.canvas.yview_scroll(scroll_delta, 'units')
            if self.ignore_scroll_events_counter == 0:
                self.canvas.event_generate('<<UserCanvasScrolled>>')
        self.schedule_render_visible_pages()
        return 'break'

    def on_key_scroll(self, event):
        if not self.pdf_document or self.pdf_document.is_closed:
            return 'break'
        scroll_amount_units = 3
        scroll_amount_pages = 1
        if event.keysym == 'Up':
            self.canvas.yview_scroll(-scroll_amount_units, 'units')
        elif event.keysym == 'Down':
            self.canvas.yview_scroll(scroll_amount_units, 'units')
        elif event.keysym == 'Left':
            self.canvas.xview_scroll(-scroll_amount_units, 'units')
        elif event.keysym == 'Right':
            self.canvas.xview_scroll(scroll_amount_units, 'units')
        elif event.keysym == 'Prior':
            self.canvas.yview_scroll(-scroll_amount_pages, 'pages')
        elif event.keysym == 'Next':
            self.canvas.yview_scroll(scroll_amount_pages, 'pages')
        elif event.keysym == 'Home':
            self.canvas.yview_moveto(0.0)
        elif event.keysym == 'End':
            self.canvas.yview_moveto(1.0)
        else:
            return
        if self.ignore_scroll_events_counter == 0:
            self.canvas.event_generate('<<UserCanvasScrolled>>')
        self.schedule_render_visible_pages()
        return 'break'

    def _apply_scroll(self, x_scroll_pixels, y_scroll_pixels):
        self.ignore_scroll_events_counter += 1
        try:
            if not self.pdf_document or self.pdf_document.is_closed:
                return
            total_width_at_zoom = self.max_document_width * self.zoom_level
            total_height_at_zoom = self.total_document_height * self.zoom_level
            max_x_scroll = max(0, total_width_at_zoom - self.canvas.winfo_width())
            max_y_scroll = max(0, total_height_at_zoom - self.canvas.winfo_height())
            x_scroll_pixels = max(0, min(x_scroll_pixels, max_x_scroll))
            y_scroll_pixels = max(0, min(y_scroll_pixels, max_y_scroll))
            x_prop = x_scroll_pixels / total_width_at_zoom if total_width_at_zoom > 0 else 0
            y_prop = y_scroll_pixels / total_height_at_zoom if total_height_at_zoom > 0 else 0
            self.canvas.xview_moveto(x_prop)
            self.canvas.yview_moveto(y_prop)
            self.schedule_render_visible_pages()
        finally:
            self.ignore_scroll_events_counter -= 1

    def start_pan(self, event):
        if not self.pdf_document or self.pdf_document.is_closed:
            return
        self.panning = True
        self.pan_start_x = event.x
        self.pan_start_y = event.y
        self.canvas_start_x_offset = self.canvas.canvasx(0)
        self.canvas_start_y_offset = self.canvas.canvasy(0)
        self.canvas.config(cursor='fleur')
        self.canvas.focus_set()

    def do_pan(self, event):
        if self.panning and self.pdf_document and (not self.pdf_document.is_closed):
            dx = event.x - self.pan_start_x
            dy = event.y - self.pan_start_y
            target_x_scroll = self.canvas_start_x_offset - dx
            target_y_scroll = self.canvas_start_y_offset - dy
            self.ignore_scroll_events_counter += 1
            try:
                total_width_at_zoom = self.max_document_width * self.zoom_level
                total_height_at_zoom = self.total_document_height * self.zoom_level
                x_prop = target_x_scroll / total_width_at_zoom if total_width_at_zoom > 0 else 0
                y_prop = target_y_scroll / total_height_at_zoom if total_height_at_zoom > 0 else 0
                x_prop = max(0.0, min(x_prop, 1.0))
                y_prop = max(0.0, min(y_prop, 1.0))
                self.canvas.xview_moveto(x_prop)
                self.canvas.yview_moveto(y_prop)
            finally:
                self.ignore_scroll_events_counter -= 1
            self.schedule_render_visible_pages()
            if self.ignore_scroll_events_counter == 0:
                self.canvas.event_generate('<<UserCanvasScrolled>>')

    def stop_pan(self, event):
        if not self.pdf_document or self.pdf_document.is_closed:
            return
        self.panning = False
        self.canvas.config(cursor='')
        self.schedule_render_visible_pages()
        self.canvas.focus_set()

    def on_canvas_configure(self, event):
        if self.pdf_document and (not self.pdf_document.is_closed):
            current_x_prop = self.canvas.xview()[0]
            current_y_prop = self.canvas.yview()[0]
            self.calculate_document_layout()
            self.schedule_fit_to_width()
            total_width_at_zoom = self.max_document_width * self.zoom_level
            total_height_at_zoom = self.total_document_height * self.zoom_level
            self._apply_scroll(current_x_prop * total_width_at_zoom, current_y_prop * total_height_at_zoom)
            self.render_visible_pages()
        self.canvas.focus_set()

    def close_pdf(self):
        if self.pdf_document:
            try:
                if 0:
                    for page_num in range(self.pdf_document.page_count):
                        page = self.pdf_document.load_page(page_num)
                        annots_to_delete = [annot for annot in page.annots() if annot.type[0] == fitz.PDF_ANNOT_HIGHLIGHT and annot.info.get('title') == 'PDFComparer']
                        for annot in annots_to_delete:
                            try:
                                page.delete_annot(annot)
                            except Exception as e:
                                pass
                self.pdf_document.close()
            except Exception as e:
                pass
            self.pdf_document = None
        self.file_name = None
        self.rendered_page_cache.clear()
        self.words_data = []
        self.page_layout_info = {}
        self.canvas.delete('all')
        self.canvas.config(scrollregion=(0, 0, 0, 0))
        if self.temp_pdf_path and os.path.exists(self.temp_pdf_path):
            try:
                os.remove(self.temp_pdf_path)
            except Exception as e:
                pass
            self.temp_pdf_path = None
        self.hide_loading_message()

class PDFViewerApp:

    def __init__(self, master):
        self.master = master
        self.master.geometry('1200x800')
        self.pdf_documents = [None, None]
        self.words_data_list = [None, None]
        self.current_active_pane = None
        self.pane1 = None
        self.pane2 = None
        self.zoom_scale_1 = None
        self.zoom_scale_2 = None
        self.zoom_percent_label_1 = None
        self.zoom_percent_label_2 = None
        self.sync_scroll_enabled = tk.BooleanVar(value=True)
        self.sync_zoom_enabled = tk.BooleanVar(value=True)
        self.case_insensitive = tk.BooleanVar(value=True)
        self.ignore_quotes = tk.BooleanVar(value=True)
        self.ignore_ligatures = tk.BooleanVar(value=True)
        self.setup_ui()
        self.update_window_title()
        self.master.after_idle(self.update_ui_state)
        self.master.after_idle(lambda: self.pane1.canvas.focus_set())
        self._process_command_line_args()
        self.scroll_time = 0
        self.scroll_y = 0
        self.scroll_pane = None
        self.scroll_height = 0
        self.scroll_target_y = 0
        self.scroll_distance = 0

    def setup_ui(self):
        control_frame = ttk.Frame(self.master, padding='10')
        control_frame.pack(fill=tk.X, side=tk.TOP)
        self.open_button_1 = ttk.Button(control_frame, text='Open (L)', command=lambda: self.open_pdf(0))
        self.open_button_1.pack(side=tk.LEFT, padx=5)
        self.open_button_2 = ttk.Button(control_frame, text='Open (R)', command=lambda: self.open_pdf(1))
        self.open_button_2.pack(side=tk.LEFT, padx=5)
        ttk.Label(control_frame, text='Zoom (L):').pack(side=tk.LEFT, padx=(15, 0))
        self.zoom_scale_1 = ttk.Scale(control_frame, from_=0.33, to_=3.0, orient=tk.HORIZONTAL, length=100)
        self.zoom_scale_1.set(1.0)
        self.zoom_scale_1.pack(side=tk.LEFT, padx=5)
        self.zoom_percent_label_1 = ttk.Label(control_frame, text='100%')
        self.zoom_percent_label_1.pack(side=tk.LEFT)
        ttk.Label(control_frame, text='Zoom (R):').pack(side=tk.LEFT, padx=(15, 0))
        self.zoom_scale_2 = ttk.Scale(control_frame, from_=0.33, to_=3.0, orient=tk.HORIZONTAL, length=100)
        self.zoom_scale_2.set(1.0)
        self.zoom_scale_2.pack(side=tk.LEFT, padx=5)
        self.zoom_percent_label_2 = ttk.Label(control_frame, text='100%')
        self.zoom_percent_label_2.pack(side=tk.LEFT)
        self.prev_change_button = ttk.Button(control_frame, text='Prev.', command=self.go_to_prev_change, underline=0)
        self.prev_change_button.pack(side=tk.LEFT, padx=(20, 5))
        self.next_change_button = ttk.Button(control_frame, text='Next', command=self.go_to_next_change, underline=0)
        self.next_change_button.pack(side=tk.LEFT, padx=5)
        self.sync_scroll_checkbox = ttk.Checkbutton(control_frame, text='Sync Scroll', variable=self.sync_scroll_enabled, onvalue=True, offvalue=False)
        self.sync_scroll_checkbox.pack(side=tk.LEFT, padx=(20, 5))
        self.sync_zoom_checkbox = ttk.Checkbutton(control_frame, text='Sync Zoom', variable=self.sync_zoom_enabled, onvalue=True, offvalue=False)
        self.sync_zoom_checkbox.pack(side=tk.LEFT, padx=5)
        self.case_insensitive_checkbox = ttk.Checkbutton(control_frame, text='Case Insensitive', variable=self.case_insensitive, onvalue=True, offvalue=False)
        self.case_insensitive_checkbox.pack(side=tk.LEFT, padx=5)
        self.tip_case_insensitive = Hovertip(self.case_insensitive_checkbox, 'Works only BEFORE loading files.\nLoad again one file if you need to change this setting.')
        self.ignore_quotes_checkbox = ttk.Checkbutton(control_frame, text='Ignore quotes type', variable=self.ignore_quotes, onvalue=True, offvalue=False)
        self.ignore_quotes_checkbox.pack(side=tk.LEFT, padx=5)
        self.tip_ignore_quotes = Hovertip(self.ignore_quotes_checkbox, 'Works only BEFORE loading files.\nLoad again one file if you need to change this setting.')
        self.ignore_ligatures_checkbox = ttk.Checkbutton(control_frame, text="Ignore 'f' ligatures", variable=self.ignore_ligatures, onvalue=True, offvalue=False)
        self.ignore_ligatures_checkbox.pack(side=tk.LEFT, padx=5)
        self.tip_ignore_ligatures = Hovertip(self.ignore_ligatures_checkbox, 'Works only BEFORE loading files.\nLoad again one file if you need to change this setting.')
        self.panes_container = ttk.Frame(self.master)
        self.panes_container.pack(fill=tk.BOTH, expand=True)
        self.pane1 = PDFViewerPane(self.panes_container, self, 'left')
        self.pane1.canvas_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.pane1.canvas.bind('<FocusIn>', lambda e: self.set_active_pane(self.pane1))
        self.pane2 = PDFViewerPane(self.panes_container, self, 'right')
        self.pane2.canvas_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.pane2.canvas.bind('<FocusIn>', lambda e: self.set_active_pane(self.pane2))
        self.zoom_scale_1.config(command=self.pane1.set_zoom_from_scale_widget)
        self.zoom_scale_2.config(command=self.pane2.set_zoom_from_scale_widget)
        self.master.bind('p', lambda event: self.go_to_prev_change())
        self.master.bind('P', lambda event: self.go_to_prev_change())
        self.master.bind('n', lambda event: self.go_to_next_change())
        self.master.bind('N', lambda event: self.go_to_next_change())

    def _process_command_line_args(self):
        if len(sys.argv) > 1:
            file_path1 = sys.argv[1]
            self.master.after_idle(lambda: self._initiate_load_process(file_path1, 0, os.path.basename(file_path1)))
        if len(sys.argv) > 2:
            file_path2 = sys.argv[2]
            self.master.after_idle(lambda: self._initiate_load_process(file_path2, 1, os.path.basename(file_path2)))

    def update_window_title(self):
        name1 = self.pane1.file_name if self.pane1.file_name else 'Panel 1'
        name2 = self.pane2.file_name if self.pane2.file_name else 'Panel 2'
        self.master.title(f'PDF Diff Viewer - {name1} vs {name2}')

    def set_active_pane(self, pane):
        self.current_active_pane = pane

    def update_zoom_label(self, pane_id, zoom_level):
        if pane_id == 'left' and self.zoom_scale_1 and self.zoom_percent_label_1:
            self.zoom_percent_label_1.config(text=f'{int(zoom_level * 100)}%')
            if abs(self.zoom_scale_1.get() - zoom_level) > 0.001:
                self.zoom_scale_1.config(command=lambda *args: None)
                self.zoom_scale_1.set(zoom_level)
                self.zoom_scale_1.config(command=self.pane1.set_zoom_from_scale_widget)
        elif pane_id == 'right' and self.zoom_scale_2 and self.zoom_percent_label_2:
            self.zoom_percent_label_2.config(text=f'{int(zoom_level * 100)}%')
            if abs(self.zoom_scale_2.get() - zoom_level) > 0.001:
                self.zoom_scale_2.config(command=lambda *args: None)
                self.zoom_scale_2.set(zoom_level)
                self.zoom_scale_2.config(command=self.pane2.set_zoom_from_scale_widget)

    def update_ui_state(self):
        doc1_loaded = self.pdf_documents[0] and (not self.pdf_documents[0].is_closed) if self.pdf_documents[0] else False
        doc2_loaded = self.pdf_documents[1] and (not self.pdf_documents[1].is_closed) if self.pdf_documents[1] else False
        self.zoom_scale_1.config(state=tk.NORMAL if doc1_loaded else tk.DISABLED)
        self.zoom_scale_2.config(state=tk.NORMAL if doc2_loaded else tk.DISABLED)
        self.prev_change_button.config(state=tk.NORMAL if doc1_loaded and doc2_loaded else tk.DISABLED)
        self.next_change_button.config(state=tk.NORMAL if doc1_loaded and doc2_loaded else tk.DISABLED)
        self.update_zoom_label('left', self.pane1.zoom_level if doc1_loaded else 1.0)
        self.update_zoom_label('right', self.pane2.zoom_level if doc2_loaded else 1.0)

    def open_pdf(self, pane_index):
        file_types = [('All supported files', '*.pdf .docx *.doc *.rtf *.txt'), ('PDF files', '*.pdf'), ('Word Documents', '*.docx *.doc'), ('Rich Text Format', '*.rtf'), ('Text files', '*.txt'), ('All files', '*.*')]
        file_path = filedialog.askopenfilename(filetypes=file_types)
        if file_path:
            self._initiate_load_process(file_path, pane_index, os.path.basename(file_path))

    def open_pdf_from_drop(self, file_path, pane_id):
        pane_index = 0 if pane_id == 'left' else 1
        self._initiate_load_process(file_path, pane_index, os.path.basename(file_path))

    def _initiate_load_process(self, file_path, pane_index, display_file_name):
        pane = self.pane1 if pane_index == 0 else self.pane2
        pane.display_loading_message(f"Loading '{display_file_name}'...")
        self.pdf_documents[pane_index] = None
        self.words_data_list[pane_index] = None
        pane.words_data = []
        pane.close_pdf()
        pane._clear_all_rendered_pages()
        load_thread = threading.Thread(target=self._load_and_process_pdf_threaded, args=(file_path, pane_index, display_file_name))
        load_thread.daemon = True
        load_thread.start()

    def _load_and_process_pdf_threaded(self, file_path, pane_index, display_file_name):
        pane = self.pane1 if pane_index == 0 else self.pane2
        pdf_doc, words_data, temp_path, error_message = pane.load_pdf_internal(file_path)
        self.master.after(1, self._on_pdf_load_complete_gui_update, pane_index, pdf_doc, words_data, temp_path, error_message, display_file_name)

    def _on_pdf_load_complete_gui_update(self, pane_index, pdf_doc, words_data, temp_path, error_message, display_file_name):
        pane = self.pane1 if pane_index == 0 else self.pane2
        pane.hide_loading_message()
        if error_message:
            messagebox.showerror('Error', f'Failed to open/process file in {pane.pane_id} pane: {error_message}')
            self.pdf_documents[pane_index] = None
            self.words_data_list[pane_index] = None
            pane.temp_pdf_path = None
            pane.canvas.config(scrollregion=(0, 0, 0, 0))
            pane.canvas.delete('all')
            self.update_ui_state()
            self.update_window_title()
            return
        pane.pdf_document = pdf_doc
        pane.words_data = words_data
        pane.temp_pdf_path = temp_path
        pane.file_name = display_file_name
        self.pdf_documents[pane_index] = pdf_doc
        self.words_data_list[pane_index] = [dict(w) for w in words_data]
        pane.calculate_document_layout()
        pane.canvas.yview_moveto(0)
        pane.canvas.xview_moveto(0)
        pane.fit_to_width()
        self.update_ui_state()
        self.update_window_title()
        self.perform_comparison_if_ready()
        pane.canvas.focus_set()

    def perform_comparison_if_ready(self):
        doc1_ready = self.pdf_documents[0] and (not self.pdf_documents[0].is_closed) if self.pdf_documents[0] else False
        doc2_ready = self.pdf_documents[1] and (not self.pdf_documents[1].is_closed) if self.pdf_documents[1] else False
        if doc1_ready and doc2_ready:
            words1_copy = [dict(w) for w in self.words_data_list[0]] if self.words_data_list[0] else []
            words2_copy = [dict(w) for w in self.words_data_list[1]] if self.words_data_list[1] else []
            self.words_data_list[0], self.words_data_list[1] = align_words(words1_copy, words2_copy, self.case_insensitive.get(), self.ignore_quotes.get())
            self.pane1.words_data = self.words_data_list[0]
            self.pane2.words_data = self.words_data_list[1]
            apply_annotations_to_pdf_pages(self.pdf_documents[0], self.pane1.words_data)
            apply_annotations_to_pdf_pages(self.pdf_documents[1], self.pane2.words_data)
            self.pane1._clear_all_rendered_pages()
            self.pane2._clear_all_rendered_pages()
            self.pane1.render_visible_pages()
            self.pane2.render_visible_pages()
            if self.current_active_pane:
                self.sync_scroll(self.current_active_pane)
            else:
                self.sync_scroll(self.pane1)
        self.update_ui_state()

    def on_pane_scrolled(self, event, source_pane):
        if self.sync_scroll_enabled.get() and source_pane.pdf_document and (not source_pane.pdf_document.is_closed):
            self.sync_scroll(source_pane)

    def sync_scroll(self, source_pane):
        if not self.sync_scroll_enabled.get():
            return
        target_pane = self.pane2 if source_pane.pane_id == 'left' else self.pane1
        if not (source_pane.pdf_document and (not source_pane.pdf_document.is_closed) and target_pane.pdf_document and (not target_pane.pdf_document.is_closed)):
            return
        if target_pane.ignore_scroll_events_counter > 0:
            return
        source_x, source_y = source_pane.get_current_view_coords()
        source_canvas_height = source_pane.canvas.winfo_height()
        prev_scroll_time = self.scroll_time
        prev_scroll_y = self.scroll_y
        prev_scroll_pane = self.scroll_pane
        prev_scroll_height = self.scroll_height
        time_scroll = time.time()
        self.scroll_time = time_scroll
        self.scroll_y = source_y
        self.scroll_pane = source_pane
        self.scroll_height = source_canvas_height
        if source_canvas_height == 0:
            return
        first_common_word_in_view = None
        if not source_pane.sorted:
            source_pane.sorted = sorted(source_pane.words_data, key=lambda x: (x['page_num'], x['y0'], x['x0']))
        for word_info in source_pane.sorted:
            if word_info['unique_id'] is not None:
                page_num = word_info['page_num']
                word_y0_doc = word_info['y0']
                page_info = source_pane.page_layout_info.get(page_num)
                if not page_info:
                    continue
                word_y_content_coord = (page_info['y_start_offset'] + word_y0_doc) * source_pane.zoom_level
                if word_y_content_coord >= source_y - source_canvas_height * 0.01:
                    if word_y_content_coord < source_y + source_canvas_height:
                        first_common_word_in_view = word_info
                        break
        if first_common_word_in_view:
            common_word_id = first_common_word_in_view['unique_id']
            target_word_info = None
            for word_info_target in target_pane.words_data:
                if word_info_target['unique_id'] == common_word_id:
                    target_word_info = word_info_target
                    break
            if target_word_info:
                target_page_num = target_word_info['page_num']
                target_word_y0_doc = target_word_info['y0']
                target_page_info = target_pane.page_layout_info.get(target_page_num)
                if not target_page_info:
                    return
                source_word_y_content_coord_exact = (first_common_word_in_view['y0'] + source_pane.page_layout_info[first_common_word_in_view['page_num']]['y_start_offset']) * source_pane.zoom_level
                y_offset_in_source_view = source_word_y_content_coord_exact - source_y
                target_word_y_content_coord = (target_page_info['y_start_offset'] + target_word_y0_doc) * target_pane.zoom_level
                target_y_scroll_pixels = target_word_y_content_coord - y_offset_in_source_view
                prev_distance = self.scroll_distance
                distance = target_word_y_content_coord - source_y
                prev_target_y = self.scroll_target_y
                is_target_word_visible = target_word_y_content_coord > target_pane.get_current_view_coords()[1] and target_word_y_content_coord < target_pane.get_current_view_coords()[1] + target_pane.canvas.winfo_height()
                if (source_y - prev_scroll_y) * (target_y_scroll_pixels - prev_target_y) > 0 or not is_target_word_visible:
                    self.scroll_distance = distance
                    self.scroll_target_y = target_y_scroll_pixels
                    source_x_prop = source_pane.canvas.xview()[0]
                    target_x_scroll_pixels = source_x_prop * target_pane.max_document_width * target_pane.zoom_level
                    target_pane._apply_scroll(target_x_scroll_pixels, target_y_scroll_pixels)
        elif 0:
            source_x_prop, source_y_prop = (source_pane.canvas.xview()[0], source_pane.canvas.yview()[0])
            target_pane._apply_scroll(source_x_prop * target_pane.max_document_width * target_pane.zoom_level, source_y_prop * target_pane.total_document_height * target_pane.zoom_level)

    def sync_zoom(self, source_pane, new_zoom_level, mouse_x_canvas_pixel, mouse_y_canvas_pixel):
        if not self.sync_zoom_enabled.get():
            return
        target_pane = self.pane2 if source_pane.pane_id == 'left' else self.pane1
        if not (source_pane.pdf_document and (not source_pane.pdf_document.is_closed) and target_pane.pdf_document and (not target_pane.pdf_document.is_closed)):
            return
        target_pane.set_zoom(new_zoom_level, mouse_x_canvas_pixel, mouse_y_canvas_pixel, from_sync=True)

    def get_word_content_y(self, pane, word_info):
        if not pane.pdf_document or pane.pdf_document.is_closed:
            return -1
        page_num = word_info['page_num']
        page_info = pane.page_layout_info.get(page_num)
        if not page_info:
            return -1
        return (page_info['y_start_offset'] + word_info['y0']) * pane.zoom_level

    def is_word_visible(self, pane, word_info):
        if not pane.pdf_document or pane.pdf_document.is_closed:
            return False
        view_x, view_y = pane.get_current_view_coords()
        canvas_width = pane.canvas.winfo_width()
        canvas_height = pane.canvas.winfo_height()
        word_x0_content = (word_info['x0'] + (pane.max_document_width - pane.page_layout_info[word_info['page_num']]['base_width']) / 2) * pane.zoom_level
        word_y0_content = self.get_word_content_y(pane, word_info)
        word_x1_content = (word_info['x1'] + (pane.max_document_width - pane.page_layout_info[word_info['page_num']]['base_width']) / 2) * pane.zoom_level
        word_y1_content = (pane.page_layout_info[word_info['page_num']]['y_start_offset'] + word_info['y1']) * pane.zoom_level
        horizontal_overlap = not (word_x1_content < view_x or word_x0_content > view_x + canvas_width)
        vertical_overlap = not (word_y1_content < view_y or word_y0_content > view_y + canvas_height)
        return horizontal_overlap and vertical_overlap

    def _find_closest_change(self, direction):
        if not self.pdf_documents[0] or self.pdf_documents[0].is_closed or (not self.pdf_documents[1]) or self.pdf_documents[1].is_closed:
            messagebox.showinfo('Navigation Error', 'Both PDF documents must be loaded to navigate changes.')
            return None
        panes = [self.pane1, self.pane2]
        current_view_y_pane1 = self.pane1.get_current_view_coords()[1]
        current_view_y_pane2 = self.pane2.get_current_view_coords()[1]
        current_view_height_pane1 = self.pane1.get_current_view_height_in_content_coords()
        current_view_height_pane2 = self.pane2.get_current_view_height_in_content_coords()
        all_highlighted_words = []
        for pane_idx, pane in enumerate(panes):
            for word_idx, word_info in enumerate(pane.words_data):
                if word_info.get('highlight_color'):
                    abs_y_pos = self.get_word_content_y(pane, word_info)
                    all_highlighted_words.append({'pane': pane, 'word_info': word_info, 'abs_y_pos': abs_y_pos, 'pane_index': pane_idx, 'word_index': word_idx})
        if not all_highlighted_words:
            messagebox.showinfo('No Changes', 'No highlighted changes found in the documents.')
            return None
        all_highlighted_words.sort(key=lambda x: (x['word_info']['page_num'], x['abs_y_pos']))
        mid_y_pane1 = current_view_y_pane1 + current_view_height_pane1 / 2
        mid_y_pane2 = current_view_y_pane2 + current_view_height_pane2 / 2
        closest_unseen_change = None
        min_distance = float('inf')
        for change in all_highlighted_words:
            pane = change['pane']
            word_info = change['word_info']
            abs_y_pos = change['abs_y_pos']
            is_visible = self.is_word_visible(pane, word_info)
            if direction == 1:
                if abs_y_pos > pane.get_current_view_coords()[1] + pane.get_current_view_height_in_content_coords():
                    distance = abs_y_pos - (pane.get_current_view_coords()[1] + pane.get_current_view_height_in_content_coords())
                    if distance + 1e-05 < min_distance and distance > 0:
                        min_distance = distance
                        target_y_scroll = abs_y_pos
                        closest_unseen_change = {'pane': pane, 'word_info': word_info, 'target_y_scroll': target_y_scroll}
            elif abs_y_pos < pane.get_current_view_coords()[1]:
                distance = pane.get_current_view_coords()[1] - abs_y_pos
                if distance + 1e-05 < min_distance and distance > 0:
                    min_distance = distance
                    word_height_at_zoom = (word_info['y1'] - word_info['y0']) * pane.zoom_level
                    target_y_scroll = abs_y_pos - (pane.get_current_view_height_in_content_coords() - word_height_at_zoom)
                    closest_unseen_change = {'pane': pane, 'word_info': word_info, 'target_y_scroll': target_y_scroll}
        if not closest_unseen_change and all_highlighted_words:
            return None
            if direction == 1:
                for change in reversed(all_highlighted_words):
                    pane = change['pane']
                    word_info = change['word_info']
                    if not self.is_word_visible(pane, word_info):
                        abs_y_pos = self.get_word_content_y(pane, word_info)
                        target_y_scroll = abs_y_pos
                        return {'pane': pane, 'word_info': word_info, 'target_y_scroll': target_y_scroll}
                messagebox.showinfo('No More Changes', 'You are at the end of the document or all changes are currently visible.')
                return None
            else:
                for change in all_highlighted_words:
                    pane = change['pane']
                    word_info = change['word_info']
                    if not self.is_word_visible(pane, word_info):
                        abs_y_pos = self.get_word_content_y(pane, word_info)
                        word_height_at_zoom = (word_info['y1'] - word_info['y0']) * pane.zoom_level
                        target_y_scroll = abs_y_pos - (pane.get_current_view_height_in_content_coords() - word_height_at_zoom)
                        return {'pane': pane, 'word_info': word_info, 'target_y_scroll': target_y_scroll}
                messagebox.showinfo('No More Changes', 'You are at the beginning of the document or all changes are currently visible.')
                return None
        return closest_unseen_change

    def go_to_next_change(self):
        change_info = self._find_closest_change(direction=1)
        if change_info:
            target_pane = change_info['pane']
            target_word_info = change_info['word_info']
            target_y_scroll = change_info['target_y_scroll']
            current_x_prop = target_pane.canvas.xview()[0]
            target_x_scroll_pixels = current_x_prop * target_pane.max_document_width * target_pane.zoom_level
            target_pane._apply_scroll(target_x_scroll_pixels, target_y_scroll)
            self.sync_scroll(target_pane)

    def go_to_prev_change(self):
        change_info = self._find_closest_change(direction=-1)
        if change_info:
            target_pane = change_info['pane']
            target_word_info = change_info['word_info']
            target_y_scroll = change_info['target_y_scroll']
            current_x_prop = target_pane.canvas.xview()[0]
            target_x_scroll_pixels = current_x_prop * target_pane.max_document_width * target_pane.zoom_level
            target_pane._apply_scroll(target_x_scroll_pixels, target_y_scroll)
            self.sync_scroll(target_pane)

    def on_closing(self):
        self.pane1.close_pdf()
        self.pane2.close_pdf()
        self.master.destroy()
if __name__ == '__main__':
    root = TkinterDnD.Tk()
    app = PDFViewerApp(root)
    root.protocol('WM_DELETE_WINDOW', app.on_closing)
    root.mainloop()
