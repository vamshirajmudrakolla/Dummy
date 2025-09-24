# -*- coding: UTF-8 -*-
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
	on_windows=1
except:
	on_windows=0
	

try:
	import pyautogui
	PYAUTOGUI_AVAILABLE = True
except ImportError:
	PYAUTOGUI_AVAILABLE = False



# python -m venv myenv
# myenv\Scripts\activate
# python -m pip install Pillow klembord tkinterdnd2 pywin32 pyinstaller pymupdf pyautogui
# python myenv\Scripts\pywin32_postinstall.py -install
# ren pdf_viewer_app.py pdf_viewer_app.pyw
# pyinstaller --noconfirm pdf_viewer_app.pyw
#
# the following files can be removed from "dist" created by pinstaller:
#
# del dist\pdf_viewer_app\_internal\libcrypto-3.dll
# del dist\pdf_viewer_app\_internal\libssl-3.dll
# del dist\pdf_viewer_app\_internal\MSVCP140.dll
# del dist\pdf_viewer_app\_internal\unicodedata.pyd
# del dist\pdf_viewer_app\_internal\PIL\_avif.cp312-win_amd64.pyd
# del dist\pdf_viewer_app\_internal\PIL\_webp.cp312-win_amd64.pyd
# rmdir /s /q dist\pdf_viewer_app\_internal\_tcl_data\tzdata
# rmdir /s /q dist\pdf_viewer_app\_internal\_tcl_data\encoding
# rmdir /s /q dist\pdf_viewer_app\_internal\idlelib
# rmdir /s /q dist\pdf_viewer_app\_internal\setuptools
# rmdir /s /q dist\pdf_viewer_app\_internal\tcl8

# copy minimal git and .dll to the main directory (dist\pdf_viewer_app)
# then, 7z the folder

# not working:
# python -m nuitka --mode=standalone --enable-plugin=tk-inter pdf_viewer_app.py


TEMP_PDF_DIR = __import__("os").path.join(__import__("tempfile").gettempdir(), "pdf_viewer_temp_pdfs")
__import__("os").makedirs(TEMP_PDF_DIR, exist_ok=True)
os.makedirs(TEMP_PDF_DIR, exist_ok=True)
try:
	windll.user32.SetThreadDpiAwarenessContext(wintypes.HANDLE(-2))
except AttributeError:
	pass
def convert_clipboard_to_pdf(output_filename="clipboard_content.pdf"):
	"""
	Converts the HTML content from the clipboard to a PDF.
	If no HTML is found, it uses the plain text content.
	"""
	try:
		klembord.init()
	except RuntimeError:
		print("Error: Could not initialize klembord. Make sure a display server is running (e.g., X server on Linux).", file=sys.stderr)
		return None
	html_content = None
	plain_text_content = None
	try:
		plain_text_content, html_content = klembord.get_with_rich_text()
	except Exception as e:
		print(f"Warning: Could not retrieve rich text from clipboard: {e}", file=sys.stderr)
		print("Attempting to get plain text only.", file=sys.stderr)
		plain_text_content = klembord.get_text()
	content_to_use = ""
	if html_content:
		content_to_use = html_content
		if content_to_use.lower().find("<html")!=-1:
			content_to_use=content_to_use[content_to_use.lower().find("<html"):]
		elif content_to_use.lower().find("<head")!=-1:
			content_to_use=content_to_use[content_to_use.lower().find("<head"):]
		print("Using HTML content from clipboard.")
		#
		# Regex to find 'style="..."' or 'style='...'
		# And then replace 'background:' within that capture group
		# This is a more complex but more precise regex approach.
		# It looks for style attributes and then performs a sub-replacement inside the matched style content.
		def replace_style_content(match):
			style_content = match.group(1) # The content inside the style attribute quotes
			# Now, replace background: with background-color: within this specific style content
			# using a nested re.sub, case-insensitively
			new_style_content = re.sub(r'(?i)background:', 'background-color:', style_content)
			return f'style="{new_style_content}"' # Reconstruct the style attribute
		# This regex captures the content of the style attribute (between the quotes).
		# We handle both single and double quotes for the style attribute value.
		# It's still not perfect for all edge cases (e.g., mismatched quotes or escaped quotes within style)
		# but is much better than a global replace.
		content_to_use = re.sub(
			r'style=["\'](.*?)["\']', # Capture everything inside style="..." or style='...'
			replace_style_content,	# Use a function to process the captured content
			content_to_use,
			flags=re.DOTALL | re.IGNORECASE # DOTALL to match across newlines, IGNORECASE for 'style' itself
		)
	elif plain_text_content:
		content_to_use = f"""
		<html>
		<head>
			<style>
				body {{
					font-family: monospace; /* Often preferred for plain text */
					white-space: pre-wrap; /* Preserves whitespace and wraps long lines */
					word-wrap: break-word; /* Breaks long words if they don't fit */
				}}
			</style>
		</head>
		<body>
			<div>{plain_text_content}</div>
		</body>
		</html>
		"""
		print("Using plain text content from clipboard (wrapped in <pre> tags).")
	else:
		print("Clipboard is empty or contains no readable content.", file=sys.stderr)
		return None
	try:
		pathlib.Path(output_filename).parent.mkdir(parents=True, exist_ok=True)
		story = fitz.Story(html=content_to_use)  
		writer = fitz.DocumentWriter(output_filename)
		mediabox = fitz.paper_rect("a4")  
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
		print(f"Clipboard content converted to PDF: {output_filename}")
		return output_filename
	except Exception as e:
		print(f"Error converting clipboard content to PDF: {e}", file=sys.stderr)
		return None


def convert_word_to_pdf_no_markup(input_file_path, output_pdf_path=None):
	"""
	Converts a Word .docx, .doc, .rtf, or .txt file to a PDF by setting various view and print options
	to hide markup, then using the SaveAs method.
	The original file is not modified.
	Crucially, this function aims to ensure it does NOT interfere with any existing
	user-opened Word instances or unsaved user documents.
	Requires Microsoft Word to be installed on a Windows system.

	Args:
		input_file_path (str): The full path to the input file.
		output_pdf_path (str, optional): The full path for the output PDF file.
										  If None, a temporary name will be generated
										  in the TEMP_PDF_DIR.
	Returns:
		str: The path to the generated PDF file, or None if conversion fails.
	"""
	if not on_windows:
		print("Only working on Windows and requiring pythoncom and win32com.client")
		return
	
	input_file_path = input_file_path.replace("/", "\\")
	if output_pdf_path:
		output_pdf_path = output_pdf_path.replace("/", "\\")

	if not os.path.exists(input_file_path):
		print(f"Error: Input file not found at {input_file_path}")
		return None

	if output_pdf_path is None:
		base_name = os.path.splitext(os.path.basename(input_file_path))[0]
		output_pdf_path = os.path.join(TEMP_PDF_DIR, f"{base_name}_temp_{os.urandom(4).hex()}.pdf")

	os.makedirs(TEMP_PDF_DIR, exist_ok=True)

	wdFormatPDF = 17
	wdRevisionsViewFinal = 0

	word_app = None
	doc = None

	try:
		print(f"--- Starting conversion for '{input_file_path}' to '{output_pdf_path}' ---")
		pythoncom.CoInitialize()

		# Always dispatch a new instance of Word.
		word_app = win32com.client.DispatchEx("Word.Application")
		print("Attempted to launch a new Word instance for conversion.")
		
		word_app.Visible = False
		word_app.DisplayAlerts = False

		doc = word_app.Documents.Open(str(input_file_path))

		# Set the WarnBeforeSavingPrintingSendingMarkup option
		#
		# still to fix this
		# An error occurred during conversion: (-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'The WarnBeforeSavingPrintingSendingMarkup method or property is not available because the current document is read-only.', 'wdmain11.chm', 37373, -2146823683), None)
		# (when opening a read-only file) 
		#
		if hasattr(word_app.Options, 'WarnBeforeSavingPrintingSendingMarkup'):
			word_app.Options.WarnBeforeSavingPrintingSendingMarkup = False
			print("1. Set word_app.Options.WarnBeforeSavingPrintingSendingMarkup to False.")
		else:
			print("Warning: 'WarnBeforeSavingPrintingSendingMarkup' property not found. Skipping.")

		# Set revision view for the document opened in this specific instance
		if doc.ActiveWindow:
			doc.ActiveWindow.View.RevisionsView = wdRevisionsViewFinal
			print("2. Set document view to 'No Markup' (ActiveWindow.View.RevisionsView).")
		else:
			print("Warning: ActiveWindow not found for the document. Could not set revision view.")

		if hasattr(doc, 'ShowRevisions'):
			doc.ShowRevisions = False
			print("3. Set doc.ShowRevisions to False.")
		else:
			print("Warning: 'ShowRevisions' property not found on document. Skipping.")

		# Set print options for this specific Word instance
		if hasattr(word_app.Options, 'PrintRevisions'):
			word_app.Options.PrintRevisions = False
			print("4. Set word_app.Options.PrintRevisions to False.")
		else:
			print("Warning: 'PrintRevisions' property not found on word_app.Options. Skipping.")

		if hasattr(word_app.Options, 'PrintComments'):
			word_app.Options.PrintComments = False
			print("5. Set word_app.Options.PrintComments to False.")
		else:
			print("Warning: 'PrintComments' property not found on word_app.Options. Skipping.")
		
		if hasattr(word_app.Options, 'PrintHiddenText'):
			word_app.Options.PrintHiddenText = False
			print("6. Set word_app.Options.PrintHiddenText to False.")
		else:
			print("Warning: 'PrintHiddenText' property not found. Skipping.")

		if hasattr(word_app.Options, 'PrintDrawingObjects'):
			word_app.Options.PrintDrawingObjects = True # Usually want drawings
			print("7. Set word_app.Options.PrintDrawingObjects to True.")
		else:
			print("Warning: 'PrintDrawingObjects' property not found. Skipping.")

		doc.SaveAs(str(output_pdf_path), FileFormat=wdFormatPDF)
		print(f"Document saved as PDF to: {output_pdf_path}")
		
		# Close only the document that was opened by this script instance.
		doc.Close(SaveChanges=False) # SaveChanges=False is crucial
		print("Document closed within the isolated Word instance.")
		print("--- Conversion successful! ---")
		return output_pdf_path

	except Exception as e:
		print(f"An error occurred during conversion: {e}")
		print(f"Error details (e.args): {e.args}")
		try:
			excepinfo = pythoncom.GetErrorInfo()
			if excepinfo:
				print(f"COM Error Info: Source={excepinfo[0]}, Description={excepinfo[2]}")
		except Exception:
			pass
		return None

	finally:
		if word_app:
			try:
				word_app.Quit(SaveChanges=0) # wdDoNotSaveChanges = 0
				print("Isolated Word application instance quit.")
			except Exception as e:
				print(f"Error quitting Word application: {e}")

		pythoncom.CoUninitialize()



def extract_words_with_styles(pdf_document):
    """
    Fast path: extract words with minimal processing and simple sorting.
    This is significantly faster on large PDFs than grouping into lines/blocks.
    """
    all_words_data = []
    if not pdf_document or pdf_document.is_closed:
        return all_words_data

    ignore_lig = getattr(app, "ignore_ligatures", None)
    ignore_lig = (ignore_lig.get() if ignore_lig else False)

    for page_num, page in enumerate(pdf_document):
        try:
            if ignore_lig:
                words = page.get_text("words", flags=0)
            else:
                words = page.get_text("words")
        except Exception:
            continue

        try:
            words.sort(key=lambda w: (w[5], w[1], w[0]))
        except Exception:
            words.sort(key=lambda w: (w[1], w[0]))

        for w in words:
            x0, y0, x1, y1, word_text = w[:5]
            all_words_data.append({
                "text": word_text,
                "x0": x0, "y0": y0, "x1": x1, "y1": y1,
                "page_num": page_num,
                "font_family": "",
                "font_size": 12,
                "font_color": "#000000",
                "font_weight": "normal",
                "font_style": "normal",
                "unique_id": None,
                "highlight_color": None
            })
    return all_words_data


def helper_case_quotes(words_data1, words_data2, case_insensitive, ignore_quotes):
	a_compare = [word_info["text"] for word_info in words_data1]
	b_compare = [word_info["text"] for word_info in words_data2]
	if case_insensitive:
		a_compare = [word.lower() for word in a_compare]
		b_compare = [word.lower() for word in b_compare]
	if ignore_quotes:
		a_compare = [word.replace("â€˜", "'").replace("â€™", "'").replace("Ê¼", "'").replace('â€œ', '"').replace('â€', '"') for word in a_compare]
		b_compare = [word.replace("â€˜", "'").replace("â€™", "'").replace("Ê¼", "'").replace('â€œ', '"').replace('â€', '"') for word in b_compare]
	return a_compare,b_compare

def align_words_with_difflib(words_data1, words_data2, case_insensitive, ignore_quotes):#difflib (standard, uses Ratcliff-Obershelp algorithm)
	import time
	print(time.time(), "inizio align_words_with_difflib")
	"""
	Aligns two sequences of words using difflib.SequenceMatcher
	and assigns common IDs or marks as unique.
	Modifies words_data1 and words_data2 in place by setting 'unique_id'
	and 'highlight_color'.
	Args:
		words_data1 (list): List of dictionaries for words in document 1.
		words_data2 (list): List of dictionaries for words in document 2.
		case_insensitive (bool): If True, comparisons ignore case.
		ignore_quotes (bool): If True, various quote types are normalized to standard quotes.
	"""
	a_compare, b_compare = helper_case_quotes(words_data1, words_data2, case_insensitive, ignore_quotes)
	s = difflib.SequenceMatcher(None, a_compare, b_compare)
	common_word_id_counter = 0
	idx1_current = 0
	idx2_current = 0
	
	# log=open("outlog.txt","w", encoding="utf8")
	
	
	
	for tag, i1, i2, j1, j2 in s.get_opcodes():
		
		
		# sx=" ".join([x["text"] for x in words_data1[i1:i2]])
		# dx=" ".join([x["text"] for x in words_data2[j1:j2]])
		#if len(sx)>85: sx=sx[:40]+"..."+sx[-40:]
		#if len(dx)>85: dx=dx[:40]+"..."+dx[-40:]
		# try:
			# log.write(f"{tag}\t{i1}\t{i2}\t{j1}\t{j2}\t{sx}\t{" -> "}\t{dx}\n")
		# except:
			# traceback.print_exc()
			# raise
		
		
		
		if tag == 'equal':
			for k in range(i2 - i1):
				common_id = f"common-word-{common_word_id_counter}"
				words_data1[idx1_current + k]["unique_id"] = common_id
				words_data2[idx2_current + k]["unique_id"] = common_id
				words_data1[idx1_current + k]["highlight_color"] = None
				words_data2[idx2_current + k]["highlight_color"] = None
				common_word_id_counter += 1
			idx1_current += (i2 - i1)
			idx2_current += (j2 - j1)
		elif tag == 'delete': 
			for k in range(i2 - i1):
				words_data1[idx1_current + k]["unique_id"] = None
				words_data1[idx1_current + k]["highlight_color"] = "red"
			idx1_current += (i2 - i1)
		elif tag == 'insert': 
			for k in range(j2 - j1):
				words_data2[idx2_current + k]["unique_id"] = None
				words_data2[idx2_current + k]["highlight_color"] = "green"
			idx2_current += (j2 - j1)
		elif tag == 'replace': 
			for k in range(i2 - i1):
				words_data1[idx1_current + k]["unique_id"] = None
				words_data1[idx1_current + k]["highlight_color"] = "red"
			for k in range(j2 - j1):
				words_data2[idx2_current + k]["unique_id"] = None
				words_data2[idx2_current + k]["highlight_color"] = "green"
			idx1_current += (i2 - i1)
			idx2_current += (j2 - j1)
	print(time.time(), "fine align_words_with_difflib")
	return words_data1, words_data2
def apply_annotations_to_pdf_pages(pdf_document, words_data):
	if not pdf_document or pdf_document.is_closed:
		return
	words_by_page = defaultdict(list)
	for word in words_data:
		if word["highlight_color"]:
			words_by_page[word["page_num"]].append(word)
	for page_num in range(pdf_document.page_count):
		page = pdf_document.load_page(page_num)
		annotations_to_delete = [
			annot for annot in page.annots()
			if annot.type[0] == fitz.PDF_ANNOT_HIGHLIGHT and annot.info.get("title") == "PDFComparer"
		]
		for annot in annotations_to_delete:
			try:
				page.delete_annot(annot)
			except Exception as e:
				print(f"Error deleting old annotation on page {page_num}: {e}")
		page_words = words_by_page[page_num]
		if not page_words:
			print(time.time(), f"page {page_num}, number of annotations: 0 (no words to highlight)")
			continue
		#words_by_page is already sorted (with a smarter logic than this one commmented, for better detection of rows)
		#page_words.sort(key=lambda w: (w["y0"], w["x0"]))
		highlights_by_color = defaultdict(list)
		for word in page_words:
			rect = fitz.Rect(word["x0"], word["y0"], word["x1"], word["y1"])
			highlights_by_color[word["highlight_color"]].append(rect)
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
					if (abs(current_merged_rect.y0 - next_rect.y0) < y_tolerance and
						abs(current_merged_rect.y1 - next_rect.y1) < y_tolerance and
						next_rect.x0 <= current_merged_rect.x1 + x_tolerance): 
						current_merged_rect = current_merged_rect | next_rect 
					else:
						merged_rects.append(current_merged_rect)
						current_merged_rect = next_rect
				merged_rects.append(current_merged_rect) 
			highlight_color_rgb_float = (0.0, 0.0, 0.0)
			if color == "red":
				highlight_color_rgb_float = (1.0, 0.0, 0.0)
			elif color == "green":
				highlight_color_rgb_float = (0.0, 1.0, 0.0)
			elif color == "blue": 
				highlight_color_rgb_float = (0.0, 0.5, 1.0)
			for merged_rect in merged_rects:
				try:
					annot = page.add_highlight_annot(merged_rect)
					annot.set_colors(stroke=highlight_color_rgb_float)
					annot.set_opacity(0.3)
					annot.set_info(title="PDFComparer")
					annot.update()
					total_annotations_added += 1
				except Exception as e:
					print(f"Error adding merged highlight annotation on page {page_num}: {e}")
		print(time.time(), f"page {page_num}, number of annotations: {total_annotations_added} (from {len(page_words)} original words)")
	print(time.time(), "fine di " + inspect.currentframe().f_code.co_name)


class GitSequenceMatcher:
	def __init__(self, a, b, temp_dir=None):
		self.a = a
		self.b = b
		self.temp_file_a = None
		self.temp_file_b = None
		self.temp_dir = temp_dir

	def _create_temp_files(self):
		"""Creates temporary files with repr() of each item in the input sequences."""
		with tempfile.NamedTemporaryFile(mode='w+', delete=False, encoding='utf-8', dir=self.temp_dir) as f_a:
			self.temp_file_a = f_a.name
			for item in self.a:
				f_a.write(repr(item) + '\n')

		with tempfile.NamedTemporaryFile(mode='w+', delete=False, encoding='utf-8', dir=self.temp_dir) as f_b:
			self.temp_file_b = f_b.name
			for item in self.b:
				f_b.write(repr(item) + '\n')

	def _cleanup_temp_files(self):
		"""Removes the temporary files."""
		if self.temp_file_a and os.path.exists(self.temp_file_a):
			os.remove(self.temp_file_a)
		if self.temp_file_b and os.path.exists(self.temp_file_b):
			os.remove(self.temp_file_b)

	def get_opcodes(self):
		"""
		Generates a list of 6-tuple opcodes (tag, i1, i2, j1, j2, is_moved)
		similar to difflib.SequenceMatcher, with 'is_moved' flag for delete/insert.
		"""
		self._create_temp_files()
		process = None
		start_time=time.time()
		try:
			command = [
				'git',
				'--no-pager',
				'diff',
				#'--diff-algorithm=myers',
				#'--diff-algorithm=minimal',
				#'--diff-algorithm=patience',
				'--diff-algorithm=histogram',
				'--color=always',
				'--color-moved',
				'--unified=99999999',
				self.temp_file_a,
				self.temp_file_b
			]
			print(f"\nRunning command: {' '.join(command)}")
			process = subprocess.run(
				command,
				capture_output=True,
				text=True,
				encoding='utf-8',
				errors='replace'
			)

			#print(f"Git diff return code: {process.returncode}")
			#print("--- Raw Git Diff Output (repr) ---")
			#print(repr(process.stdout))
			#print("--- End Raw Git Diff Output ---")
			#print("Stderr from git (if any):")
			#print(process.stderr)
			#print("--- End Stderr ---")

			diff_output = process.stdout


			if process.returncode == 0 and not diff_output.strip():# in case the two files are equal and therefore git diff returns 0 and empty stdout
				with open(self.temp_file_a, 'r', encoding='utf-8', errors='replace') as f:
					num_lines = sum(1 for _ in f)
				return [('equal', 0, num_lines, 0, num_lines, False)]


			COLOR_RED_FG = r'\x1b\[31m'# deletions
			COLOR_GREEN_FG = r'\x1b\[32m'# insertions
			COLOR_BOLD_MAGENTA_FG = r'\x1b\[1;35m'# deletions (move)
			COLOR_BLUE_FG =		 r'\x1b\[1;34m'# deletions (move)
			COLOR_BOLD_CYAN_FG = r'\x1b\[1;36m'# insertions (move)
			COLOR_BOLD_YELLOW_FG = r'\x1b\[1;33m'# insertions (move)
			COLOR_RED_BG = r'\x1b\[41m'

			current_a_idx = 0
			current_b_idx = 0

			lines = diff_output.splitlines()
			in_hunk = False

			# Stores granular changes with internal tags and content
			# (internal_tag, content, a_start, a_end, b_start, b_end)
			granular_changes = []

			print("\n--- Line-by-Line Parsing Debug ---")
			for line_num, line in enumerate(lines):
				line_without_ansi = re.sub(r'\x1b\[[0-9;]*m', '', line)

				if line.startswith('\x1b[1mdiff --git'):
					in_hunk = True
					continue
				if not in_hunk:
					continue

				if line_without_ansi.strip().startswith('index ') or \
				   line_without_ansi.strip().startswith('--- a/') or \
				   line_without_ansi.strip().startswith('+++ b/'):
					continue
				
				if line_without_ansi.strip().startswith('@@'):
					match = re.match(r'@@ -(\d+)(?:,(\d+))? \+(\d+)(?:,(\d+))? @@', line_without_ansi.strip())
					if match:
						current_a_idx = int(match.group(1)) - 1
						current_b_idx = int(match.group(3)) - 1
					else:
						print(f"  -> ERROR: '@@' line did not match regex: {repr(line_without_ansi.strip())}")
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
				else:
					if not line_without_ansi.strip():
						continue
					else:
						print(f"  -> WARNING: Line not classified by any rule: {repr(line)}")
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

			#print("\n--- End Line-by-Line Parsing Debug ---")
			#print(f"DEBUG: granular_changes after first pass: {granular_changes}")

			# --- Move Detection and Marking (before consolidation) ---
			# Create a dictionary to map content to a list of its occurrences in source/dest
			moved_candidates = {} # content -> [(a_idx, b_idx, 'moved_delete'/'moved_insert', original_granular_idx)]

			for idx, (g_tag, g_content, g_a1, g_a2, g_b1, g_b2) in enumerate(granular_changes):
				if g_tag in ['moved_delete', 'moved_insert']:
					if g_content not in moved_candidates:
						moved_candidates[g_content] = []
					# Store (a_start, b_start, original_tag, original_index_in_granular_changes)
					# We use start indices (g_a1, g_b1) for simpler matching
					moved_candidates[g_content].append((g_a1, g_b1, g_tag, idx))

			# Keep track of granular indices that are part of a detected 'moved' pair
			# These will be marked with is_moved=True and treated as delete/insert ops
			is_moved_flags = {} # (original_granular_idx) -> True

			for content, candidates in moved_candidates.items():
				deletes = [c for c in candidates if c[2] == 'moved_delete']
				inserts = [c for c in candidates if c[2] == 'moved_insert']

				# Attempt to pair up deletes and inserts of the same content
				matched_deletes = set()
				matched_inserts = set()

				for d_a1, d_b1, d_tag, d_idx in deletes:
					if d_idx in matched_deletes: continue # Already used

					for i_a1, i_b1, i_tag, i_idx in inserts:
						if i_idx in matched_inserts: continue # Already used

						# If content matches and both are marked as moved_delete/insert by Git
						# We consider them a 'move'
						is_moved_flags[d_idx] = True
						is_moved_flags[i_idx] = True
						matched_deletes.add(d_idx)
						matched_inserts.add(i_idx)
						break # Found a match for this delete, move to next delete

			# --- Consolidation into difflib-style opcodes (with is_moved flag) ---
			final_opcodes_pre_replace = []
			
			current_tag = None
			current_i1, current_i2, current_j1, current_j2 = -1, -1, -1, -1
			current_is_moved_flag = False

			for idx, (g_tag, g_content, g_a1, g_a2, g_b1, g_b2) in enumerate(granular_changes):
				# Determine the actual output tag and its moved status
				# If a 'moved_delete' or 'moved_insert' was paired, its is_moved_flag is True
				# Otherwise, they become regular delete/insert.
				actual_tag = g_tag
				if actual_tag in ['moved_delete', 'moved_insert']:
					actual_tag = 'delete' if g_tag == 'moved_delete' else 'insert'

				is_moved_for_this_item = is_moved_flags.get(idx, False)

				# Initialize for the first valid change or a new type of change
				if current_tag is None:
					current_tag = actual_tag
					current_i1, current_i2 = g_a1, g_a2
					current_j1, current_j2 = g_b1, g_b2
					current_is_moved_flag = is_moved_for_this_item
					continue

				# Check if the current granular change can extend the current block
				can_extend = False
				# A block can only extend if the tags are the same, AND the is_moved_flag is the same
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
					# Corrected lines:
					current_i2 = g_a2
					current_j2 = g_b2
				else:
					# Current block ends, add it to final opcodes
					final_opcodes_pre_replace.append((current_tag, current_i1, current_i2, current_j1, current_j2, current_is_moved_flag))
					# Start a new block
					current_tag = actual_tag
					current_i1, current_i2 = g_a1, g_a2
					current_j1, current_j2 = g_b1, g_b2
					current_is_moved_flag = is_moved_for_this_item
			
			# Add the last block if any
			if current_tag is not None:
				final_opcodes_pre_replace.append((current_tag, current_i1, current_i2, current_j1, current_j2, current_is_moved_flag))

			# --- Post-consolidation for 'replace' operations ---
			# Now, merge adjacent delete and insert opcodes into 'replace' where appropriate.
			# This must happen after the initial consolidation of same-type operations.
			# Note: We do NOT convert 'moved' deletes/inserts into 'replace' if they are flagged as moved.
			# This is because 'replace' means content changed *in place*, while 'moved' means it changed location.

			consolidated_opcodes = []
			i = 0
			while i < len(final_opcodes_pre_replace):
				current_op = final_opcodes_pre_replace[i]
				tag, i1, i2, j1, j2, is_moved = current_op

				# Check for replace only if the current delete/insert is NOT marked as 'moved'
				if (tag == 'delete' and not is_moved) and i + 1 < len(final_opcodes_pre_replace):
					next_op = final_opcodes_pre_replace[i+1]
					next_tag, next_i1, next_i2, next_j1, next_j2, next_is_moved = next_op

					# If delete is immediately followed by an insert, and neither are marked as moved
					if (next_tag == 'insert' and not next_is_moved) and i2 == next_i1 and j2 == next_j1:
						# Create a 'replace' opcode. The is_moved flag for 'replace' is always False.
						consolidated_opcodes.append(('replace', i1, i2, j1, next_j2, False))
						i += 2 # Skip both current delete and the next insert
						continue
				
				consolidated_opcodes.append(current_op)
				i += 1

			# Final sort for consistent output order
			opcodes = sorted(consolidated_opcodes, key=lambda x: (x[1], x[3]))

		except Exception as e:
			print(f"An unexpected error occurred during parsing: {e}")
			traceback.print_exc()
			if process:
				print(f"Stdout from git: {process.stdout}")
				print(f"Stderr from git: {process.stderr}")
			return []
		finally:
			self._cleanup_temp_files()
		print("\n\n\n\n****time to run git diff and parse (in sec): ",time.time()-start_time)
		return opcodes



def align_words_with_git_diff(words_data1, words_data2, case_insensitive, ignore_quotes):
	a_compare, b_compare = helper_case_quotes(words_data1, words_data2, case_insensitive, ignore_quotes)
	s = GitSequenceMatcher(a_compare, b_compare,temp_dir='.')
	common_word_id_counter = 0
	idx1_current = 0
	idx2_current = 0
	for tag, i1, i2, j1, j2, is_moved in s.get_opcodes():
		# sx=" ".join([x["text"] for x in words_data1[i1:i2]])
		# dx=" ".join([x["text"] for x in words_data2[j1:j2]])
		if tag == 'equal':
			for k in range(i2 - i1):
				common_id = f"common-word-{common_word_id_counter}"
				words_data1[idx1_current + k]["unique_id"] = common_id
				words_data2[idx2_current + k]["unique_id"] = common_id
				words_data1[idx1_current + k]["highlight_color"] = None
				words_data2[idx2_current + k]["highlight_color"] = None
				common_word_id_counter += 1
			idx1_current += (i2 - i1)
			idx2_current += (j2 - j1)
		elif tag == 'delete' and not is_moved:
			for k in range(i2 - i1):
				words_data1[idx1_current + k]["unique_id"] = None
				words_data1[idx1_current + k]["highlight_color"] = "red"
			idx1_current += (i2 - i1)
		elif tag == 'insert' and not is_moved:
			for k in range(j2 - j1):
				words_data2[idx2_current + k]["unique_id"] = None
				words_data2[idx2_current + k]["highlight_color"] = "green"
			idx2_current += (j2 - j1)
		elif tag == 'replace':
			for k in range(i2 - i1):
				words_data1[idx1_current + k]["unique_id"] = None
				words_data1[idx1_current + k]["highlight_color"] = "red"
			for k in range(j2 - j1):
				words_data2[idx2_current + k]["unique_id"] = None
				words_data2[idx2_current + k]["highlight_color"] = "green"
			idx1_current += (i2 - i1)
			idx2_current += (j2 - j1)
		elif tag == 'insert' and is_moved: 
			for k in range(j2 - j1):
				words_data2[idx2_current + k]["unique_id"] = None 
				words_data2[idx2_current + k]["highlight_color"] = "blue"
				#print(idx2_current + k,words_data2[idx2_current + k]["text"])
			idx2_current += (j2 - j1)
		elif tag == 'delete' and is_moved:
			for k in range(i2 - i1):
				words_data1[idx1_current + k]["unique_id"] = None
				words_data1[idx1_current + k]["highlight_color"] = "blue"
			idx1_current += (i2 - i1)
	return words_data1, words_data2

def is_git_diff_available():
	"""
	Checks if git diff is available by running it with --quiet.
	"""
	try:
		subprocess.run(['git', 'diff'], check=False,  stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
		return True
	except (subprocess.CalledProcessError, FileNotFoundError):
		return False


align_words = align_words_with_difflib
print('git diff disabled; using difflib for speed')