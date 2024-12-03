import os
import PyPDF2
import fitz  # PyMuPDF
from docx import Document
import tkinter as tk
from tkinter import filedialog, messagebox
import shlex
import shelve
import tempfile
import atexit
import re

# Open the cache file
cache_file = shelve.open('search_cache')

# Create a temporary directory for storing images
temp_dir = tempfile.TemporaryDirectory()

def extract_snippet(text, keyword, context_size=30):
    keyword_lower = keyword.lower()
    start = max(text.lower().find(keyword_lower) - context_size, 0)
    end = min(start + len(keyword_lower) + 2 * context_size, len(text))
    snippet = text[start:end]
    return snippet

def search_pdf(file_path, keyword):
    results = []
    keyword_lower = keyword.lower()
    pattern = re.compile(re.escape(keyword_lower), re.IGNORECASE)
    with open(file_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            text = page.extract_text()
            if text and pattern.search(text.lower()):
                snippet = extract_snippet(text, keyword)
                results.append((file_path, page_num + 1, snippet))
    return results

def search_docx(file_path, keyword):
    results = []
    keyword_lower = keyword.lower()
    pattern = re.compile(re.escape(keyword_lower), re.IGNORECASE)
    doc = Document(file_path)
    for para in doc.paragraphs:
        if pattern.search(para.text.lower()):
            snippet = extract_snippet(para.text, keyword)
            results.append((file_path, None, snippet))
    return results

def search_files(directory, keyword):
    # Check if the keyword is in the cache
    if keyword in cache_file:
        return cache_file[keyword]

    results = []
    for root, _, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            if file.lower().endswith('.pdf'):
                results.extend(search_pdf(file_path, keyword))
            elif file.lower().endswith('.docx'):
                results.extend(search_docx(file_path, keyword))

    # Store the results in the cache
    cache_file[keyword] = results
    return results

def browse_directory():
    directory = filedialog.askdirectory()
    if directory:
        entry_directory.delete(0, tk.END)
        entry_directory.insert(0, directory)

def open_pdf(file_path, page_num, keyword):
    doc = fitz.open(file_path)
    page = doc.load_page(page_num - 1)  # Page numbers are 0-based in PyMuPDF
    text_instances = page.search_for(keyword, quads=True)
    
    for inst in text_instances:
        highlight = page.add_highlight_annot(inst)
        highlight.update()

    rect = page.rect
    zoom = 2  # Zoom factor
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, clip=rect)
    output = os.path.join(temp_dir.name, f"page_{page_num}.png")
    pix.save(output)
    os.system(f"open {output}")

def on_result_select(event):
    selection = event.widget.curselection()
    if selection:
        index = selection[0]
        result = event.widget.get(index)
        file_path, page_info = result.split(" - Page ")
        page_num, snippet = page_info.split(" - Snippet: ")
        page_num = int(page_num)
        keyword = entry_keyword.get()
        open_pdf(file_path, page_num, keyword)

def search():
    directory = entry_directory.get()
    keyword = entry_keyword.get()
    if not directory or not keyword:
        messagebox.showwarning("Input Error", "Please enter both directory and keyword.")
        return
    results = search_files(directory, keyword)
    listbox_results.delete(0, tk.END)
    if results:
        for file_path, page_num, snippet in results:
            if page_num:
                listbox_results.insert(tk.END, f"{file_path} - Page {page_num} - Snippet: {snippet}")
            else:
                listbox_results.insert(tk.END, f"{file_path} - Snippet: {snippet}")
    else:
        messagebox.showinfo("No Results", "No files found with the keyword.")

# Create the main window
root = tk.Tk()
root.title("PDF and DOCX Search")

# Ensure the window pops up in the foreground
root.attributes('-topmost', True)
root.update()
root.attributes('-topmost', False)
root.focus_force()
root.lift()

# Create and place the widgets
label_directory = tk.Label(root, text="Directory:")
label_directory.grid(row=0, column=0, padx=10, pady=10)

entry_directory = tk.Entry(root, width=50)
entry_directory.grid(row=0, column=1, padx=10, pady=10)

button_browse = tk.Button(root, text="Browse", command=browse_directory)
button_browse.grid(row=0, column=2, padx=10, pady=10)

label_keyword = tk.Label(root, text="Keyword:")
label_keyword.grid(row=1, column=0, padx=10, pady=10)

entry_keyword = tk.Entry(root, width=50)
entry_keyword.grid(row=1, column=1, padx=10, pady=10)

button_search = tk.Button(root, text="Search", command=search)
button_search.grid(row=1, column=2, padx=10, pady=10)

listbox_results = tk.Listbox(root, width=100, height=20)
listbox_results.grid(row=2, column=0, columnspan=3, padx=10, pady=10)
listbox_results.bind('<<ListboxSelect>>', on_result_select)

# Register a cleanup function to delete the temporary directory when the application exits
def cleanup():
    cache_file.close()
    temp_dir.cleanup()

atexit.register(cleanup)

# Start the main event loop
root.mainloop()
