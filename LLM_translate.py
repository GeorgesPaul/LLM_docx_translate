import docx
from docx import Document
import os
import logging
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import configparser
from pathlib import Path

logging.basicConfig(filename='translation_log.txt', level=logging.DEBUG)
CONFIG_FILE = 'translation_config.ini'

def load_config():
    config = configparser.ConfigParser()
    if os.path.exists(CONFIG_FILE):
        config.read(CONFIG_FILE)
    return config

def save_config(config):
    with open(CONFIG_FILE, 'w') as configfile:
        config.write(configfile)

def translate_text(text, api_url, api_key, target_language):
    # This is a placeholder - you'll need to implement the actual API call
    result = "kak"
    return result  # For now, just return a placeholder text

def translate_document(input_file, output_file, api_url, api_key, target_language):
    input_file = Path(input_file)
    output_file = Path(output_file)
    try:
        doc = Document(input_file)
        
        # Translate main document text
        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                if run.text.strip():
                    run.text = translate_text(run.text, api_url, api_key, target_language)
        
        # Translate text in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    # Process paragraphs within each cell
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            if run.text.strip():
                                run.text = translate_text(run.text, api_url, api_key, target_language)
                    
                    # Process any nested tables
                    for nested_table in cell.tables:
                        for nested_row in nested_table.rows:
                            for nested_cell in nested_row.cells:
                                for paragraph in nested_cell.paragraphs:
                                    for run in paragraph.runs:
                                        if run.text.strip():
                                            run.text = translate_text(run.text, api_url, api_key, target_language)
        
        # Translate text in text boxes (shapes)
        for shape in doc.inline_shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.text.strip():
                            run.text = translate_text(run.text, api_url, api_key, target_language)
        
        # Translate headers and footers
        for section in doc.sections:
            for header in section.header.paragraphs:
                for run in header.runs:
                    if run.text.strip():
                        run.text = translate_text(run.text, api_url, api_key, target_language)
            for footer in section.footer.paragraphs:
                for run in footer.runs:
                    if run.text.strip():
                        run.text = translate_text(run.text, api_url, api_key, target_language)
        
        doc.save(output_file)
        print(f"Successfully translated and saved: {output_file}")
    
    except Exception as e:
        print(f"Error in translate_document: {str(e)}")
        raise

class TranslationGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Document Translator")
        self.master.geometry("800x400")
        self.config = load_config()

        self.create_widgets()
        self.load_api_list()
        self.load_target_language()

    def create_widgets(self):
        # API List
        self.api_frame = ttk.LabelFrame(self.master, text="API Settings")
        self.api_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        self.api_list = tk.Listbox(self.api_frame, width=80, height=6)
        self.api_list.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.BOTH, expand=True)

        self.scroll = ttk.Scrollbar(self.api_frame, orient=tk.VERTICAL, command=self.api_list.yview)
        self.scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.api_list.config(yscrollcommand=self.scroll.set)

        self.delete_button = ttk.Button(self.api_frame, text="Delete Selected", command=self.delete_api)
        self.delete_button.pack(side=tk.RIGHT, padx=5, pady=5)

        # Add API
        self.add_frame = ttk.Frame(self.master)
        self.add_frame.pack(padx=10, pady=5, fill=tk.X)

        ttk.Label(self.add_frame, text="Name:").pack(side=tk.LEFT, padx=5)
        self.name_entry = ttk.Entry(self.add_frame, width=20)
        self.name_entry.pack(side=tk.LEFT, padx=5)

        ttk.Label(self.add_frame, text="URL:").pack(side=tk.LEFT, padx=5)
        self.url_entry = ttk.Entry(self.add_frame, width=30)
        self.url_entry.pack(side=tk.LEFT, padx=5)

        ttk.Label(self.add_frame, text="API Key:").pack(side=tk.LEFT, padx=5)
        self.key_entry = ttk.Entry(self.add_frame, width=20, show="*")
        self.key_entry.pack(side=tk.LEFT, padx=5)

        ttk.Button(self.add_frame, text="Add API", command=self.add_api).pack(side=tk.LEFT, padx=5)

        # File Selection
        self.file_frame = ttk.Frame(self.master)
        self.file_frame.pack(padx=10, pady=5, fill=tk.X)

        ttk.Label(self.file_frame, text="Input File:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.input_entry = ttk.Entry(self.file_frame, width=70)
        self.input_entry.grid(row=0, column=1, padx=5, pady=2)
        ttk.Button(self.file_frame, text="Browse", command=self.browse_input).grid(row=0, column=2, padx=5, pady=2)

        ttk.Label(self.file_frame, text="Output File:").grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.output_entry = ttk.Entry(self.file_frame, width=70)
        self.output_entry.grid(row=1, column=1, padx=5, pady=2)

        # Target Language
        ttk.Label(self.file_frame, text="Target Language:").grid(row=2, column=0, padx=5, pady=2, sticky="w")
        self.lang_entry = ttk.Entry(self.file_frame, width=20)
        self.lang_entry.grid(row=2, column=1, padx=5, pady=2, sticky="w")

        # Translation
        self.trans_frame = ttk.Frame(self.master)
        self.trans_frame.pack(padx=10, pady=5)

        self.open_var = tk.BooleanVar()
        ttk.Checkbutton(self.trans_frame, text="Open file when complete", variable=self.open_var).pack(side=tk.LEFT)
        ttk.Button(self.trans_frame, text="Do Translation", command=self.do_translation).pack(side=tk.LEFT, padx=5)

    def load_api_list(self):
        if 'APIs' in self.config:
            for name, details in self.config['APIs'].items():
                url, key = details.split(',')
                self.api_list.insert(tk.END, f"{name}: {url}")

    def load_target_language(self):
        if 'Settings' in self.config and 'target_language' in self.config['Settings']:
            self.lang_entry.insert(0, self.config['Settings']['target_language'])

    def add_api(self):
        name = self.name_entry.get()
        url = self.url_entry.get()
        key = self.key_entry.get()
        if name and url:
            if 'APIs' not in self.config:
                self.config['APIs'] = {}
            self.config['APIs'][name] = f"{url},{key}"
            save_config(self.config)
            self.api_list.insert(tk.END, f"{name}: {url}")
            self.name_entry.delete(0, tk.END)
            self.url_entry.delete(0, tk.END)
            self.key_entry.delete(0, tk.END)
        else:
            messagebox.showerror("Error", "Please fill all fields")

    def delete_api(self):
        selection = self.api_list.curselection()
        if selection:
            index = selection[0]
            name = self.api_list.get(index).split(':')[0]
            del self.config['APIs'][name]
            save_config(self.config)
            self.api_list.delete(index)
        else:
            messagebox.showerror("Error", "Please select an API to delete")

    def browse_input(self):
        filename = filedialog.askopenfilename(filetypes=[("Word Document", "*.docx")])
        if filename:
            input_path = Path(filename)
            if input_path.exists():
                self.input_entry.delete(0, tk.END)
                self.input_entry.insert(0, str(input_path))
                # Set default output filename
                output_path = input_path.with_stem(input_path.stem + "OUT")
                self.output_entry.delete(0, tk.END)
                self.output_entry.insert(0, str(output_path))
            else:
                messagebox.showerror("Error", f"Selected file does not exist: {input_path}")

    def do_translation(self):
        input_file = Path(self.input_entry.get())
        output_file = Path(self.output_entry.get())
        target_language = self.lang_entry.get()
        selection = self.api_list.curselection()
        
        if not input_file or not output_file or not target_language:
            messagebox.showerror("Error", "Please fill all fields")
            return
        
        if not selection:
            messagebox.showerror("Error", "Please select an API")
            return

        if not input_file.exists():
            error_message = f"Input file does not exist: {input_file}"
            logging.error(error_message)
            messagebox.showerror("Error", error_message)
            return

        if not input_file.is_file():
            error_message = f"Input path is not a file: {input_file}"
            logging.error(error_message)
            messagebox.showerror("Error", error_message)
            return

        name = self.api_list.get(selection[0]).split(':')[0]
        url, key = self.config['APIs'][name].split(',')

        # Save target language to config
        if 'Settings' not in self.config:
            self.config['Settings'] = {}
        self.config['Settings']['target_language'] = target_language
        save_config(self.config)

        try:
            translate_document(input_file, output_file, url, key, target_language)
            messagebox.showinfo("Success", "Translation complete!")
            if self.open_var.get():
                os.startfile(str(output_file))
        except Exception as e:
            error_message = f"An error occurred: {str(e)}"
            logging.error(error_message)
            messagebox.showerror("Error", error_message)

if __name__ == "__main__":
    root = tk.Tk()
    app = TranslationGUI(root)
    root.mainloop()