import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
from tkinter import filedialog
import os
import threading
from translator_core import ExcelTranslator, WordTranslator, PDFTranslator
from deep_translator import GoogleTranslator

class XlsxTranslatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QQy-Translator")
        self.root.geometry("600x750")
        
        # Variables
        self.file_path_var = ttk.StringVar()
        self.source_lang_var = ttk.StringVar(value="auto")
        self.target_lang_var = ttk.StringVar(value="en")
        self.progress_var = ttk.DoubleVar(value=0)
        self.is_translating = False
        self.translator = None

        # Get supported languages
        try:
            # Using a temporary instance to get languages
            langs_dict = GoogleTranslator(source='auto', target='en').get_supported_languages(as_dict=True)
            self.languages = list(langs_dict.keys())
            self.lang_codes = langs_dict
        except:
            # Fallback if offline or API changes
            self.languages = ['english', 'chinese (simplified)', 'japanese', 'korean', 'french', 'german', 'spanish']
            self.lang_codes = {l: l for l in self.languages} # Simplified mapping for fallback
        
        self.languages.sort()
        
        self.create_widgets()

    def create_widgets(self):
        # Main Container
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=BOTH, expand=YES)

        # Title
        title_label = ttk.Label(
            main_frame, 
            text="QQY 翻译器 (XLSX/DOCX/PDF)", 
            font=("Helvetica", 18, "bold"),
            bootstyle="primary"
        )
        title_label.pack(pady=(0, 20))

        # File Selection Section
        file_frame = ttk.Labelframe(main_frame, text="选择文件", padding="10")
        file_frame.pack(fill=X, pady=5)

        entry = ttk.Entry(file_frame, textvariable=self.file_path_var, state="readonly")
        entry.pack(side=LEFT, fill=X, expand=YES, padx=(0, 10))

        browse_btn = ttk.Button(
            file_frame, 
            text="浏览", 
            command=self.browse_file,
            bootstyle="secondary-outline"
        )
        browse_btn.pack(side=RIGHT)

        # Language Settings Section
        settings_frame = ttk.Labelframe(main_frame, text="翻译设置", padding="10")
        settings_frame.pack(fill=X, pady=10)

        # Grid layout for settings
        ttk.Label(settings_frame, text="源语言:").grid(row=0, column=0, padx=5, pady=5, sticky=W)
        source_combo = ttk.Combobox(
            settings_frame, 
            textvariable=self.source_lang_var, 
            values=['auto'] + self.languages,
            state="readonly"
        )
        source_combo.grid(row=0, column=1, padx=5, pady=5, sticky=EW)
        
        ttk.Label(settings_frame, text="目标语言:").grid(row=0, column=2, padx=5, pady=5, sticky=W)
        target_combo = ttk.Combobox(
            settings_frame, 
            textvariable=self.target_lang_var, 
            values=self.languages,
            state="readonly"
        )
        target_combo.grid(row=0, column=3, padx=5, pady=5, sticky=EW)
        
        settings_frame.columnconfigure(1, weight=1)
        settings_frame.columnconfigure(3, weight=1)

        # Progress Section
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=X, pady=10)

        self.progress_bar = ttk.Floodgauge(
            progress_frame, 
            bootstyle="success", 
            font=("Helvetica", 12), 
            mask="进度: {}%",
            variable=self.progress_var
        )
        self.progress_bar.pack(fill=X)

        # Log Area
        log_frame = ttk.Labelframe(main_frame, text="日志", padding="10")
        log_frame.pack(fill=BOTH, expand=YES, pady=5)

        self.log_text = ttk.Text(log_frame, height=8, state="disabled", font=("Consolas", 9))
        self.log_text.pack(fill=BOTH, expand=YES)

        # Control Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=X, pady=10)

        self.start_btn = ttk.Button(
            btn_frame, 
            text="开始翻译", 
            command=self.start_translation,
            bootstyle="primary",
            width=20
        )
        self.start_btn.pack(side=LEFT, padx=5)

        self.stop_btn = ttk.Button(
            btn_frame, 
            text="停止", 
            command=self.stop_translation,
            bootstyle="danger-outline",
            state="disabled"
        )
        self.stop_btn.pack(side=RIGHT, padx=5)

    def log(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert(END, message + "\n")
        self.log_text.see(END)
        self.log_text.config(state="disabled")

    def browse_file(self):
        filename = filedialog.askopenfilename(
            filetypes=[
                ("支持的文件", "*.xlsx *.docx *.pdf"),
                ("Excel 文件", "*.xlsx"),
                ("Word 文档", "*.docx"),
                ("PDF 文档", "*.pdf")
            ]
        )
        if filename:
            self.file_path_var.set(filename)
            self.log(f"已选择文件: {filename}")

    def toggle_controls(self, translating):
        self.is_translating = translating
        if translating:
            self.start_btn.config(state="disabled")
            self.stop_btn.config(state="normal")
            self.progress_bar.start()
        else:
            self.start_btn.config(state="normal")
            self.stop_btn.config(state="disabled")
            self.progress_bar.stop()

    def start_translation(self):
        input_path = self.file_path_var.get()
        if not input_path:
            Messagebox.show_error("请先选择一个文件。", "错误")
            return

        if not os.path.exists(input_path):
            Messagebox.show_error("未找到文件。", "错误")
            return

        source = self.source_lang_var.get()
        target = self.target_lang_var.get()
        
        # Map display name to code
        source_code = 'auto' if source == 'auto' else self.lang_codes.get(source, source)
        target_code = self.lang_codes.get(target, target)

        ext = os.path.splitext(input_path)[1].lower()
        if ext == '.xlsx':
            self.translator = ExcelTranslator(source_lang=source_code, target_lang=target_code)
            output_path = os.path.splitext(input_path)[0] + f"_translated_{target_code}.xlsx"
        elif ext == '.docx':
            self.translator = WordTranslator(source_lang=source_code, target_lang=target_code)
            output_path = os.path.splitext(input_path)[0] + f"_translated_{target_code}.docx"
        elif ext == '.pdf':
            self.translator = PDFTranslator(source_lang=source_code, target_lang=target_code)
            # PDF translations are saved as DOCX for better editability and layout preservation
            output_path = os.path.splitext(input_path)[0] + f"_translated_{target_code}.docx"
        else:
            Messagebox.show_error("不支持的文件格式。", "错误")
            return

        self.log("-" * 30)
        self.log(f"开始翻译: {source} -> {target}")
        self.log(f"输出文件: {output_path}")
        
        self.toggle_controls(True)
        self.progress_var.set(0)

        # Run in thread
        thread = threading.Thread(
            target=self.run_translation,
            args=(input_path, output_path)
        )
        thread.daemon = True
        thread.start()

    def run_translation(self, input_path, output_path):
        def update_progress(val):
            self.progress_var.set(val)
            self.progress_bar.configure(text=f"进度: {val:.1f}%")

        def update_status(msg):
            self.root.after(0, lambda: self.log(msg))

        success = self.translator.process_file(
            input_path, 
            output_path, 
            progress_callback=lambda v: self.root.after(0, lambda: update_progress(v)),
            status_callback=update_status
        )

        self.root.after(0, lambda: self.finish_translation(success, output_path))

    def finish_translation(self, success, output_path):
        self.toggle_controls(False)
        if success:
            Messagebox.show_info(f"翻译完成！\n已保存至: {output_path}", "成功")
        else:
            if not self.translator.stop_requested: # Only show error if not stopped manually
                Messagebox.show_error("翻译失败。请查看日志获取详细信息。", "错误")

    def stop_translation(self):
        if self.translator:
            self.translator.stop()
            self.log("正在停止翻译...")
            self.stop_btn.config(state="disabled") # Prevent double click

if __name__ == "__main__":
    # Setup theme
    app = ttk.Window(themename="flatly")
    XlsxTranslatorApp(app)
    app.mainloop()
