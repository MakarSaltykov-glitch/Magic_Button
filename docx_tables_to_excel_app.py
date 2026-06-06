from pathlib import Path
from queue import Queue
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd
from docx import Document


APP_TITLE = "Таблицы Word в Excel"
MAX_SHEET_NAME_LENGTH = 31
INVALID_SHEET_CHARS = r"[\[\]\:\*\?\/\\]"


def clean_cell_text(text: str) -> str:
    """Remove extra spaces, line breaks, and tabs from a Word table cell."""
    if not text:
        return ""
    return " ".join(text.split())


def make_unique_sheet_name(file_name: str, table_number: int, used_names: set[str]) -> str:
    """Build a valid unique Excel sheet name no longer than 31 characters."""
    base_name = Path(file_name).stem
    base_name = re.sub(INVALID_SHEET_CHARS, "_", base_name).strip()

    if not base_name:
        base_name = "Table"

    suffix = f"_tbl{table_number}"
    max_base_length = MAX_SHEET_NAME_LENGTH - len(suffix)
    sheet_name = f"{base_name[:max_base_length]}{suffix}"

    original_name = sheet_name
    counter = 1

    while sheet_name in used_names:
        counter_suffix = f"_{counter}"
        max_length = MAX_SHEET_NAME_LENGTH - len(counter_suffix)
        sheet_name = f"{original_name[:max_length]}{counter_suffix}"
        counter += 1

    used_names.add(sheet_name)
    return sheet_name


def table_to_dataframe(table) -> pd.DataFrame:
    """Convert a Word table to a DataFrame and skip fully empty rows."""
    rows = []

    for row in table.rows:
        cleaned_row = [clean_cell_text(cell.text) for cell in row.cells]

        if any(cleaned_row):
            rows.append(cleaned_row)

    if not rows:
        return pd.DataFrame()

    max_columns = max(len(row) for row in rows)
    normalized_rows = [row + [""] * (max_columns - len(row)) for row in rows]

    return pd.DataFrame(normalized_rows)


def convert_docx_tables_to_excel(input_folder: Path, output_excel_file: Path, log) -> int:
    """Extract all tables from all .docx files in a folder into one .xlsx file."""
    docx_files = sorted(
        file for file in input_folder.glob("*.docx") if not file.name.startswith("~$")
    )

    if not docx_files:
        raise RuntimeError("В выбранной папке нет файлов .docx.")

    output_excel_file.parent.mkdir(parents=True, exist_ok=True)
    used_sheet_names = set()
    written_tables = 0

    with pd.ExcelWriter(output_excel_file, engine="openpyxl") as writer:
        for docx_file in docx_files:
            try:
                document = Document(docx_file)
            except Exception as error:
                log(f"Ошибка: не удалось открыть '{docx_file.name}': {error}")
                continue

            if not document.tables:
                log(f"Файл '{docx_file.name}': таблицы не найдены.")
                continue

            for table_number, table in enumerate(document.tables, start=1):
                log(f"Обработка: {docx_file.name}, таблица №{table_number}")
                dataframe = table_to_dataframe(table)

                if dataframe.empty:
                    log(f"Пропуск пустой таблицы: {docx_file.name}, таблица №{table_number}")
                    continue

                sheet_name = make_unique_sheet_name(
                    file_name=docx_file.name,
                    table_number=table_number,
                    used_names=used_sheet_names,
                )

                dataframe.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    index=False,
                    header=False,
                )
                written_tables += 1

        if written_tables == 0:
            pd.DataFrame([["Таблицы не найдены или все таблицы пустые."]]).to_excel(
                writer,
                sheet_name="Info",
                index=False,
                header=False,
            )

    return written_tables


class ConverterApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("760x520")
        self.minsize(680, 440)

        self.input_folder_var = tk.StringVar(value=str(Path.home() / "Documents"))
        self.output_file_var = tk.StringVar(value=str(Path.home() / "Documents" / "tables.xlsx"))
        self.message_queue: Queue[tuple[str, str | int | None]] = Queue()
        self.worker_thread: threading.Thread | None = None

        self._build_ui()
        self._poll_queue()

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)

        title_label = ttk.Label(
            self,
            text="Конвертация таблиц Word (.docx) в Excel (.xlsx)",
            font=("Segoe UI", 14, "bold"),
        )
        title_label.grid(row=0, column=0, padx=16, pady=(16, 10), sticky="w")

        form_frame = ttk.Frame(self)
        form_frame.grid(row=1, column=0, padx=16, pady=6, sticky="ew")
        form_frame.columnconfigure(1, weight=1)

        ttk.Label(form_frame, text="Папка с .docx:").grid(row=0, column=0, padx=(0, 8), pady=6, sticky="w")
        ttk.Entry(form_frame, textvariable=self.input_folder_var).grid(row=0, column=1, pady=6, sticky="ew")
        ttk.Button(form_frame, text="Выбрать", command=self.choose_input_folder).grid(row=0, column=2, padx=(8, 0), pady=6)

        ttk.Label(form_frame, text="Excel-файл:").grid(row=1, column=0, padx=(0, 8), pady=6, sticky="w")
        ttk.Entry(form_frame, textvariable=self.output_file_var).grid(row=1, column=1, pady=6, sticky="ew")
        ttk.Button(form_frame, text="Сохранить как", command=self.choose_output_file).grid(row=1, column=2, padx=(8, 0), pady=6)

        control_frame = ttk.Frame(self)
        control_frame.grid(row=2, column=0, padx=16, pady=10, sticky="ew")
        control_frame.columnconfigure(1, weight=1)

        self.start_button = ttk.Button(control_frame, text="Начать конвертацию", command=self.start_conversion)
        self.start_button.grid(row=0, column=0, sticky="w")

        self.progress_bar = ttk.Progressbar(control_frame, mode="indeterminate")
        self.progress_bar.grid(row=0, column=1, padx=(12, 0), sticky="ew")

        log_frame = ttk.LabelFrame(self, text="Журнал обработки")
        log_frame.grid(row=3, column=0, padx=16, pady=(4, 16), sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, height=12, wrap="word", state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def choose_input_folder(self) -> None:
        folder = filedialog.askdirectory(title="Выберите папку с .docx-файлами")

        if folder:
            self.input_folder_var.set(folder)

    def choose_output_file(self) -> None:
        file_path = filedialog.asksaveasfilename(
            title="Куда сохранить Excel-файл",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )

        if file_path:
            self.output_file_var.set(file_path)

    def start_conversion(self) -> None:
        input_folder = Path(self.input_folder_var.get().strip())
        output_file = Path(self.output_file_var.get().strip())

        if not input_folder.exists() or not input_folder.is_dir():
            messagebox.showerror(APP_TITLE, "Выберите существующую папку с .docx-файлами.")
            return

        if output_file.suffix.lower() != ".xlsx":
            messagebox.showerror(APP_TITLE, "Выходной файл должен иметь расширение .xlsx.")
            return

        self._clear_log()
        self._set_running(True)
        self._append_log("Старт обработки...")

        self.worker_thread = threading.Thread(
            target=self._run_conversion,
            args=(input_folder, output_file),
            daemon=True,
        )
        self.worker_thread.start()

    def _run_conversion(self, input_folder: Path, output_file: Path) -> None:
        try:
            written_tables = convert_docx_tables_to_excel(
                input_folder=input_folder,
                output_excel_file=output_file,
                log=lambda message: self.message_queue.put(("log", message)),
            )
        except Exception as error:
            self.message_queue.put(("error", str(error)))
            return

        self.message_queue.put(("done", written_tables))

    def _poll_queue(self) -> None:
        while not self.message_queue.empty():
            message_type, payload = self.message_queue.get()

            if message_type == "log":
                self._append_log(str(payload))
            elif message_type == "error":
                self._set_running(False)
                self._append_log(f"Ошибка: {payload}")
                messagebox.showerror(APP_TITLE, str(payload))
            elif message_type == "done":
                self._set_running(False)
                self._append_log(f"Готово. Таблиц сохранено: {payload}")
                messagebox.showinfo(APP_TITLE, f"Готово. Таблиц сохранено: {payload}")

        self.after(100, self._poll_queue)

    def _append_log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _set_running(self, is_running: bool) -> None:
        if is_running:
            self.start_button.configure(state="disabled")
            self.progress_bar.start(10)
        else:
            self.progress_bar.stop()
            self.start_button.configure(state="normal")


if __name__ == "__main__":
    app = ConverterApp()
    app.mainloop()
