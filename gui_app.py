import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import json
import os
from pathlib import Path
import sys
import pythoncom

def get_resource_path(relative_path):
    """ Получает абсолютный путь к ресурсу """
    try:
        # PyInstaller создает временную папку и хранит путь в sys._MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Импорт ваших модулей
from kompas_connection import get_kompas_api7
from json_runner import JsonKompasRunner
from parse_models import KompasM3DParser

# Имя файла с промптом
PROMPT_FILE = get_resource_path("prompt.txt")  # Промпт читаем из ресурсов сборки
TEMP_JSON_FILE = os.path.abspath("temp_gui_model.json") # Временный файл сохраняем рядом с .exe

def copy_prompt_to_clipboard():
    """Считывает промпт из файла и копирует в буфер обмена."""
    if not os.path.exists(PROMPT_FILE):
        messagebox.showerror("Ошибка", f"Файл {PROMPT_FILE} не найден в папке проекта!")
        return
    
    try:
        with open(PROMPT_FILE, "r", encoding="utf-8") as f:
            prompt_text = f.read()
        root.clipboard_clear()
        root.clipboard_append(prompt_text)
        messagebox.showinfo("Скопировано", "Промпт успешно скопирован в буфер обмена! \nТеперь вы можете вставить его в ChatGPT/Claude.")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось прочитать файл:\n{e}")

def run_kompas_build(json_path):
    """Фоновая задача для запуска процесса построения."""
    def background_task():
        try:
            # ИНИЦИАЛИЗАЦИЯ COM ДЛЯ ФОНОВОГО ПОТОКА!
            pythoncom.CoInitialize() 
            
            # Отключаем актуальные кнопки (через root.after для безопасности UI)
            root.after(0, lambda: btn_clipboard_build.config(state=tk.DISABLED))
            root.after(0, lambda: btn_file_build.config(state=tk.DISABLED))
            
            module7, app, const7 = get_kompas_api7()
            app.Visible = True
            
            runner = JsonKompasRunner(module7, app, const7)
            runner.execute_script(json_path)
            
            root.after(0, lambda: messagebox.showinfo("Успех", "Модель успешно построена!"))
        except Exception as e:
            error_msg = str(e)
            root.after(0, lambda: messagebox.showerror("Ошибка построения", f"Произошла ошибка в КОМПАС-3D:\n{error_msg}"))
        finally:
            # Освобождаем COM (хорошая практика)
            pythoncom.CoUninitialize()
            
            # Включаем кнопки обратно
            root.after(0, lambda: btn_clipboard_build.config(state=tk.NORMAL))
            root.after(0, lambda: btn_file_build.config(state=tk.NORMAL))

    threading.Thread(target=background_task, daemon=True).start()

def build_from_clipboard():
    """Забирает JSON напрямую из буфера обмена, проверяет его и запускает построение."""
    try:
        # Пытаемся получить текст из буфера обмена
        json_content = root.clipboard_get().strip()
    except tk.TclError:
        # Ошибка возникает, если в буфере пусто или находится не текст (например, картинка или файл)
        messagebox.showwarning("Внимание", "Буфер обмена пуст или содержит не текстовые данные!")
        return

    if not json_content:
        messagebox.showwarning("Внимание", "Буфер обмена пуст!")
        return
    
    # Проверяем валидность JSON
    try:
        json.loads(json_content)
    except json.JSONDecodeError as e:
        messagebox.showerror("Ошибка JSON", f"В буфере обмена находится невалидный JSON:\n{e}\n\nПроверьте скопированный текст.")
        return
        
    # Сохраняем во временный файл и запускаем Компас
    try:
        with open(TEMP_JSON_FILE, "w", encoding="utf-8") as f:
            f.write(json_content)
        run_kompas_build(TEMP_JSON_FILE)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось создать временный файл:\n{e}")

def build_from_file():
    """Открывает диалог выбора файла и запускает построение."""
    file_path = filedialog.askopenfilename(
        title="Выберите JSON файл модели",
        filetypes=[("JSON файлы", "*.json"), ("Все файлы", "*.*")]
    )
    if file_path:
        run_kompas_build(file_path)

def parse_single_model():
    """Парсинг одного .m3d файла с выбором места сохранения JSON."""
    input_file = filedialog.askopenfilename(
        title="Выберите модель .m3d для парсинга",
        filetypes=[("КОМПАС-3D Деталь", "*.m3d"), ("Все файлы", "*.*")]
    )
    if not input_file:
        return

    output_file = filedialog.asksaveasfilename(
        title="Сохранить структуру как JSON",
        defaultextension=".json",
        initialfile=Path(input_file).stem + "_dump.json",
        filetypes=[("JSON файлы", "*.json")]
    )
    if not output_file:
        return

    def background_parse():
        try:
            btn_parse_single.config(state=tk.DISABLED, text="Парсинг...")
            parser = KompasM3DParser()
            
            model_data = parser.parse_model(input_file)
            if not model_data:
                root.after(0, lambda: messagebox.showerror("Ошибка", "Не удалось извлечь данные из модели."))
                return

            with open(output_file, "w", encoding="utf-8") as f:
                json.dump({Path(input_file).name: model_data}, f, ensure_ascii=False, indent=4)
                
            root.after(0, lambda: messagebox.showinfo("Успех", f"Модель успешно распарсена!\nСохранено в:\n{output_file}"))
        except Exception as e:
            root.after(0, lambda: messagebox.showerror("Ошибка парсинга", f"Произошла ошибка:\n{e}"))
        finally:
            root.after(0, lambda: btn_parse_single.config(state=tk.NORMAL, text="Парсинг одного файла .m3d"))
            
    threading.Thread(target=background_parse, daemon=True).start()

def parse_directory_models():
    """Парсинг всей папки с моделями."""
    input_dir = filedialog.askdirectory(title="Выберите папку с моделями .m3d")
    if not input_dir:
        return

    output_file = filedialog.asksaveasfilename(
        title="Сохранить общий дамп как JSON",
        defaultextension=".json",
        initialfile="kompas_directory_dump.json",
        filetypes=[("JSON файлы", "*.json")]
    )
    if not output_file:
        return

    def background_parse():
        try:
            btn_parse_dir.config(state=tk.DISABLED, text="Парсинг папки...")
            parser = KompasM3DParser()
            
            parsed_data = parser.parse_directory(input_dir)
            
            if not parsed_data:
                root.after(0, lambda: messagebox.showwarning("Внимание", "В выбранной папке не найдено файлов .m3d или парсинг не удался."))
                return

            with open(output_file, "w", encoding="utf-8") as f:
                json.dump(parsed_data, f, ensure_ascii=False, indent=4)
                
            root.after(0, lambda: messagebox.showinfo("Успех", f"Успешно обработано {len(parsed_data)} моделей!\nСохранено в:\n{output_file}"))
        except Exception as e:
            root.after(0, lambda: messagebox.showerror("Ошибка парсинга", f"Произошла ошибка:\n{e}"))
        finally:
            root.after(0, lambda: btn_parse_dir.config(state=tk.NORMAL, text="Парсинг папки с моделями"))

    threading.Thread(target=background_parse, daemon=True).start()

root = tk.Tk()
root.title("КОМПАС-3D AI Коннектор")
root.geometry("450x600")
root.resizable(False, False)
root.attributes('-topmost', True) 

# Блок 1: Промпт
frame_prompt = tk.LabelFrame(root, text="Шаг 1: Подготовка", padx=10, pady=10)
frame_prompt.pack(fill="x", padx=10, pady=5)

lbl_prompt = tk.Label(frame_prompt, text="Скопируйте системный промпт для нейросети:", justify=tk.LEFT)
lbl_prompt.pack(anchor="w")

btn_copy = tk.Button(frame_prompt, text="📋 Скопировать промпт", command=copy_prompt_to_clipboard, width=30)
btn_copy.pack(pady=5)

# Блок 2: Построение
frame_text = tk.LabelFrame(root, text="Шаг 2: Построение (Text to CAD)", padx=10, pady=10)
frame_text.pack(fill="x", padx=10, pady=5)

btn_clipboard_build = tk.Button(
    frame_text, 
    text="📋 Построить из буфера обмена", 
    command=build_from_clipboard, 
    bg="#4CAF50", 
    fg="white", 
    width=30
)
btn_clipboard_build.pack(pady=5)

btn_file_build = tk.Button(
    frame_text, 
    text="📁 Выбрать .json файл с диска", 
    command=build_from_file, 
    bg="#2196F3", 
    fg="white", 
    width=30
)
btn_file_build.pack(pady=5)

# Блок 3: Парсинг (CAD to JSON)
frame_parse = tk.LabelFrame(root, text="Шаг 3: Парсинг (CAD to JSON)", padx=10, pady=10)
frame_parse.pack(fill="x", padx=10, pady=5)

btn_parse_single = tk.Button(frame_parse, text="🔍 Парсинг одного файла .m3d", command=parse_single_model, bg="#FF9800", fg="white", width=30)
btn_parse_single.pack(pady=2)

btn_parse_dir = tk.Button(frame_parse, text="📂 Парсинг папки с моделями", command=parse_directory_models, bg="#FF9800", fg="white", width=30)
btn_parse_dir.pack(pady=2)

if __name__ == "__main__":
    root.mainloop()