import tkinter as tk
from tkinter import filedialog
from docx2pdf import convert
from tkinter.messagebox import showinfo
from threading import Thread
from tkinter import ttk
import os

root = tk.Tk()
root.title("Конвертатор Word на PDF")

icon_path = 'settings.png'  # Replace with the path to your icon file
root.iconbitmap(icon_path)
# Load and set the window icon (replace 'icon.png' with your icon file)


def openfile():
    file_path = filedialog.askopenfilename(filetypes=[('Word Files', '*.docx')])
    if file_path:
        button.config(state="disabled")  # Disable the button during conversion
        convert_thread = Thread(target=convert_and_show_info, args=(file_path,))
        convert_thread.start()
    else:
        showinfo("Прозошла ошибка", "Попробуйте позже!")

def convert_and_show_info(file_path):
    try:
        progress_bar.start()  # Start the progress bar
        base_filename = os.path.splitext(os.path.basename(file_path))[0]
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        save_folder = os.path.join(desktop_path, 'Конвертированные PDF-файлы')
        os.makedirs(save_folder, exist_ok=True)  # Create the 'saved' folder if it doesn't exist
        pdf_output_path = os.path.join(save_folder, f"{base_filename}.pdf")
        convert(file_path, pdf_output_path)
        showinfo("Готово", "Файл успешно конвертирован! \nСохранено в папке <<Конвертированные PDF-файлы>>.")
    except Exception as e:
        showinfo("Ошибка", f"Не удалось конвертация: {str(e)}")
    finally:
        progress_bar.stop()  # Stop the progress bar
        button.config(state="normal")

label = tk.Label(root, text="Выберите файл")
label.grid(row=0, column=0, padx=5, pady=5)

button = tk.Button(root, text="Выбирать", width=30, command=openfile)
button.grid(row=0, column=1, padx=5, pady=5)

progress_bar = ttk.Progressbar(root, mode="indeterminate")
progress_bar.grid(row=1, column=0, columnspan=2, padx=5, pady=5)

root.mainloop()
