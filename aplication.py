import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from PIL import Image, ImageTk
import base64
from io import BytesIO
import os
from main import main

# Ваша строка Base64
encoded_image = "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PPLAAAAJLKIAAAIIIAEE08844880000000000wwwAAAAAIIIY8833333/PvXzDDLLLDDH33/AP8A/wD/APzzxx1//wDf/wD3HHHHHDDDLKIJzy0089//AN9996yyywwwwwwwwwyyyyMMMOMOOOO+/wD/AH3333/LLLDAAQc880332w4LIIILLDDLb77HHHHHDDDDDCAIIIIIE031/wD+ywwgHPPPPPNsMMMM8w8wwwCCCCCCCGOOOIIV999y2qS1NNdu++yyxhBBBBHPPPPOOOOO88888889999++6G+rFNN99+yyDBBFMO++/zzzzzxxxxxxxzzz3+//wDvusgmPATfffOsogwRTPPv+8xzDDDHffffff8A/H3/AP8A/vPPDPOJLDUf/wD77KAEEQz/AP8A0hDHfffffusogggggggvfPPPPvvvvog043fPNHQQwlv+4BDff/usgggwQRTRTDDDDDDjjnvvPPOPfaQffv8A8sIBX33SNd776sEEECQzzzzv/PPPM4www777/wD/AP8A8Nbz30mAEO/7rtfzz0EAIJ7/AP8A+868csoggjPPPffffffdQQLP/wD/ACCDU89x9++jCCMd994kwwwNNNd88+PPPP8A/wD3/u7DAz33yAMNb7+Pxz2AAU//AG+DPP8AvvvPPPPMMP8A/PPvPOZqL7//AP8AAAQfOIz/ALrcAgX3wA1z7rLPOMMMIMY47PPPKII5/wD/AP8A/KIMUzjnVz7OIZ/vgY3zjHEEAYww577/AP8A/wA44457/wD/APw0sgzxXuORfPQQlv8AsI/zyEEERz//AH6ySmCGOO//AP8AvPPMMIIMc/3zkV3zEMZ7/I/7qEE1z3nG445z3/8A/wD/APMMMMMEEEIE3333A31zDsML/wDG/wCAAXfOY0jvv/8A+sIIIIIIIE00U3iAE3/nT+d/74IPbyWT3yIM/r+d775PEQIICU00013/xAAcEQEBAAIDAQEAAAAAAAAAAAARACBAEDBQYAH/2gAIAQMBAT8Q3SPkjIj4EiIiIiIiMiIiIiOSIiIiIiIiIiIjwyIiIiIiI7iNQ5IiIiIjuIyIjk8E5O0iOSMSIiIjQIzMyNAiNM5OTkiOCIiIiIiOwwIwI6yMDoNE0DAjMjIjcOwiNkzPFPpjqMDwSIiI3jMzN80CIzPHOk1jrPnzQ/N8iIv/xAAZEQEBAQEBAQAAAAAAAAAAAAARABAgMED/2gAIAQIBAT8QIiIjwIiODgiOSPAjg8CMIiNOTDSIiIiIiIiMI8CI6IiMMOSI4PAiIiIiIiIiMIiIiNIiIiIiIiIiIiIiIiIiI4IjSIiIiIjDCIjCI0iIwwiIiIiIiIiIiIiIiIiIiIiIiMIiIwiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIwiIiIiIiIiIiIiIiIiIwiIiIwiIiIiIiNIiIiIiIiIiIiIiIjCIiIjgiIiIiIw00jCIiIiIiIiIiIiIiIjkjCMIjCMMIiIwjTkjDSOTSMMMIw04MMI6IjSIiIwiMOjTxIjwODTxIjow+U0jw//xAAgEAEBAAIDAQEBAQEBAAAAAAARACEwECBAAVBgQTGB/9oACAEBAAE/EPz2ZnoRwzMzMz4Gfz3xMzM2LHJHZmZmZmbFixYsWNDMzMzMzMzO/PDM9v8AePszPpZvl9/LZmZmZsWLHVmdLMzM/wBKzMzM9cas2bNmzZs2bMzMzMzMzMzMzMz+wfnszMzycfejMzMzO9mZmf5d/jiIiPy2ZmZmZmZmZmZmZmZ4xoeGZmZmZnQ7GZmZmZmZmZmZmZmZmZ/sWZ8ebNmzZs2bNmzfeWZmZmZnhmZ8TMzMzMzMzMzMzMzMzOw+WP6NmZmZmZmZ/gWZmZmZmZmZn9VmZmZmZmZmZmZ6szPlZ5ZmZmf4VmZ3M/gM+NmZmdL53hmZmZnlmZ6vhZmZmZmZmZmZnxszMzPLwzO18DMzMzMzMzPuZmZmdrPDMzPDMzPvZmZ4zZ3/AHx/+zMzM62Z4ZmenzlmZmfGzMzPrdzMzMzod7wzPH2ZnozM+hmZnezOxmZmZn8/Ez7WfGzMzMzMzOhmZmZmZmZmeM+E4ZmZ8z0ZmZmfLixY3Ox7fdLM7nx/ZmeMzPciOT0MzrfMzMz+Y6Dys63ez1zzmzZs2fWzsZ0/dLOtnUzM/iN90ml0s33RnqzM/lMzrZ1OjGv77yIiNpHJw6DVjacszuxYsTM2bNmzZs2bNnS+EiNJ3dDPhZ6HGLHZmZmZmfUzqe5ueD8o8BHd5O5fezodLMz5mZmZ7kRrfATO775HlmbNmzZs9vsRERHGJmdf2OX1uhnwMzM6XjNnoRERtZ3s9n1szPlZmZ/DezO44Z/MZ0s9mdr3fTixY9P3wvd2PdmdLMzM+JmZ4+zydGZ1Oz71/wA5eX3szMzMzsZmdbOh7PZ2M92dzMzMzPRnoRERGt1Pd1vd7M6WZ8bMzMzqZ6vodv8AnnZmZmZnxY0/Y4fF96fe3/eHQ68WLHH3zMzOtnu9nq9Xs92dLMz4mZmZ8Doe71er2ez3ZnczMzMzPkZ0vZ4N71eDxZ/BI0Pd4enzh6fev3W8PV6szPpZ7up6/ev3h6ut6/dLPkeWZm+dfnT73+fen3v95+8/dX3r97fdD1dDMzo//9k="

def select_ekp_files():
    global ekp_files
    ekp_files = filedialog.askopenfilenames(
        title="Выберите файлы ЕКП",
        filetypes=[("Excel files", "*.xlsx *.xls *.xml")]
    )
    if ekp_files:
        ekp_label.config(text=f"Файл ЕКП: {', '.join([os.path.basename(file) for file in ekp_files])}")
    else:
        ekp_label.config(text="Файл ЕКП не выбран")

def select_gbk_files():
    global gbk_files
    gbk_files = filedialog.askopenfilenames(
        title="Выберите файлы ГБК",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if gbk_files:
        gbk_label.config(text=f"Файл ГБК: {', '.join([os.path.basename(file) for file in gbk_files])}")
    else:
        gbk_label.config(text="Файл ГБК не выбран")

def create_report_file():
    global ekp_files, gbk_files

    if not ekp_files:
        messagebox.showwarning("Предупреждение", "Вы не выбрали файл для ЕКП.")
        return

    if not gbk_files:
        messagebox.showwarning("Предупреждение", "Вы не выбрали файл для ГБК.")
        return

    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    
    if not save_path:
        messagebox.showwarning("Предупреждение", "Вы не выбрали путь для сохранения отчета.")
        return

    try:
        main(ekp_files, gbk_files, save_path)
        
        ekp_files = []
        gbk_files = []
        ekp_label.config(text="Файл ЕКП не выбран")
        gbk_label.config(text="Файл ГБК не выбран")

        messagebox.showinfo("Успех", f"Отчет сохранён: {save_path}")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось создать отчет: {e}")

# Основной root
root = tk.Tk()
root.title("Сверка ОСВ ГБК")
root.state('zoomed')

# Загрузка фонового изображения
image_data = base64.b64decode(encoded_image)
image = Image.open(BytesIO(image_data))
image = image.resize((root.winfo_screenwidth(), root.winfo_screenheight()), Image.Resampling.LANCZOS)
background_photo = ImageTk.PhotoImage(image)

background_label = tk.Label(root, image=background_photo)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

# Стилизация
style = ttk.Style()
style.configure("GlassButton.TButton", 
                font=("Arial", 18, "bold"),
                foreground="#4BA3C3",
                background="#4BA3C3",
                relief="flat",
                padding=10)
style.map("GlassButton.TButton",
          background=[("active", "#68D391")],
          foreground=[("active", "white")])

# Кнопки
ekp_button = ttk.Button(root, text="Выбрать файлы ЕКП", command=select_ekp_files, style="GlassButton.TButton")
ekp_button.place(relx=0.25, rely=0.4, anchor='center')

gbk_button = ttk.Button(root, text="Выбрать файлы ГБК", command=select_gbk_files, style="GlassButton.TButton")
gbk_button.place(relx=0.75, rely=0.4, anchor='center')

create_button = ttk.Button(root, text="Создать отчет", command=create_report_file, style="GlassButton.TButton")
create_button.place(relx=0.5, rely=0.6, anchor='center')

# Метки
ekp_label = tk.Label(root, text="Файлы ЕКП не выбраны", fg="white", bg="#4BA3C3", font=("Arial", 14))
ekp_label.place(relx=0.25, rely=0.45, anchor='center')

gbk_label = tk.Label(root, text="Файлы ГБК не выбраны", fg="white", bg="#4BA3C3", font=("Arial", 14))
gbk_label.place(relx=0.75, rely=0.45, anchor='center')

root.mainloop()
