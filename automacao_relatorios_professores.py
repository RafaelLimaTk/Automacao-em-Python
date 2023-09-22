from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from tkinter import filedialog, PhotoImage
from zipfile import ZipFile

import pandas as pd
import tkinter as tk
import os
import shutil

class ProfessorEvaluation:
    def __init__(self, professor_name, question, evaluation_data, evaluation_data_value, weighted_average):
        self.professor_name = professor_name
        self.question = question
        self.evaluation_data = evaluation_data
        self.evaluation_data_value = evaluation_data_value
        self.weighted_average = weighted_average

def create_excel_report_for_professor(professor_evaluation_list, output_file_path, folder_path):
    wb = Workbook()
    ws = wb.active

    for idx, prof_eval in enumerate(professor_evaluation_list):
        if idx > 0:
            ws = wb.create_sheet(title=f"Q{idx+1}")

        ws.append([f"QUestão: {prof_eval.question}"])
        ws.append([f"Professor: {prof_eval.professor_name}"])
        ws.append([])
        ws.append(['Características de Avaliação', 'Porcentagem'])
        ws.append(['Média ponderada', 'Média Ponderada ({})'.format(prof_eval.weighted_average)])

        for eval_char, percentage in prof_eval.evaluation_data.items():
            corresponding_key = next((key for key, value in prof_eval.evaluation_data_value.items() if value == percentage), None)
            if corresponding_key is not None:
                ws.append([eval_char, corresponding_key])

        chart = BarChart()
        chart.title = prof_eval.question
        chart.x_axis.title = prof_eval.professor_name
        chart.y_axis.title = "Porcentagem"

        chart.width = 25
        chart.height = 12

        data = Reference(ws, min_col=2, min_row=5, max_row=3 + len(prof_eval.evaluation_data), max_col=2)
        cats = Reference(ws, min_col=1, min_row=6, max_row=5 + len(prof_eval.evaluation_data))

        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)

        ws.add_chart(chart, f"H{1 + len(prof_eval.evaluation_data)}")

    save_path = os.path.join(folder_path, output_file_path)
    wb.save(save_path)

def process_excel_file(input_file_path, folder_path):
    df = pd.read_excel(input_file_path, header=None)
    question_rows = df[df.iloc[:, 0].str.contains("Q[0-9]+", na=False, regex=True)].index.tolist()
    all_professors_data = {}

    for i in range(len(question_rows) - 1):
        start_row = question_rows[i]
        end_row = question_rows[i + 1] if i < len(question_rows) - 1 else len(df)

        question = df.iloc[start_row, 0]

        evaluation_characteristics = df.iloc[start_row + 1, 1::2].dropna()
        evaluation_characteristics_value = df.iloc[start_row + 2, 1::2].dropna()
        evaluation_characteristics_value[evaluation_characteristics_value.index <= 11] *= 100

        professor_data = df.iloc[start_row + 2 : end_row].dropna(subset=[df.columns[0]])
        professor_data = professor_data.iloc[:, ::2]
        
        eval_chars_list = evaluation_characteristics.tolist()
        eval_value_list = evaluation_characteristics_value.tolist()
        for index, row in professor_data.iterrows():
            professor_name = row.iloc[0]
            evaluation_data = {eval_chars_list[i]: row.iloc[i+1] for i in range(len(eval_chars_list))}
            evaluation_data_value = {eval_value_list[i]: row.iloc[i+1] for i in range(len(eval_value_list))}
            weighted_average = row.iloc[-1]      
            prof_eval = ProfessorEvaluation(professor_name, question, evaluation_data, evaluation_data_value, weighted_average)
            
            if professor_name not in all_professors_data:
                all_professors_data[professor_name] = []
            
            all_professors_data[professor_name].append(prof_eval)

    for professor_name, professor_evaluation_list in all_professors_data.items():
        output_file_path = f"Avaliação_{professor_name.replace('/', '_').replace(' ', '_')}.xlsx"
        create_excel_report_for_professor(professor_evaluation_list, output_file_path, folder_path)

def download_zip_file(folder_name):
    save_path = filedialog.asksaveasfilename(initialfile=f"{folder_name}.zip", defaultextension=".zip", filetypes=[("ZIP files", "*.zip")])
    if not save_path:
        return
    shutil.copy(zip_file_path, save_path)

def select_and_process_files():
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
    if not file_paths:
        return
    
    os.makedirs("Avaliações", exist_ok=True)
    
    for file_path in file_paths:
        if not file_path.lower().endswith('.xlsx'):
            print(f"O arquivo {file_path} não é um arquivo .xlsx válido.")
            continue

    for file_path in file_paths:
        folder_name = os.path.splitext(os.path.basename(file_path))[0]
        folder_path = os.path.join("Avaliações", folder_name)
        os.makedirs(folder_path, exist_ok=True)

        process_excel_file(file_path, folder_path)

    global zip_file_path
    zip_file_path = os.path.join("Avaliações", f"{folder_name}.zip")
    with ZipFile(zip_file_path, 'w') as zipf:
        for root, _, files in os.walk(folder_path):
            for file in files:
                zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), folder_path))
    
    download_button.config(state=tk.NORMAL, command=lambda: download_zip_file(folder_name))
    return folder_name

root = tk.Tk()
root.title("Processador de Arquivos Excel")
root.geometry("450x560+500+200")
root.configure(bg="#f4fdfe")
root.resizable(False,False)

button_frame = tk.Frame(root)
button_frame.pack(side=tk.TOP, pady=20)
button_frame.configure(bg="#f4fdfe")

img_select_file_button = tk.PhotoImage(file="img/btn_selecionar_arquivo.png")
button = tk.Button(button_frame, image=img_select_file_button, command=select_and_process_files, borderwidth=0, highlightthickness=0, relief='flat')
button.pack(side=tk.LEFT, padx=10)

img_download_file_button = tk.PhotoImage(file="img/btn_baixar_zip.png")
download_button = tk.Button(button_frame, image=img_download_file_button, command=download_zip_file, state=tk.DISABLED, borderwidth=0, highlightthickness=0, relief='flat')
download_button.pack(side=tk.LEFT, padx=10)

root.mainloop()
