import os
import csv
import tkinter as tk
from tkinter import filedialog, messagebox

def parse_questions(text):
    lines = text.strip().splitlines()
    questions = []
    buffer = []
    answer = ''
    for line in lines:
        line = line.strip()
        if not line:
            continue

        # 偵測答案：單獨一行 A/B/C/D
        if line.upper() in ['A', 'B', 'C', 'D'] and buffer:
            answer = line.upper()
            questions.append((buffer, answer))
            buffer = []
        else:
            buffer.append(line)

    return questions

def extract_components(question_lines):
    # 合併多行成一行以便解析
    combined = ' '.join(question_lines)
    
    # 嘗試分割出題目與四個選項
    try:
        q_end = combined.index('(A)')
        question = combined[:q_end].strip()
        rest = combined[q_end:]

        options = []
        for tag in ['(A)', '(B)', '(C)', '(D)']:
            start = rest.index(tag) + len(tag)
            end = rest.index(f'({chr(ord(tag[1])+1)})') if tag != '(D)' else len(rest)
            options.append(rest[start:end].strip())
    except Exception as e:
        print(f"錯誤解析題目：{question_lines}，錯誤：{e}")
        return None

    return [question] + options

def convert_txt_to_csv(txt_path, csv_path):
    with open(txt_path, 'r', encoding='utf-8') as f:
        text = f.read()

    questions = parse_questions(text)

    with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['A', '題目', '選項A', '選項B', '選項C', '選項D'])  # 欄位標題

        for q_lines, answer in questions:
            components = extract_components(q_lines)
            if components:
                writer.writerow([answer] + components)

def batch_process():
    input_dir = os.path.join(os.getcwd(), 'txt')
    output_dir = os.path.join(os.getcwd(), 'csv')

    if not os.path.exists(input_dir):
        messagebox.showerror("錯誤", f"找不到 txt 資料夾：{input_dir}")
        return
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    processed = 0
    for filename in os.listdir(input_dir):
        if filename.endswith('.txt'):
            txt_path = os.path.join(input_dir, filename)
            csv_filename = os.path.splitext(filename)[0] + '.csv'
            csv_path = os.path.join(output_dir, csv_filename)
            convert_txt_to_csv(txt_path, csv_path)
            processed += 1

    messagebox.showinfo("完成", f"已成功轉換 {processed} 筆檔案")

def create_gui():
    root = tk.Tk()
    root.title("TXT to CSV 題庫轉換器")
    root.geometry("400x200")

    label = tk.Label(root, text="按下按鈕轉換 txt 資料夾中的所有題庫", font=('Arial', 12))
    label.pack(pady=30)

    convert_btn = tk.Button(root, text="開始轉換", command=batch_process, font=('Arial', 14), bg='lightblue')
    convert_btn.pack()

    root.mainloop()

if __name__ == '__main__':
    create_gui()
