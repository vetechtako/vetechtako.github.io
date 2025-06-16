import os
import csv
import tkinter as tk
from tkinter import filedialog, messagebox
import re

def parse_questions(text):
    lines = text.strip().splitlines()
    questions = []
    buffer = []
    answer = ''
    
    for line in lines:
        line = line.strip()
        if not line:
            continue

        # 偵測答案：單獨一行 A/B/C/D（支援原格式）
        if line.upper() in ['A', 'B', 'C', 'D'] and buffer:
            answer = line.upper()
            questions.append((buffer, answer))
            buffer = []
        else:
            buffer.append(line)
    
    # 處理最後一題（如果沒有空行結束）
    if buffer:
        # 檢查最後一行是否為答案
        if len(buffer) > 1 and buffer[-1].upper() in ['A', 'B', 'C', 'D']:
            answer = buffer[-1].upper()
            questions.append((buffer[:-1], answer))
        else:
            # 沒有找到答案，設為空
            questions.append((buffer, ''))

    return questions

def extract_components(question_lines):
    # 合併多行成一行以便解析
    combined = ' '.join(question_lines)
    
    # 支援兩種括號格式：(A) 和 （A）
    patterns = [
        r'（([ABCD])）',  # 全形括號格式
        r'\(([ABCD])\)'   # 半形括號格式
    ]
    
    question = ""
    options = ["", "", "", ""]
    
    for pattern in patterns:
        matches = list(re.finditer(pattern, combined))
        if len(matches) >= 4:
            # 找到題目部分（第一個選項之前的內容）
            first_match = matches[0]
            question = combined[:first_match.start()].strip()
            
            # 提取四個選項
            for i, match in enumerate(matches[:4]):
                option_letter = match.group(1)
                option_index = ord(option_letter) - ord('A')
                
                # 找到選項內容
                start = match.end()
                if i < len(matches) - 1:
                    end = matches[i + 1].start()
                else:
                    end = len(combined)
                
                option_text = combined[start:end].strip()
                options[option_index] = option_text
            
            break
    
    # 如果沒有找到完整的四個選項，返回None
    if not question or not all(options):
        print(f"錯誤解析題目：{question_lines}")
        return None

    return [question] + options

def convert_txt_to_csv(txt_path, csv_path):
    with open(txt_path, 'r', encoding='utf-8') as f:
        text = f.read()

    questions = parse_questions(text)

    with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['答案', '題目', '選項A', '選項B', '選項C', '選項D'])  # 欄位標題

        for q_lines, answer in questions:
            components = extract_components(q_lines)
            if components:
                writer.writerow([answer] + components)
            else:
                print(f"跳過無法解析的題目：{q_lines}")

def batch_process():
    input_dir = os.path.join(os.getcwd(), 'txt')
    output_dir = os.path.join(os.getcwd(), 'csv')

    if not os.path.exists(input_dir):
        messagebox.showerror("錯誤", f"找不到 txt 資料夾：{input_dir}")
        return
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    processed = 0
    total_questions = 0
    
    for filename in os.listdir(input_dir):
        if filename.endswith('.txt'):
            txt_path = os.path.join(input_dir, filename)
            csv_filename = os.path.splitext(filename)[0] + '.csv'
            csv_path = os.path.join(output_dir, csv_filename)
            
            # 計算題目數量
            with open(txt_path, 'r', encoding='utf-8') as f:
                text = f.read()
            questions = parse_questions(text)
            total_questions += len(questions)
            
            convert_txt_to_csv(txt_path, csv_path)
            processed += 1
            print(f"已轉換：{filename} ({len(questions)} 題)")

    messagebox.showinfo("完成", f"已成功轉換 {processed} 個檔案，共 {total_questions} 題")

def test_single_file():
    """測試單一檔案的功能"""
    file_path = filedialog.askopenfilename(
        title="選擇要測試的TXT檔案",
        filetypes=[("文本檔案", "*.txt")]
    )
    
    if file_path:
        output_path = file_path.replace('.txt', '_test.csv')
        convert_txt_to_csv(file_path, output_path)
        messagebox.showinfo("測試完成", f"測試檔案已儲存至：{output_path}")

def create_gui():
    root = tk.Tk()
    root.title("TXT to CSV 題庫轉換器 (改進版)")
    root.geometry("450x250")

    title_label = tk.Label(root, text="題庫轉換器", font=('Arial', 16, 'bold'))
    title_label.pack(pady=10)
    
    desc_label = tk.Label(root, text="支援 (A) 和 （A） 兩種括號格式", font=('Arial', 10))
    desc_label.pack(pady=5)

    info_label = tk.Label(root, text="將 txt 資料夾中的所有題庫轉換為 CSV 格式", font=('Arial', 12))
    info_label.pack(pady=20)

    # 批量轉換按鈕
    batch_btn = tk.Button(root, text="批量轉換", command=batch_process, 
                         font=('Arial', 14), bg='lightblue', width=15)
    batch_btn.pack(pady=10)

    # 測試單一檔案按鈕
    test_btn = tk.Button(root, text="測試單一檔案", command=test_single_file, 
                        font=('Arial', 12), bg='lightgreen', width=15)
    test_btn.pack(pady=5)

    root.mainloop()

if __name__ == '__main__':
    create_gui()