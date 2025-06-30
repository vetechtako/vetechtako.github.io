import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import csv
import json
import random
from datetime import datetime
import os
from pathlib import Path

class QuizSystem:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("選擇題練習系統")
        self.root.geometry("1024x768")
        self.root.configure(bg='#f0f0f0')
        
        # 數據存儲
        self.all_questions = []  # 一般隨機測驗用的100題
        self.complete_question_bank = []  # 完整題庫（新增）
        self.current_questions = []  # 當前測驗題目
        self.current_index = 0
        self.user_answers = {}
        self.wrong_questions = []
        self.quiz_mode = "random"  # "random", "review", "unique_random"
        self.questions_per_page = 50  # 每頁顯示題數
        self.current_page = 0  # 當前頁數
        
        # 創建題庫資料夾
        self.used_questions_file = "used_questions.json"
        self.used_wrong_questions_file = "used_wrong_questions.json"  # 新增：錯題複習的已使用記錄
        self.questions_dir = "questions"
        os.makedirs(self.questions_dir, exist_ok=True)
        
        # 創建界面
        self.setup_ui()
        self.load_questions_from_folder()
        
    def setup_ui(self):
        """設置用戶界面"""
        # 主框架
        main_frame = tk.Frame(self.root, bg='#f0f0f0')
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # 標題
        title_label = tk.Label(main_frame, text="選擇題練習系統", 
                              font=('Microsoft JhengHei', 24, 'bold'),
                              bg='#f0f0f0', fg='#2c3e50')
        title_label.pack(pady=(0, 20))
        
        # 按鈕框架
        button_frame = tk.Frame(main_frame, bg='#f0f0f0')
        button_frame.pack(pady=(0, 20))
        
        # 載入題庫按鈕
        load_btn = tk.Button(button_frame, text="重新載入題庫", 
                           command=self.load_questions_from_folder,
                           font=('Microsoft JhengHei', 12),
                           bg='#3498db', fg='white',
                           padx=20, pady=10)
        load_btn.pack(side='left', padx=(0, 10))
        
        # 循環式隨機測驗按鈕
        unique_btn = tk.Button(button_frame, text="循環式隨機100題測驗", 
                               command=self.start_unique_random_quiz,
                               font=('Microsoft JhengHei', 12),
                               bg='#8e44ad', fg='white',
                               padx=20, pady=10)
        unique_btn.pack(side='left', padx=(0, 10))

        # 開始隨機測驗按鈕
        start_btn = tk.Button(button_frame, text="隨機100題測驗", 
                            command=self.start_random_quiz,
                            font=('Microsoft JhengHei', 12),
                            bg='#27ae60', fg='white',
                            padx=20, pady=10)
        start_btn.pack(side='left', padx=(0, 10))
        
        # 循環式錯題複習按鈕（修改）
        review_btn = tk.Button(button_frame, text="循環式錯題複習", 
                             command=self.start_review_quiz,
                             font=('Microsoft JhengHei', 12),
                             bg='#e74c3c', fg='white',
                             padx=20, pady=10)
        review_btn.pack(side='left')
        
        # 題庫狀態標籤
        self.status_label = tk.Label(main_frame, text="正在載入questions/資料夾中的題庫...", 
                                   font=('Microsoft JhengHei', 10),
                                   bg='#f0f0f0', fg='#7f8c8d')
        self.status_label.pack(pady=(0, 10))
        
        # 創建滾動框架
        self.canvas = tk.Canvas(main_frame, bg='#f0f0f0')
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg='#f0f0f0')
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 綁定滑鼠滾輪事件
        self.root.bind("<MouseWheel>", self.on_mousewheel)
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)
        
        # 測驗內容框架
        self.quiz_frame = tk.Frame(self.scrollable_frame, bg='#f0f0f0')
        self.quiz_frame.pack(fill='both', expand=True)
        
    def on_mousewheel(self, event):
        """處理滑鼠滾輪事件"""
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
    def load_questions_from_folder(self):
        """從questions資料夾載入所有CSV題庫"""
        self.complete_question_bank = []  # 清空完整題庫
        self.all_questions = []
        csv_files = []
        
        # 搜尋questions資料夾中的所有CSV檔案
        if os.path.exists(self.questions_dir):
            for file_name in os.listdir(self.questions_dir):
                if file_name.lower().endswith('.csv'):
                    csv_files.append(os.path.join(self.questions_dir, file_name))
        
        if not csv_files:
            # 如果沒有找到CSV檔案，創建範例檔案
            self.create_sample_csv()
            csv_files = [os.path.join(self.questions_dir, "sample_questions.csv")]
        
        # 載入所有CSV檔案的題目
        all_question_banks = []
        total_questions = 0
        
        for csv_file in csv_files:
            try:
                with open(csv_file, 'r', encoding='utf-8', newline='') as file:
                    csv_reader = csv.reader(file)
                    questions = []
                    for row in csv_reader:
                        if len(row) >= 6:  # 答案,題目,選項1,選項2,選項3,選項4,詳解
                            questions.append(row)
                    
                    if questions:
                        all_question_banks.append({
                            'file': os.path.basename(csv_file),
                            'questions': questions
                        })
                        # 將所有題目加入完整題庫
                        self.complete_question_bank.extend(questions)
                        total_questions += len(questions)
                        
            except Exception as e:
                print(f"載入檔案失敗: {csv_file}, 錯誤: {str(e)}")
        
        if all_question_banks:
            # 為了保持原有的隨機測驗功能，從完整題庫中平均抽取100題
            questions_per_bank = max(1, 100 // len(all_question_banks))
            remaining_questions = 100 % len(all_question_banks)
            
            selected_questions = []
            
            for i, bank in enumerate(all_question_banks):
                # 計算這個題庫要抽取的題目數
                num_to_select = questions_per_bank
                if i < remaining_questions:
                    num_to_select += 1
                
                # 從題庫中隨機選取題目
                bank_questions = bank['questions']
                if len(bank_questions) <= num_to_select:
                    # 如果題庫題目不足，全部選取
                    selected_questions.extend(bank_questions)
                else:
                    # 隨機選取指定數量的題目
                    selected_questions.extend(random.sample(bank_questions, num_to_select))
            
            self.all_questions = selected_questions
            
            # 更新狀態顯示
            bank_info = []
            for bank in all_question_banks:
                bank_info.append(f"{bank['file']}({len(bank['questions'])}題)")
            
            status_text = f"已載入 {len(all_question_banks)} 個題庫，共 {total_questions} 題\n"
            status_text += f"題庫: {', '.join(bank_info)}\n"
            status_text += f"隨機測驗準備 {len(self.all_questions)} 題，循環測驗使用完整 {len(self.complete_question_bank)} 題"
            
            self.status_label.config(text=status_text)
        else:
            self.status_label.config(text="questions/資料夾中沒有找到有效的CSV題庫檔案")

    def create_sample_csv(self):
        """創建範例CSV檔案"""
        sample_questions = [
            ["D", "請問2025年5月15日是鳥羽水族館開館幾週年的紀念日？", "67週年", "68週年", "69週年", "70週年", "鳥羽水族館７０周年祝う!!"],
            ["B", "2025年加入Monterey Bay Aquarium的海獺名字是？", "Hazel", "Opal", "Quinn", "Quatse", "Meeeeep!!"],
            ["C", "請問下列哪一個英文詞彙專指海獺的群體？", "herd", "pod", "raft", "school", "Sea otter raft!!"],
            ["A", "《五感之外的世界》一書中，曾住在聖克魯茲（Santa Cruz）隆恩海洋實驗室（Long Marine Laboratory）參於研究的海獺名字是？", "Selka", "Ivy", "Rosa", "Kit", "https://doi.org/10.1242/jeb.181347"],
            ["C", "日本北海道ＮＰＯ法人エトピリカ基金的代表、攝影師片岡義廣曾經以一隻野生海獺的故事為主題出版攝影書籍，請問該書中的主角是哪一隻野生海獺？", "Ｇ子", "Ｆ子", "Ａ子", "Ｈ子", "《ラッコー霧多布で生まれたA子の物語》"],
        ]
        
        sample_file = os.path.join(self.questions_dir, "sample_questions.csv")
        try:
            with open(sample_file, 'w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                for question in sample_questions:
                    writer.writerow(question)
            print(f"已創建範例題庫: {sample_file}")
        except Exception as e:
            print(f"創建範例題庫失敗: {e}")

    def load_default_questions(self):
        """載入預設範例題目（保留以供備用）"""
        pass
        
    def load_questions(self):
        """從CSV檔案載入題目"""
        file_paths = filedialog.askopenfilenames(
            title="選擇題庫CSV檔案",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if not file_paths:
            return
            
        self.all_questions = []
        total_loaded = 0
        
        for file_path in file_paths:
            try:
                with open(file_path, 'r', encoding='utf-8', newline='') as file:
                    csv_reader = csv.reader(file)
                    file_questions = []
                    for row in csv_reader:
                        if len(row) >= 6:  # 答案,題目,選項1,選項2,選項3,選項4,詳解
                            file_questions.append(row)
                    
                    self.all_questions.extend(file_questions)
                    total_loaded += len(file_questions)
                    
            except Exception as e:
                messagebox.showerror("錯誤", f"載入檔案失敗: {file_path}\n錯誤: {str(e)}")
                
        if self.all_questions:
            self.status_label.config(text=f"已載入 {total_loaded} 道題目")
            messagebox.showinfo("成功", f"成功載入 {total_loaded} 道題目")
        else:
            messagebox.showwarning("警告", "沒有載入任何有效題目")
            
    def start_unique_random_quiz(self):
        """開始循環式隨機100題測驗"""
        if not self.complete_question_bank:
            messagebox.showwarning("警告", "沒有可用的題庫，請先載入題庫")
            return

        # 讀取已使用題目
        used_questions = set()
        if os.path.exists(self.used_questions_file):
            try:
                with open(self.used_questions_file, 'r', encoding='utf-8') as f:
                    used_questions = set(json.load(f))
            except:
                used_questions = set()

        # 過濾出未出現過的題目（使用完整題庫）
        remaining_questions = [q for q in self.complete_question_bank if q[1] not in used_questions]

        # 如果剩餘題目不足100題，檢查是否需要重置
        if len(remaining_questions) < 100:
            if len(used_questions) > 0:  # 確實有用過題目，才需要重置
                reset_msg = f"題庫已完成一輪循環！\n"
                reset_msg += f"總題目數：{len(self.complete_question_bank)} 題\n"
                reset_msg += f"已使用題目：{len(used_questions)} 題\n"
                reset_msg += f"剩餘題目：{len(remaining_questions)} 題\n\n"
                reset_msg += f"現在將重置循環，重新從完整題庫中抽取題目。"
                
                messagebox.showinfo("循環重置", reset_msg)
                
                # 重置已使用題目記錄
                used_questions = set()
                remaining_questions = self.complete_question_bank[:]

        # 確保有足夠的題目可選
        if len(remaining_questions) == 0:
            remaining_questions = self.complete_question_bank[:]

        # 隨機選取100題（或剩餘全部題目）
        num_questions = min(100, len(remaining_questions))
        selected = random.sample(remaining_questions, num_questions)
        
        self.current_questions = selected
        self.quiz_mode = "unique_random"
        self.current_page = 0

        # 更新已使用題目記錄
        for q in selected:
            used_questions.add(q[1])  # 使用題目文字作為唯一標識

        # 保存更新後的已使用題目記錄
        try:
            with open(self.used_questions_file, 'w', encoding='utf-8') as f:
                json.dump(list(used_questions), f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"寫入 used_questions.json 錯誤: {e}")

        # 顯示選取資訊
        info_msg = f"本次選取了 {len(selected)} 題\n"
        info_msg += f"題庫總數：{len(self.complete_question_bank)} 題\n"
        info_msg += f"已使用：{len(used_questions)} 題\n"
        info_msg += f"剩餘：{len(self.complete_question_bank) - len(used_questions)} 題"
        
        # 在狀態標籤中顯示循環進度
        current_status = self.status_label.cget("text")
        if "\n\n循環進度：" in current_status:
            # 如果已有進度資訊，更新它
            status_lines = current_status.split("\n\n循環進度：")[0]
        else:
            status_lines = current_status
        
        progress_text = status_lines + f"\n\n循環進度：{len(used_questions)}/{len(self.complete_question_bank)} 題已使用"
        self.status_label.config(text=progress_text)

        self.start_quiz()

    def start_random_quiz(self):
        """開始隨機測驗"""
        if not self.all_questions:
            messagebox.showwarning("警告", "沒有可用的題庫，請檢查questions/資料夾")
            return
            
        # 使用已經平均抽取的100題
        self.current_questions = self.all_questions[:]
        random.shuffle(self.current_questions)  # 打亂順序
        self.quiz_mode = "random"
        self.current_page = 0
        self.start_quiz()
        
    def start_review_quiz(self):
        """開始循環式錯題複習"""
        # 載入錯題記錄
        all_wrong_questions = self.load_wrong_questions()
        if not all_wrong_questions:
            messagebox.showinfo("提示", "沒有錯題記錄")
            return

        # 讀取已複習過的錯題
        used_wrong_questions = set()
        if os.path.exists(self.used_wrong_questions_file):
            try:
                with open(self.used_wrong_questions_file, 'r', encoding='utf-8') as f:
                    used_wrong_questions = set(json.load(f))
            except:
                used_wrong_questions = set()

        # 過濾出未複習過的錯題
        remaining_wrong_questions = [q for q in all_wrong_questions if q[1] not in used_wrong_questions]

        # 如果沒有剩餘的錯題，檢查是否需要重置
        if len(remaining_wrong_questions) == 0:
            if len(used_wrong_questions) > 0:  # 確實有復習過題目，才需要重置
                reset_msg = f"錯題已完成一輪循環複習！\n"
                reset_msg += f"總錯題數：{len(all_wrong_questions)} 題\n"
                reset_msg += f"已複習：{len(used_wrong_questions)} 題\n\n"
                reset_msg += f"現在將重置循環，重新從所有錯題中選取。"
                
                messagebox.showinfo("錯題循環重置", reset_msg)
                
                # 重置已複習錯題記錄
                used_wrong_questions = set()
                remaining_wrong_questions = all_wrong_questions[:]

        # 確保有足夠的題目可選
        if len(remaining_wrong_questions) == 0:
            remaining_wrong_questions = all_wrong_questions[:]

        # 選取錯題（最多100題，如果錯題不足100題就全部選取）
        num_questions = min(100, len(remaining_wrong_questions))
        if num_questions == len(remaining_wrong_questions):
            # 如果要全部選取，就不用隨機了
            selected_wrong = remaining_wrong_questions[:]
        else:
            # 隨機選取
            selected_wrong = random.sample(remaining_wrong_questions, num_questions)
        
        self.current_questions = selected_wrong
        self.quiz_mode = "review"
        self.current_page = 0

        # 顯示選取資訊
        info_msg = f"本次選取了 {len(selected_wrong)} 題錯題\n"
        info_msg += f"總錯題數：{len(all_wrong_questions)} 題\n"
        info_msg += f"已複習：{len(used_wrong_questions)} 題\n"
        info_msg += f"剩餘：{len(all_wrong_questions) - len(used_wrong_questions)} 題"
        
        messagebox.showinfo("錯題複習", info_msg)

        # 在狀態標籤中顯示錯題復習進度
        current_status = self.status_label.cget("text")
        if "\n\n錯題復習進度：" in current_status:
            # 如果已有進度資訊，更新它
            status_lines = current_status.split("\n\n錯題復習進度：")[0]
        else:
            status_lines = current_status
        
        progress_text = status_lines + f"\n\n錯題復習進度：{len(used_wrong_questions)}/{len(all_wrong_questions)} 題已複習"
        self.status_label.config(text=progress_text)

        self.start_quiz()
        
    def start_quiz(self):
        """開始測驗"""
        self.current_index = self.current_page * self.questions_per_page
        self.user_answers = {}
        self.clear_quiz_frame()
        self.show_questions_page()
        
    def clear_quiz_frame(self):
        """清空測驗框架"""
        for widget in self.quiz_frame.winfo_children():
            widget.destroy()
            
    def show_questions_page(self):
        """顯示當前頁的所有題目（50題）"""
        if self.current_page * self.questions_per_page >= len(self.current_questions):
            self.show_results()
            return
        
        # 計算當前頁的題目範圍
        start_idx = self.current_page * self.questions_per_page
        end_idx = min(start_idx + self.questions_per_page, len(self.current_questions))
        
        # 頁面標題
        page_frame = tk.Frame(self.quiz_frame, bg='#f0f0f0')
        page_frame.pack(fill='x', padx=20, pady=(0, 10))
        
        total_pages = (len(self.current_questions) + self.questions_per_page - 1) // self.questions_per_page
        page_title = f"第 {self.current_page + 1} 頁 / 共 {total_pages} 頁 (題目 {start_idx + 1} - {end_idx})"
        
        # 根據測驗模式調整標題
        if self.quiz_mode == "review":
            page_title = f"循環式錯題複習 - " + page_title
        elif self.quiz_mode == "unique_random":
            page_title = f"循環式隨機測驗 - " + page_title
        elif self.quiz_mode == "random":
            page_title = f"隨機測驗 - " + page_title
        
        title_label = tk.Label(page_frame, text=page_title,
                             font=('Microsoft JhengHei', 16, 'bold'),
                             bg='#f0f0f0', fg='#2c3e50')
        title_label.pack()
        
        # 存儲所有答案變數
        self.page_answer_vars = {}
        
        # 顯示當前頁的所有題目
        for i in range(start_idx, end_idx):
            question_data = self.current_questions[i]
            correct_answer, question_text, option_a, option_b, option_c, option_d, explanation = question_data
            
            # 題目框架
            question_frame = tk.Frame(self.quiz_frame, bg='white', relief='raised', bd=2)
            question_frame.pack(fill='x', padx=20, pady=10)
            
            # 題號和題目內容
            question_header = f"第 {i + 1} 題"
            header_label = tk.Label(question_frame, text=question_header,
                                  font=('Microsoft JhengHei', 12, 'bold'),
                                  bg='white', fg='#2c3e50')
            header_label.pack(pady=(15, 5), padx=20, anchor='w')
            
            question_label = tk.Label(question_frame, text=question_text,
                                    font=('Microsoft JhengHei', 12),
                                    bg='white', fg='#2c3e50',
                                    wraplength=800, justify='left')
            question_label.pack(pady=(0, 10), padx=20, anchor='w')
            
            # 選項
            answer_var = tk.StringVar()
            self.page_answer_vars[i] = answer_var
            
            options = [
                ('A', option_a),
                ('B', option_b),
                ('C', option_c),
                ('D', option_d)
            ]
            
            options_frame = tk.Frame(question_frame, bg='white')
            options_frame.pack(fill='x', padx=40, pady=(0, 15))
            
            for option_key, option_text in options:
                radio = tk.Radiobutton(options_frame, text=f"({option_key}) {option_text}",
                                     variable=answer_var, value=option_key,
                                     font=('Microsoft JhengHei', 11),
                                     bg='white', fg='#2c3e50',
                                     wraplength=700, justify='left')
                radio.pack(anchor='w', pady=2)
            
            # 恢復之前的答案
            if i in self.user_answers:
                answer_var.set(self.user_answers[i])
        
        # 頁面導航按鈕
        nav_frame = tk.Frame(self.quiz_frame, bg='#f0f0f0')
        nav_frame.pack(pady=20)
        
        # 上一頁按鈕
        if self.current_page > 0:
            prev_page_btn = tk.Button(nav_frame, text="上一頁",
                                    command=self.prev_page,
                                    font=('Microsoft JhengHei', 12),
                                    bg='#95a5a6', fg='white',
                                    padx=20, pady=10)
            prev_page_btn.pack(side='left', padx=(0, 10))
        
        # 保存當前頁答案按鈕
        save_btn = tk.Button(nav_frame, text="保存當前頁答案",
                           command=self.save_current_page,
                           font=('Microsoft JhengHei', 12),
                           bg='#f39c12', fg='white',
                           padx=20, pady=10)
        save_btn.pack(side='left', padx=(0, 10))
        
        # 下一頁/提交按鈕
        if end_idx < len(self.current_questions):
            next_page_btn = tk.Button(nav_frame, text="下一頁",
                                    command=self.next_page,
                                    font=('Microsoft JhengHei', 12),
                                    bg='#3498db', fg='white',
                                    padx=20, pady=10)
            next_page_btn.pack(side='left')
        else:
            submit_btn = tk.Button(nav_frame, text="提交答案",
                                 command=self.submit_quiz,
                                 font=('Microsoft JhengHei', 12),
                                 bg='#27ae60', fg='white',
                                 padx=20, pady=10)
            submit_btn.pack(side='left')
        
        # 滾動到頂部
        self.root.after(100, lambda: self.canvas.yview_moveto(0))
    
    def save_current_page(self):
        """保存當前頁的所有答案"""
        start_idx = self.current_page * self.questions_per_page
        end_idx = min(start_idx + self.questions_per_page, len(self.current_questions))
        
        saved_count = 0
        for i in range(start_idx, end_idx):
            if i in self.page_answer_vars and self.page_answer_vars[i].get():
                self.user_answers[i] = self.page_answer_vars[i].get()
                saved_count += 1
        
        messagebox.showinfo("保存成功", f"已保存 {saved_count} 題答案")
    
    def prev_page(self):
        """上一頁"""
        self.save_current_page_silently()
        self.current_page -= 1
        self.clear_quiz_frame()
        self.show_questions_page()
    
    def next_page(self):
        """下一頁"""
        self.save_current_page_silently()
        self.current_page += 1
        self.clear_quiz_frame()
        self.show_questions_page()
    
    def save_current_page_silently(self):
        """靜默保存當前頁答案（不顯示提示）"""
        start_idx = self.current_page * self.questions_per_page
        end_idx = min(start_idx + self.questions_per_page, len(self.current_questions))
        
        for i in range(start_idx, end_idx):
            if i in self.page_answer_vars and self.page_answer_vars[i].get():
                self.user_answers[i] = self.page_answer_vars[i].get()
        
    def submit_quiz(self):
        """提交測驗"""
        # 保存當前頁的答案
        self.save_current_page_silently()
            
        # 檢查是否所有題目都已作答
        unanswered = []
        for i in range(len(self.current_questions)):
            if i not in self.user_answers:
                unanswered.append(i + 1)
                
        if unanswered:
            if len(unanswered) <= 10:
                unanswered_str = "、".join(map(str, unanswered))
                message = f"第 {unanswered_str} 題未作答，確定要提交嗎？"
            else:
                message = f"還有 {len(unanswered)} 題未作答，確定要提交嗎？"
                
            result = messagebox.askyesno("確認提交", message)
            if not result:
                return
                
        self.show_results()
        
    def show_results(self):
        """顯示測驗結果"""
        self.clear_quiz_frame()
        
        # 計算成績
        correct_count = 0
        self.wrong_questions = []
        
        for i, question_data in enumerate(self.current_questions):
            correct_answer = question_data[0]
            user_answer = self.user_answers.get(i, "")
            
            if user_answer == correct_answer:
                correct_count += 1
            else:
                self.wrong_questions.append({
                    'question_index': i,
                    'question_data': question_data,
                    'user_answer': user_answer,
                    'correct_answer': correct_answer
                })
                
        total_questions = len(self.current_questions)
        score = (correct_count / total_questions) * 100 if total_questions > 0 else 0
        
        # 如果是錯題複習模式，更新已複習題目記錄
        if self.quiz_mode == "review":
            self.update_reviewed_wrong_questions()
        
        # 結果標題
        result_frame = tk.Frame(self.quiz_frame, bg='#f0f0f0')
        result_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # 根據測驗模式調整標題
        if self.quiz_mode == "review":
            title_text = "循環式錯題複習結果"
        elif self.quiz_mode == "unique_random":
            title_text = "循環式隨機測驗結果"
        else:
            title_text = "測驗結果"
            
        title_label = tk.Label(result_frame, text=title_text,
                             font=('Microsoft JhengHei', 24, 'bold'),
                             bg='#f0f0f0', fg='#2c3e50')
        title_label.pack(pady=(0, 20))
        
        # 成績顯示
        score_text = f"答對: {correct_count} / {total_questions} 題\n成績: {score:.1f} 分"
        score_label = tk.Label(result_frame, text=score_text,
                             font=('Microsoft JhengHei', 18),
                             bg='#f0f0f0', fg='#e74c3c' if score < 60 else '#27ae60')
        score_label.pack(pady=(0, 20))
        
        # 按鈕框架
        button_frame = tk.Frame(result_frame, bg='#f0f0f0')
        button_frame.pack(pady=(0, 20))
        
        # 重新開始按鈕
        restart_btn = tk.Button(button_frame, text="重新開始",
                              command=self.restart_quiz,
                              font=('Microsoft JhengHei', 12),
                              bg='#3498db', fg='white',
                              padx=20, pady=10)
        restart_btn.pack(side='left', padx=(0, 10))
        
        # 匯出錯題按鈕
        if self.wrong_questions:
            export_btn = tk.Button(button_frame, text="匯出錯題",
                                 command=self.export_wrong_questions,
                                 font=('Microsoft JhengHei', 12),
                                 bg='#e74c3c', fg='white',
                                 padx=20, pady=10)
            export_btn.pack(side='left')
            
        # 顯示錯題詳解
        if self.wrong_questions:
            self.show_wrong_questions_detail(result_frame)
            
        # 保存錯題記錄（只有在隨機測驗模式下才保存）
        if self.quiz_mode in ["random", "unique_random"] and self.wrong_questions:
            self.save_wrong_questions()
            
        # 滾動到頂部
        self.root.after(100, lambda: self.canvas.yview_moveto(0))

    def update_reviewed_wrong_questions(self):
        """更新已複習過的錯題記錄"""
        try:
            # 讀取現有的已複習錯題記錄
            used_wrong_questions = set()
            if os.path.exists(self.used_wrong_questions_file):
                try:
                    with open(self.used_wrong_questions_file, 'r', encoding='utf-8') as f:
                        used_wrong_questions = set(json.load(f))
                except:
                    used_wrong_questions = set()

            # 將本次複習的題目加入已複習記錄
            for question_data in self.current_questions:
                used_wrong_questions.add(question_data[1])  # 使用題目文字作為唯一標識

            # 保存更新後的已複習錯題記錄
            with open(self.used_wrong_questions_file, 'w', encoding='utf-8') as f:
                json.dump(list(used_wrong_questions), f, ensure_ascii=False, indent=2)

            # 更新狀態標籤中的錯題復習進度
            all_wrong_questions = self.load_wrong_questions()
            current_status = self.status_label.cget("text")
            if "\n\n錯題復習進度：" in current_status:
                status_lines = current_status.split("\n\n錯題復習進度：")[0]
            else:
                status_lines = current_status
            
            progress_text = status_lines + f"\n\n錯題復習進度：{len(used_wrong_questions)}/{len(all_wrong_questions)} 題已複習"
            self.status_label.config(text=progress_text)

        except Exception as e:
            print(f"更新已複習錯題記錄失敗: {e}")
        
    def show_wrong_questions_detail(self, parent_frame):
        """顯示錯題詳解"""
        detail_label = tk.Label(parent_frame, text="錯題詳解",
                              font=('Microsoft JhengHei', 18, 'bold'),
                              bg='#f0f0f0', fg='#e74c3c')
        detail_label.pack(pady=(20, 10))
        
        for wrong_q in self.wrong_questions:
            question_data = wrong_q['question_data']
            user_answer = wrong_q['user_answer']
            correct_answer = wrong_q['correct_answer']
            
            correct_answer_text, question_text, option_a, option_b, option_c, option_d, explanation = question_data
            
            # 錯題框架
            wrong_frame = tk.Frame(parent_frame, bg='white', relief='raised', bd=2)
            wrong_frame.pack(fill='x', pady=10)
            
            # 題目
            q_label = tk.Label(wrong_frame, text=f"題目: {question_text}",
                             font=('Microsoft JhengHei', 12, 'bold'),
                             bg='white', fg='#2c3e50',
                             wraplength=800, justify='left')
            q_label.pack(pady=10, padx=20, anchor='w')
            
            # 選項
            options = [
                ('A', option_a),
                ('B', option_b),
                ('C', option_c),
                ('D', option_d)
            ]
            
            for option_key, option_text in options:
                color = '#00a381' if option_key == correct_answer else '#e5abbe' if option_key == user_answer else '#2c3e50'
                prefix = "✓ " if option_key == correct_answer else "✗ " if option_key == user_answer else "  "
                
                option_label = tk.Label(wrong_frame, text=f"{prefix}({option_key}) {option_text}",
                                      font=('Microsoft JhengHei', 11),
                                      bg='white', fg=color,
                                      wraplength=750, justify='left')
                option_label.pack(padx=40, pady=2, anchor='w')
                
            # 詳解
            if explanation:
                exp_label = tk.Label(wrong_frame, text=f"詳解: {explanation}",
                                   font=('Microsoft JhengHei', 11),
                                   bg='#ecf0f1', fg='#2c3e50',
                                   wraplength=750, justify='left')
                exp_label.pack(fill='x', padx=20, pady=(10, 15))
                
    def save_wrong_questions(self):
        """保存錯題記錄到JSON檔案"""
        try:
            # 創建錯題記錄資料夾
            os.makedirs("wrong_questions", exist_ok=True)
            
            # 載入現有錯題記錄
            json_file = "wrong_questions/wrong_questions.json"
            existing_wrong = []
            
            if os.path.exists(json_file):
                try:
                    with open(json_file, 'r', encoding='utf-8') as f:
                        existing_wrong = json.load(f)
                except:
                    existing_wrong = []
                    
            # 添加新的錯題記錄
            for wrong_q in self.wrong_questions:
                record = {
                    'timestamp': datetime.now().isoformat(),
                    'question_data': wrong_q['question_data'],
                    'user_answer': wrong_q['user_answer'],
                    'correct_answer': wrong_q['correct_answer']
                }
                existing_wrong.append(record)
                
            # 保存到JSON檔案
            with open(json_file, 'w', encoding='utf-8') as f:
                json.dump(existing_wrong, f, ensure_ascii=False, indent=2)
                
        except Exception as e:
            print(f"保存錯題記錄失敗: {e}")
            
    def load_wrong_questions(self):
        """載入錯題記錄"""
        json_file = "wrong_questions/wrong_questions.json"
        
        if not os.path.exists(json_file):
            return []
            
        try:
            with open(json_file, 'r', encoding='utf-8') as f:
                wrong_records = json.load(f)
                
            # 提取題目資料，去除重複
            questions = []
            seen_questions = set()
            
            for record in wrong_records:
                question_data = record['question_data']
                question_text = question_data[1]  # 題目文字
                
                if question_text not in seen_questions:
                    questions.append(question_data)
                    seen_questions.add(question_text)
                    
            return questions
            
        except Exception as e:
            print(f"載入錯題記錄失敗: {e}")
            return []
            
    def export_wrong_questions(self):
        """匯出錯題到CSV檔案"""
        if not self.wrong_questions:
            messagebox.showinfo("提示", "沒有錯題可匯出")
            return
            
        try:
            # 創建錯題記錄資料夾
            os.makedirs("wrong_questions", exist_ok=True)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            csv_file = f"wrong_questions/wrong_questions_{timestamp}.csv"
            
            with open(csv_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                
                # 寫入標題行
                writer.writerow(["正確答案", "題目", "選項A", "選項B", "選項C", "選項D", "詳解", "您的答案"])
                
                # 寫入錯題資料
                for wrong_q in self.wrong_questions:
                    question_data = wrong_q['question_data']
                    user_answer = wrong_q['user_answer']
                    row = question_data + [user_answer]
                    writer.writerow(row)
                    
            messagebox.showinfo("成功", f"錯題已匯出至: {csv_file}")
            
        except Exception as e:
            messagebox.showerror("錯誤", f"匯出失敗: {str(e)}")
            
    def restart_quiz(self):
        """重新開始測驗"""
        self.clear_quiz_frame()
        
        # 顯示初始界面內容
        welcome_label = tk.Label(self.quiz_frame, 
                               text="歡迎使用選擇題練習系統，請選擇測驗模式：\n➡［循環式隨機100題測驗］：題目出過一次後，下一輪隨機100題不會再出現，直到所有題目都出過後，曾出過的題目才會再次出現。\n➡［隨機100題測驗］：完全的隨機100題，上一輪出現過的題目，下一輪可能再次出現。\n➡［循環式錯題複習］：循環式練習之前答錯的題目。錯題複習過一次後，下一輪不會再出現，直到所有錯題都複習過後，曾複習過的錯題才會再次出現。\n\n系統說明：\n100題隨機出題形式，每頁顯示50題。\n程式將從各題庫平均抽取題目，總共100題。\n作答後可看正解和詳解。\n錯題記錄及錯題複習功能。\n單機程式，無須網路，不用擔心題庫從雲端外流。\n\n注意事項：\n請自備題庫和詳解。\n題庫請使用UTF-8字元編碼。\n➡新增題庫CSV或修改現有題庫CSV內容之後，請按［重新載入題庫］按鈕更新題庫。\n\n請將題庫CSV放在 questions/ 資料夾。\n題庫CSV格式請見 questions/ 資料夾內的範例題庫CSV。為確保題庫正確讀取，題庫CSV第一列內容請勿更動。\n\n錯題紀錄將匯出於 wrong_questions/ 資料夾內。",
                               font=('Microsoft JhengHei', 12),
                               bg='#f0f0f0', fg='#2c3e50', justify='left' )
        welcome_label.pack(expand=True, pady=50)
        
    def run(self):
        """運行程式"""
        # 顯示歡迎訊息
        self.restart_quiz()
        self.root.mainloop()

if __name__ == "__main__":
    app = QuizSystem()
    app.run()