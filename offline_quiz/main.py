import csv
import json
import os
import random
from datetime import datetime

QUESTIONS_DIR = 'questions'
WRONG_LOG_JSON = 'results/wrong_log.json'
WRONG_LOG_CSV = 'results/wrong_log.csv'

def load_questions_per_file(directory=QUESTIONS_DIR):
    file_questions = []
    for file in sorted(os.listdir(directory)):
        if file.endswith(".csv"):
            with open(os.path.join(directory, file), newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                questions = []
                for i, row in enumerate(reader):
                    questions.append({
                        'id': f"{file}_{i+1}",
                        'question': row['題目'],
                        'A': row['選項1'],
                        'B': row['選項2'],
                        'C': row['選項3'],
                        'D': row['選項4'],
                        'answer': row['答案'],
                        'explanation': row.get('詳解', '')
                    })
                file_questions.append(questions)
    return file_questions

def sample_balanced_questions(file_questions, total_questions=100):
    num_files = len(file_questions)
    base = total_questions // num_files
    extra = total_questions % num_files

    sampled = []
    for i, qlist in enumerate(file_questions):
        count = base + (1 if i < extra else 0)
        if len(qlist) < count:
            raise ValueError(f"Not enough questions in file {i+1} to sample {count}")
        sampled.extend(random.sample(qlist, count))
    random.shuffle(sampled)
    return sampled

def ask_questions(questions):
    score = 0
    wrong = []
    print("\n=== 測驗開始 ===\n")
    for idx, q in enumerate(questions, start=1):
        print(f"{idx}. {q['question']}")
        print(f"   (A) {q['A']}")
        print(f"   (B) {q['B']}")
        print(f"   (C) {q['C']}")
        print(f"   (D) {q['D']}")
        while True:
            ans = input("請作答 (A/B/C/D): ").strip().upper()
            if ans in ['A', 'B', 'C', 'D']:
                break
            else:
              print("不接受A、B、C、D以外的答案，請重新作答。")
        if ans == q['answer'].strip().upper():
            score += 1
        else:
            q['your_answer'] = ans
            wrong.append(q)
        print()
    return score, wrong

def save_wrong_log(wrong_list, user='anonymous'):
    os.makedirs('results', exist_ok=True)
    try:
        with open(WRONG_LOG_JSON, 'r', encoding='utf-8') as f:
            logs = json.load(f)
    except FileNotFoundError:
        logs = []

    for w in wrong_list:
        logs.append({
            'timestamp': datetime.now().isoformat(),
            'user': user,
            'id': w['id'],
            'question': w['question'],
            'your_answer': w['your_answer'],
            'correct_answer': w['answer'],
            'explanation': w.get('explanation', '')
        })

    with open(WRONG_LOG_JSON, 'w', encoding='utf-8') as f:
        json.dump(logs, f, indent=2, ensure_ascii=False)

def export_wrong_log_csv():
    try:
        with open(WRONG_LOG_JSON, 'r', encoding='utf-8') as f:
            logs = json.load(f)
    except FileNotFoundError:
        print("No wrong log found.")
        return

    with open(WRONG_LOG_CSV, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=["timestamp", "user", "id", "question", "your_answer", "correct_answer", "explanation"])
        writer.writeheader()
        writer.writerows(logs)
    print(f"Wrong log exported to {WRONG_LOG_CSV}")

def review_wrong_questions():
    if not os.path.exists(WRONG_LOG_JSON):
        print("No wrong questions to review.")
        return
    with open(WRONG_LOG_JSON, 'r', encoding='utf-8') as f:
        logs = json.load(f)
    unique_questions = {}
    for item in logs:
        qid = item['id']
        if qid not in unique_questions:
            unique_questions[qid] = {
                'id': qid,
                'question': item['question'],
                'A': 'A', 'B': 'B', 'C': 'C', 'D': 'D',
                'answer': item['correct_answer'],
                'explanation': item.get('explanation', '')
            }
    print(f"\n=== Reviewing {len(unique_questions)} unique wrong questions ===\n")
    ask_questions(list(unique_questions.values()))

def main():
    file_questions = load_questions_per_file()
    user = input("請輸入英文名字: ").strip() or "anonymous"
    mode = input("請選擇測驗模式: (1) 隨機 100 題  (2) 錯題複習: ").strip()

    if mode == "2":
        review_wrong_questions()
        return

    questions = sample_balanced_questions(file_questions, 100)
    score, wrong = ask_questions(questions)
    print(f"Your score: {score}/100")
    print(f"You got {len(wrong)} wrong.")
    for w in wrong:
        print(f"Q: {w['question']}")
        print(f"  Your answer: {w['your_answer']} | Correct: {w['answer']}")
        if w.get('explanation'):
            print(f"  Explanation: {w['explanation']}")
    save_wrong_log(wrong, user=user)

    export = input("請問錯題是否匯出成CSV? (y/n): ").strip().lower()
    if export == 'y':
        export_wrong_log_csv()

if __name__ == "__main__":
    main()
