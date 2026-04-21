"""
patch_csv_only.py
Xóa hoàn toàn mbox khỏi notebook, chỉ dùng CSV (mail1.csv + messages.csv)
"""
import sys
sys.stdout.reconfigure(encoding="utf-8")

import nbformat, os, json

NB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "spam_classifier.ipynb")

with open(NB_PATH, "r", encoding="utf-8") as f:
    nb = nbformat.read(f, as_version=4)

print(f"Loaded: {len(nb.cells)} cells")

def find_cell(keyword):
    for i, cell in enumerate(nb.cells):
        if keyword in cell.source:
            return i
    return -1

# ─────────────────────────────────────────────────────────────
# CELL 1 — IMPORTS: xóa mailbox + email.header
# ─────────────────────────────────────────────────────────────
idx = find_cell("import mailbox")
if idx >= 0:
    nb.cells[idx].source = """\
# --- Import các thư viện cần dùng cho dự án ---

import sys
import os
import re
import pickle
import warnings

import numpy as np           # tính toán ma trận
import pandas as pd          # đọc và xử lý dữ liệu bảng
import matplotlib.pyplot as plt  # vẽ biểu đồ
import matplotlib
import seaborn as sns        # vẽ biểu đồ đẹp hơn

from tqdm.notebook import tqdm       # thanh tiến trình trong Jupyter
from collections import Counter
from IPython.display import display, HTML

# scikit-learn: thư viện machine learning chính
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.naive_bayes import MultinomialNB
from sklearn.model_selection import train_test_split
from sklearn.metrics import (
    classification_report, confusion_matrix,
    accuracy_score, precision_score, recall_score, f1_score
)

# xử lý mất cân bằng nhãn (spam ít hơn ham)
try:
    from imblearn.over_sampling import RandomOverSampler
    IMBLEARN_OK = True
except ImportError:
    IMBLEARN_OK = False
    print("chua cai imbalanced-learn -> chay: pip install imbalanced-learn")

# tách từ tiếng Việt
try:
    from underthesea import word_tokenize as vn_tokenize
    UNDERTHESEA_OK = True
except ImportError:
    UNDERTHESEA_OK = False
    print("chua cai underthesea -> chay: pip install underthesea")

# xuất báo cáo Excel
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from tabulate import tabulate

warnings.filterwarnings('ignore')

# cấu hình font cho matplotlib
matplotlib.rcParams['font.family'] = 'DejaVu Sans'
matplotlib.rcParams['figure.dpi']  = 100
plt.style.use('seaborn-v0_8-whitegrid')

print("Import xong!")
print(f"Python : {sys.version.split()[0]}")
print(f"Pandas : {pd.__version__}")
print(f"Sklearn: {__import__('sklearn').__version__}")
"""
    print(f"  [OK] cell {idx}: imports — da xoa mailbox + email.header")

# ─────────────────────────────────────────────────────────────
# CELL 2 — CONFIG: xóa MBOX_PATH + GMAIL_SPAM_LABELS
# ─────────────────────────────────────────────────────────────
idx = find_cell("MBOX_PATH")
if idx >= 0:
    nb.cells[idx].source = """\
# --- Cấu hình dự án ---
# Nếu bạn đặt file ở chỗ khác thì chỉ cần sửa phần này

# thư mục chứa notebook này
BASE_DIR = os.getcwd()

# file CSV dữ liệu gốc
# ⚠️ Đặt mail1.csv và messages.csv vào cùng thư mục với notebook này
MAIL1_CSV    = os.path.join(BASE_DIR, "mail1.csv")     # Gmail 1 (6,854 email)
MESSAGES_CSV = os.path.join(BASE_DIR, "messages.csv")  # Gmail 2 (32,152 email)

# tự động tạo thư mục output nếu chưa có
OUTPUT_DIR  = os.path.join(BASE_DIR, "output")
REPORTS_DIR = os.path.join(OUTPUT_DIR, "reports")
MODELS_DIR  = os.path.join(OUTPUT_DIR, "models")
for d in [OUTPUT_DIR, REPORTS_DIR, MODELS_DIR]:
    os.makedirs(d, exist_ok=True)

# đường dẫn các file kết quả
LABELED_CSV   = os.path.join(OUTPUT_DIR, "labeled_dataset.csv")
PROCESSED_CSV = os.path.join(OUTPUT_DIR, "processed_dataset.csv")
MODEL_PKL     = os.path.join(MODELS_DIR, "naive_bayes_model.pkl")
TFIDF_PKL     = os.path.join(MODELS_DIR, "tfidf_vectorizer.pkl")
REPORT_XLSX   = os.path.join(REPORTS_DIR, "spam_report.xlsx")
REPORT_TXT    = os.path.join(REPORTS_DIR, "spam_summary.txt")

# nhãn
SPAM, HAM = "spam", "ham"

# tham số chia dữ liệu
TEST_SIZE    = 0.20  # 20% test, 80% train
RANDOM_STATE = 42    # seed cố định để kết quả ổn định

# tham số TF-IDF
MAX_FEATURES = 50000   # tối đa 50,000 từ
NGRAM_RANGE  = (1, 2)  # unigram + bigram

# danh sách người gửi/từ khóa để gán nhãn tự động (heuristics)
SPAM_SENDERS = [
    "ecomm.lenovo.com", "recommendations@ted.com", "recommends@ted.com",
    "discover.pinterest.com", "inspire.pinterest.com", "ideas.pinterest.com",
    "explore.pinterest.com", "pinterest.com", "no-reply@grab.com",
    "facebookmail.com", "news@insideapple.apple.com", "noreply@autocode.com",
]
HAM_SENDERS = [
    "security@facebookmail.com", "accounts.google.com",
    "forms-receipts-noreply@google.com", "appstore@insideapple.apple.com",
    "no_reply@email.apple.com",
]
SPAM_KEYWORDS = [
    "ưu đãi", "khuyến mãi", "giảm giá", "voucher", "deal", "offer",
    "sale", "promo", "discount", "unsubscribe", "miễn phí", "free",
    "recommendations", "recommended for you",
]
PHISHING_KW = [
    "tài khoản của bạn bị", "your account has been", "password reset",
    "verify your account", "xác minh tài khoản", "trúng thưởng", "bạn đã thắng",
]

print("Cấu hình xong!")
print(f"Thư mục dự án : {BASE_DIR}")
print(f"mail1.csv     : {'OK' if os.path.exists(MAIL1_CSV) else 'KHONG THAY - dat file vao thu muc du an'}")
print(f"messages.csv  : {'OK' if os.path.exists(MESSAGES_CSV) else 'KHONG THAY - dat file vao thu muc du an'}")
"""
    print(f"  [OK] cell {idx}: config — da xoa MBOX_PATH + GMAIL_SPAM_LABELS")

# ─────────────────────────────────────────────────────────────
# CELL 3 — LABEL FUNCTIONS: xóa decode_header, extract_body, label_by_gmail
# ─────────────────────────────────────────────────────────────
idx = find_cell("def decode_header")
if idx >= 0:
    nb.cells[idx].source = """\
# --- Hàm gán nhãn bằng quy tắc (heuristics) ---
# Dùng cho cả Gmail1 và Gmail2 vì chỉ có file CSV, không có nhãn thật

# gán nhãn dựa trên người gửi và tiêu đề email
# ưu tiên: ham sender > spam sender > spam keyword > mặc định ham
def label_by_heuristics(sender, subject):
    s = str(sender).lower()
    q = str(subject).lower()

    if any(h in s for h in HAM_SENDERS):    # email quan trọng -> ham chắc chắn
        return HAM
    if any(sp in s for sp in SPAM_SENDERS): # người gửi marketing -> spam
        return SPAM
    if any(kw in q for kw in SPAM_KEYWORDS): # tiêu đề có từ spam
        return SPAM
    return HAM  # không rõ -> mặc định là ham


print("Ham gan nhan san sang!")
"""
    print(f"  [OK] cell {idx}: label functions — giu lai label_by_heuristics")

# ─────────────────────────────────────────────────────────────
# CELL 4 — DATA LOADING: thay load_gmail2() mbox bằng CSV
# ─────────────────────────────────────────────────────────────
idx = find_cell("def load_gmail2")
if idx >= 0:
    nb.cells[idx].source = """\
# --- Đọc dữ liệu từ 2 file CSV và gán nhãn ---

def load_gmail2_csv():
    \"\"\"Đọc Gmail2 từ messages.csv, gán nhãn bằng heuristics.\"\"\"
    records = []
    df2 = pd.read_csv(MESSAGES_CSV, encoding="utf-8-sig", header=0, low_memory=False)
    df2.columns = ["subject", "sender", "to", "date", "starred", "size", "body"]
    df2.fillna("", inplace=True)
    for _, row in tqdm(df2.iterrows(), total=len(df2), desc="Gmail 2 (messages.csv)", unit="email"):
        label = label_by_heuristics(str(row["sender"]), str(row["subject"]))
        records.append({
            "source":       "gmail2",
            "subject":      str(row["subject"]),
            "sender":       str(row["sender"]),
            "date":         str(row["date"]),
            "body":         str(row["body"])[:8000],
            "label":        label,
            "label_source": "heuristic",
        })
    s = sum(1 for r in records if r["label"] == SPAM)
    print(f"  Gmail 2: {len(records):,} email -> {s:,} spam | {len(records)-s:,} ham")
    return records


def load_gmail1():
    \"\"\"Đọc Gmail1 từ mail1.csv, gán nhãn bằng heuristics.\"\"\"
    records = []
    df1 = pd.read_csv(MAIL1_CSV, encoding="utf-8-sig", header=0, low_memory=False)
    df1.columns = ["subject", "sender", "to", "date", "starred", "size", "body"]
    df1.fillna("", inplace=True)
    for _, row in tqdm(df1.iterrows(), total=len(df1), desc="Gmail 1 (mail1.csv)", unit="email"):
        label = label_by_heuristics(str(row["sender"]), str(row["subject"]))
        records.append({
            "source":       "gmail1",
            "subject":      str(row["subject"]),
            "sender":       str(row["sender"]),
            "date":         str(row["date"]),
            "body":         str(row["body"])[:8000],
            "label":        label,
            "label_source": "heuristic",
        })
    s = sum(1 for r in records if r["label"] == SPAM)
    print(f"  Gmail 1: {len(records):,} email -> {s:,} spam | {len(records)-s:,} ham")
    return records


# --- Load dữ liệu (có cache) ---
# Lần đầu: đọc 2 file CSV + gán nhãn -> lưu cache (~1-2 phút)
# Lần sau: tải thẳng từ cache (~5 giây)
if os.path.exists(LABELED_CSV):
    print(f"Tim thay cache: {LABELED_CSV}")
    df = pd.read_csv(LABELED_CSV, encoding="utf-8-sig", low_memory=False)
    df.fillna("", inplace=True)
    print(f"  Tai cache: {len(df):,} email")
else:
    print("Chua co cache, doc tu CSV...")
    recs = load_gmail2_csv() + load_gmail1()
    df   = pd.DataFrame(recs)

    # sửa nhãn phishing bị gán nhầm ham
    fixed = 0
    for idx_r, row in df.iterrows():
        if row["label"] == HAM:
            combo = (str(row["subject"]) + " " + str(row["body"])[:500]).lower()
            if any(kw in combo for kw in PHISHING_KW):
                df.at[idx_r, "label"] = SPAM
                fixed += 1
    if fixed:
        print(f"  Da sua {fixed} email phishing -> spam")

    df.to_csv(LABELED_CSV, index=False, encoding="utf-8-sig")
    print(f"  Da luu cache: {LABELED_CSV}")

df = df.reset_index(drop=True)

spam_n = (df["label"] == SPAM).sum()
ham_n  = (df["label"] == HAM).sum()
print(f"\\nTONG KET:")
print(f"  Tong : {len(df):,} email")
print(f"  Spam : {spam_n:,} ({spam_n/len(df)*100:.1f}%)")
print(f"  Ham  : {ham_n:,} ({ham_n/len(df)*100:.1f}%)")
"""
    print(f"  [OK] cell {idx}: data loading — thay mbox bang messages.csv")

# ─────────────────────────────────────────────────────────────
# CELL Phase 1 markdown — cập nhật mô tả
# ─────────────────────────────────────────────────────────────
idx = find_cell("Phase 1 — Thu thập & Gán nhãn")
if idx >= 0:
    nb.cells[idx].source = "## 📂 Phase 1 — Thu thập & Gán nhãn Dữ liệu\n\nDữ liệu từ **2 file CSV** (mail1.csv + messages.csv), gán nhãn tự động bằng heuristics."
    print(f"  [OK] cell {idx}: Phase 1 markdown")

# ─────────────────────────────────────────────────────────────
# XÓA CACHE CŨ để force re-label từ CSV
# ─────────────────────────────────────────────────────────────
cached = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                      "output", "labeled_dataset.csv")
if os.path.exists(cached):
    os.remove(cached)
    print(f"\n[OK] Da xoa cache cu: {cached}")
    print("     (se tu tao lai tu CSV khi chay notebook)")

# ─────────────────────────────────────────────────────────────
# LƯU
# ─────────────────────────────────────────────────────────────
with open(NB_PATH, "w", encoding="utf-8") as f:
    nbformat.write(nb, f)

size_kb = os.path.getsize(NB_PATH) // 1024
print(f"\nDone! {NB_PATH}")
print(f"Kich thuoc: {size_kb} KB | Cells: {len(nb.cells)}")
