"""
patch_notebook.py
- Sửa các cell hiện có cho giống code sinh viên hơn
- Thêm Phase 9: Tự kiểm tra (gõ tay + load CSV)
Chạy: python patch_notebook.py
"""
import sys
sys.stdout.reconfigure(encoding="utf-8")

import nbformat
import os

NB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "spam_classifier.ipynb")

# ─────────────────────────────────────────────
# LOAD NOTEBOOK
# ─────────────────────────────────────────────
with open(NB_PATH, "r", encoding="utf-8") as f:
    nb = nbformat.read(f, as_version=4)

print(f"Đã load notebook: {len(nb.cells)} cells")

# ─────────────────────────────────────────────
# HÀM HELPER
# ─────────────────────────────────────────────
def find_cell(keyword):
    """Tìm index cell đầu tiên chứa keyword."""
    for i, cell in enumerate(nb.cells):
        if keyword in cell.source:
            return i
    return -1

def make_code(src):
    return nbformat.v4.new_code_cell(src)

def make_md(src):
    return nbformat.v4.new_markdown_cell(src)

# ─────────────────────────────────────────────
# SỬA CELL 1 — IMPORT (sinh viên hơn)
# ─────────────────────────────────────────────
idx = find_cell("IMPORT TẤT CẢ THƯ VIỆN")
if idx >= 0:
    nb.cells[idx].source = """\
# --- Import các thư viện cần dùng cho dự án ---

import sys
import os
import re
import pickle
import warnings
import mailbox
import email.header

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

# cấu hình font cho matplotlib (hỗ trợ tiếng Việt trên Windows)
matplotlib.rcParams['font.family'] = 'DejaVu Sans'
matplotlib.rcParams['figure.dpi']  = 100
plt.style.use('seaborn-v0_8-whitegrid')

print("Import xong!")
print(f"Python : {sys.version.split()[0]}")
print(f"Pandas : {pd.__version__}")
print(f"Sklearn: {__import__('sklearn').__version__}")
"""
    print(f"  [OK] Sửa cell {idx}: imports")

# ─────────────────────────────────────────────
# SỬA CELL 2 — CẤU HÌNH (sinh viên hơn)
# ─────────────────────────────────────────────
idx = find_cell("CẤU HÌNH — chỉnh sửa đường dẫn")
if idx >= 0:
    nb.cells[idx].source = """\
# --- Cấu hình dự án ---
# Nếu bạn đặt file ở chỗ khác thì chỉ cần sửa phần này, không cần đụng code bên dưới

# thư mục chứa notebook này
BASE_DIR = os.getcwd()

# ⚠️ SỬA đường dẫn file .mbox nếu bạn lưu ở nơi khác
MBOX_PATH = r"D:\\UTH\\AI\\Spam\\data\\Mail\\Tất cả thư bao gồm spam và thư rác.mbox"

# file CSV dữ liệu gốc
MAIL1_CSV    = os.path.join(BASE_DIR, "mail1.csv")     # Gmail 1
MESSAGES_CSV = os.path.join(BASE_DIR, "messages.csv")  # Gmail 2 (dùng khi không có mbox)

# tự động tạo thư mục output nếu chưa có
OUTPUT_DIR  = os.path.join(BASE_DIR, "output")
REPORTS_DIR = os.path.join(OUTPUT_DIR, "reports")
MODELS_DIR  = os.path.join(OUTPUT_DIR, "models")
for d in [OUTPUT_DIR, REPORTS_DIR, MODELS_DIR]:
    os.makedirs(d, exist_ok=True)

# đường dẫn các file kết quả
LABELED_CSV   = os.path.join(OUTPUT_DIR, "labeled_dataset.csv")    # sau bước gán nhãn
PROCESSED_CSV = os.path.join(OUTPUT_DIR, "processed_dataset.csv")  # sau bước tiền xử lý
MODEL_PKL     = os.path.join(MODELS_DIR, "naive_bayes_model.pkl")  # model đã train
TFIDF_PKL     = os.path.join(MODELS_DIR, "tfidf_vectorizer.pkl")   # vectorizer
REPORT_XLSX   = os.path.join(REPORTS_DIR, "spam_report.xlsx")
REPORT_TXT    = os.path.join(REPORTS_DIR, "spam_summary.txt")

# nhãn — dùng biến cho chắc ăn, tránh typo
SPAM, HAM = "spam", "ham"

# tham số chia dữ liệu
TEST_SIZE    = 0.20  # 20% để test, 80% để train
RANDOM_STATE = 42    # để kết quả giống nhau mỗi lần chạy

# tham số TF-IDF
MAX_FEATURES = 50000   # tối đa 50,000 từ
NGRAM_RANGE  = (1, 2)  # dùng cả 1 từ và 2 từ ghép

# nhãn Gmail2 nào được coi là SPAM
GMAIL_SPAM_LABELS = ["Spam", "Danh mục Khuyến mại"]

# danh sách người gửi/từ khóa dùng để gán nhãn Gmail1 (không có mbox)
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
print(f"File mbox     : {'OK - tìm thấy' if os.path.exists(MBOX_PATH) else 'KHÔNG thấy (sẽ dùng CSV)'}")
"""
    print(f"  [OK] Sửa cell {idx}: config")

# ─────────────────────────────────────────────
# SỬA CELL — HÀM GÁN NHÃN (sinh viên hơn)
# ─────────────────────────────────────────────
idx = find_cell("CÁC HÀM GÁN NHÃN DỮ LIỆU")
if idx >= 0:
    nb.cells[idx].source = """\
# --- Các hàm để đọc và gán nhãn email ---

# hàm giải mã tiêu đề email bị mã hóa
# vd: "=?UTF-8?B?S2h1eeG6v24gbcOjaQ==?=" -> "Khuyến mãi"
def decode_header(raw):
    if not raw:
        return ""
    try:
        parts = email.header.decode_header(raw)
        result = ""
        for part, enc in parts:
            if isinstance(part, bytes):
                result += part.decode(enc or "utf-8", errors="replace")
            else:
                result += str(part)
        return result.strip()
    except Exception:
        return str(raw).strip()


# hàm lấy nội dung text từ email (bỏ qua file đính kèm)
def extract_body(msg):
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            # bỏ qua file đính kèm
            if "attachment" in str(part.get("Content-Disposition", "")):
                continue
            if ctype == "text/plain":
                raw = part.get_payload(decode=True)
                if raw:
                    charset = part.get_content_charset() or "utf-8"
                    body = raw.decode(charset, errors="replace")
                    break
            elif ctype == "text/html" and not body:
                raw = part.get_payload(decode=True)
                if raw:
                    charset = part.get_content_charset() or "utf-8"
                    body = raw.decode(charset, errors="replace")
    else:
        raw = msg.get_payload(decode=True)
        if raw:
            charset = msg.get_content_charset() or "utf-8"
            body = raw.decode(charset, errors="replace")
    return body[:8000]  # giới hạn 8000 ký tự, đủ để phân tích rồi


# gán nhãn dựa trên nhãn thật của Gmail (X-Gmail-Labels header)
# Gmail đã gán sẵn cho mình rồi, chỉ cần đọc ra
def label_by_gmail(label_str):
    if not label_str:
        return HAM
    labels = [l.strip() for l in label_str.split(",")]
    return SPAM if any(sl in labels for sl in GMAIL_SPAM_LABELS) else HAM


# gán nhãn bằng quy tắc thủ công (dùng cho Gmail1 không có mbox)
# ưu tiên: ham sender > spam sender > spam keyword > mặc định ham
def label_by_heuristics(sender, subject):
    s = sender.lower()
    q = subject.lower()
    if any(h in s for h in HAM_SENDERS):   # email quan trọng -> ham chắc chắn
        return HAM
    if any(sp in s for sp in SPAM_SENDERS): # người gửi marketing -> spam
        return SPAM
    if any(kw in q for kw in SPAM_KEYWORDS): # tiêu đề có từ spam -> spam
        return SPAM
    return HAM  # không rõ -> mặc định là ham


print("Hàm gán nhãn sẵn sàng!")
"""
    print(f"  [OK] Sửa cell {idx}: label functions")

# ─────────────────────────────────────────────
# SỬA CELL — HÀM TIỀN XỬ LÝ (sinh viên hơn)
# ─────────────────────────────────────────────
idx = find_cell("CÁC HÀM TIỀN XỬ LÝ VĂN BẢN")
if idx >= 0:
    nb.cells[idx].source = """\
# --- Các hàm tiền xử lý văn bản ---
from sklearn.feature_extraction.text import ENGLISH_STOP_WORDS

# danh sách từ dừng tiếng Việt (từ không có giá trị phân loại)
VN_STOP = {
    "và","của","cho","là","có","được","trong","không","với","các","một",
    "để","này","từ","tôi","bạn","đã","đang","sẽ","khi","như","hay","hoặc",
    "nhưng","mà","thì","vì","bởi","nên","ra","đi","lên","xuống","theo",
    "về","qua","trên","dưới","sau","trước","đây","đó","ở","tại","vào",
    "rất","lắm","quá","khá","cũng","vậy","thế","đây","đó","còn","vẫn",
}
EN_STOP = set(ENGLISH_STOP_WORDS)  # từ dừng tiếng Anh từ sklearn


# bước 1: làm sạch text — xóa HTML, link, ký tự lạ
def clean_text(text):
    if not text or not isinstance(text, str):
        return ""
    text = re.sub(r'<[^>]+>',      ' ', text)  # xóa thẻ HTML như <div>, <p>
    text = re.sub(r'&[a-z]+;',     ' ', text)  # xóa &amp; &nbsp; ...
    text = re.sub(r'https?://\S+', ' ', text)  # xóa link http://...
    text = re.sub(r'www\.\S+',     ' ', text)  # xóa www....
    text = re.sub(r'\S+@\S+',      ' ', text)  # xóa địa chỉ email
    text = re.sub(r'[^\w\s]',      ' ', text)  # xóa ký tự đặc biệt ! @ # ...
    text = re.sub(r'\b\d+\b',      ' ', text)  # xóa số đứng một mình
    text = re.sub(r'\s+',          ' ', text)  # nhiều khoảng trắng -> 1
    return text.lower().strip()


# bước 2: phát hiện ngôn ngữ đơn giản (đếm ký tự đặc trưng tiếng Việt)
def detect_lang(text):
    vn_chars = set('àáâãèéêìíòóôõùúýăđơưạặầẩẫậắẳẵặẹẽếềểễệỉịọộốồổỗợớờởỡụứừửữựỳỵ')
    count = sum(1 for c in text if c in vn_chars)
    return 'vi' if count > len(text) * 0.04 else 'en'  # >4% ký tự VN -> tiếng Việt


# bước 3: tách từ tiếng Việt bằng underthesea
# "khuyến mãi" -> "khuyến_mãi" (giữ nguyên nghĩa cụm từ)
def tokenize_vi(text):
    if UNDERTHESEA_OK:
        try:
            return vn_tokenize(text, format="text")
        except Exception:
            return text
    return text


# bước 4: xóa từ dừng và từ quá ngắn
def remove_stops(text):
    return ' '.join(
        w for w in text.split()
        if len(w) >= 2 and w not in VN_STOP and w not in EN_STOP
    )


# pipeline đầy đủ cho 1 email: ghép subject + body rồi qua 4 bước
def preprocess(subject, body):
    # nhân subject lên 3 lần vì tiêu đề thường quan trọng hơn nội dung
    raw = (str(subject) + ' ') * 3 + ' ' + str(body)
    cleaned = clean_text(raw)
    if detect_lang(cleaned) == 'vi':
        cleaned = tokenize_vi(cleaned)
    return remove_stops(cleaned)


print("Hàm tiền xử lý sẵn sàng!")

# thử với 1 ví dụ cho dễ hiểu
vi_du_subject = "Khuyến mãi đặc biệt: giảm 50% hôm nay!"
vi_du_body    = "Chúc mừng bạn đã trúng thưởng! Click vào link để nhận quà."
vi_du_result  = preprocess(vi_du_subject, vi_du_body)
print(f"\\nVí dụ tiền xử lý:")
print(f"  Input : {vi_du_subject}")
print(f"  Output: {vi_du_result[:120]}")
"""
    print(f"  [OK] Sửa cell {idx}: preprocessing functions")

# ─────────────────────────────────────────────
# SỬA CELL — TRAIN MODEL (sinh viên hơn)
# ─────────────────────────────────────────────
idx = find_cell("HUẤN LUYỆN MULTINOMIAL NAIVE BAYES")
if idx >= 0:
    nb.cells[idx].source = """\
# --- Huấn luyện mô hình Multinomial Naive Bayes ---
# MultinomialNB phù hợp với dữ liệu đếm (tần suất từ)
# alpha=1.0 là Laplace smoothing: tránh xác suất = 0 khi gặp từ mới

print(f"Đang train model...")
print(f"  Số email train: {X_train_res.shape[0]:,}")
print(f"  Số features   : {X_train_res.shape[1]:,}")

model = MultinomialNB(alpha=1.0)
model.fit(X_train_res, y_train_res)

# lưu model vào file để dùng lại sau (không phải train lại)
with open(MODEL_PKL, 'wb') as f:
    pickle.dump(model, f)

# dự đoán luôn trên tập test để xem kết quả sơ bộ
y_pred = model.predict(X_test)

print(f"\\nTrain xong!")
print(f"  Model lưu tại: {MODEL_PKL}")
print(f"  Accuracy sơ bộ: {accuracy_score(y_test, y_pred)*100:.2f}%")
"""
    print(f"  [OK] Sửa cell {idx}: train model")

# ─────────────────────────────────────────────
# THÊM PHASE 9 — TỰ KIỂM TRA
# ─────────────────────────────────────────────
phase9_cells = [

    # --- Markdown header ---
    make_md("""\
---
## 🧪 Phase 9 — Tự Kiểm Tra Email

Phần này **không ảnh hưởng** đến các Phase 1-8.
Dùng để kiểm tra nhanh sau khi đã chạy xong toàn bộ notebook.

| Mục | Mô tả |
|-----|-------|
| **9.1** | Load lại model (nếu mở notebook mới) |
| **9.2** | Gõ tay 1 email để kiểm tra |
| **9.3** | Test nhiều email cùng lúc |
| **9.4** | Load file CSV của bạn để kiểm tra hàng loạt |

> ⚠️ **Yêu cầu:** Phải có file model trong `output/models/` (chạy Phase 1-8 ít nhất 1 lần)
"""),

    # --- 9.1: Load model ---
    make_code("""\
# ============================================================
# 9.1 — LOAD LẠI MODEL
# Chạy cell này nếu bạn mở notebook mới mà chưa chạy Phase 1-8
# ============================================================

import pickle, os

_model_path = os.path.join(os.getcwd(), "output", "models", "naive_bayes_model.pkl")
_tfidf_path = os.path.join(os.getcwd(), "output", "models", "tfidf_vectorizer.pkl")

if os.path.exists(_model_path) and os.path.exists(_tfidf_path):
    with open(_model_path, 'rb') as f:
        model = pickle.load(f)
    with open(_tfidf_path, 'rb') as f:
        tfidf = pickle.load(f)
    print("Load model thanh cong!")
    print(f"  Classes: {list(model.classes_)}")
    print(f"  So features: {tfidf.get_feature_names_out().shape[0]:,}")
else:
    print("Chua co model! Hay chay Phase 1-8 truoc.")
    print(f"  Can file: {_model_path}")
"""),

    # --- 9.2: Test 1 email gõ tay ---
    make_md("### ✏️ 9.2 — Gõ tay 1 email để kiểm tra"),

    make_code("""\
# ============================================================
# 9.2 — TỰ GÕ NỘI DUNG EMAIL ĐỂ KIỂM TRA
#
# 👇 CHỈ CẦN THAY ĐỔI 2 DÒNG NÀY RỒI BẤM RUN CELL
# ============================================================

tieu_de  = "Chúc mừng! Bạn đã trúng thưởng 500k"          # <- đổi tiêu đề ở đây
noi_dung = "Click vào link để nhận thưởng ngay hôm nay!"   # <- đổi nội dung ở đây

# ============================================================
# (không cần sửa từ đây trở xuống)
# ============================================================

# tiền xử lý giống hệt lúc train
import re
van_ban = preprocess(tieu_de, noi_dung)

# chuyển sang vector số
vector = tfidf.transform([van_ban])

# dự đoán
ket_qua  = model.predict(vector)[0]
xac_suat = model.predict_proba(vector)[0]

# lấy xác suất từng nhãn
classes      = list(model.classes_)
spam_pct     = xac_suat[classes.index('spam')] * 100
ham_pct      = xac_suat[classes.index('ham')]  * 100

# in kết quả
print("-" * 55)
print(f"Tieu de  : {tieu_de}")
print(f"Noi dung : {noi_dung[:70]}")
print("-" * 55)

if ket_qua == 'spam':
    print(f"  KET QUA: SPAM")
else:
    print(f"  KET QUA: HAM (email binh thuong)")

print(f"  Xac suat spam : {spam_pct:.1f}%")
print(f"  Xac suat ham  : {ham_pct:.1f}%")
print("-" * 55)
"""),

    # --- 9.3: Test nhiều email ---
    make_md("### 📋 9.3 — Test nhiều email cùng lúc"),

    make_code("""\
# ============================================================
# 9.3 — TEST NHIỀU EMAIL CÙNG LÚC
#
# 👇 THÊM / XÓA / SỬA email trong danh sách bên dưới
# Format: ("tiêu đề", "nội dung")
# ============================================================

danh_sach_test = [
    ("Khuyến mãi 50% hôm nay!", "Giảm giá cực sốc, click ngay để mua hàng"),
    ("Meeting at 3pm tomorrow", "Hi team, please join the meeting on Zoom"),
    ("Trúng thưởng iPhone 15!", "Ban la nguoi may man, nhan thuong ngay!"),
    ("Your order has been shipped", "Your package #12345 is on the way"),
    ("Ưu đãi voucher 200k", "Dùng code SALE200 để được giảm giá ngay"),
    ("Password reset request", "Click here to reset your password"),
    ("Re: báo cáo tuần này", "Mình gửi file báo cáo, anh xem và góp ý nhé"),
    ("FREE GIFT WAITING FOR YOU", "Claim your free gift now before it expires!!!"),
    ("Lịch thi cuối kỳ", "Lịch thi học kỳ 2 đã được cập nhật trên portal"),
    ("Win $1000 cash prize!", "You have been selected. Click to claim your prize"),
]

# ============================================================
print(f"{'STT':<5} {'KET QUA':<12} {'SPAM%':>7}   TIEU DE")
print("-" * 65)

for i, (tieu_de, noi_dung) in enumerate(danh_sach_test, 1):
    van_ban  = preprocess(tieu_de, noi_dung)
    vector   = tfidf.transform([van_ban])
    ket_qua  = model.predict(vector)[0]
    xac_suat = model.predict_proba(vector)[0]
    spam_pct = xac_suat[list(model.classes_).index('spam')] * 100

    nhan = "SPAM" if ket_qua == 'spam' else "ham "
    print(f"[{i:<3}]  {nhan:<12} {spam_pct:>6.1f}%   {tieu_de[:42]}")

print("-" * 65)
n_spam = sum(
    1 for td, nd in danh_sach_test
    if model.predict(tfidf.transform([preprocess(td, nd)]))[0] == 'spam'
)
print(f"Tong: {n_spam} SPAM  |  {len(danh_sach_test) - n_spam} HAM")
"""),

    # --- 9.4: Test từ CSV ---
    make_md("""\
### 📂 9.4 — Load file CSV của bạn để kiểm tra hàng loạt

**Hỗ trợ 3 format CSV:**

| Format | Cột cần có | Ví dụ |
|--------|-----------|-------|
| **Format A** — giống mail1.csv | 7 cột không header (Column1=subject, Column7=body) | File export Gmail dạng CSV |
| **Format B** — CSV đơn giản | Header: `subject`, `body` | File tự tạo |
| **Format C** — chỉ 1 cột text | Header: `text` | Gộp subject+body vào 1 cột |

> 💡 Kết quả sẽ tự động lưu vào `output/reports/test_result.csv`
"""),

    make_code("""\
# ============================================================
# 9.4 — TEST TỪ FILE CSV CỦA BẠN
#
# 👇 CHỈ CẦN THAY ĐỔI DÒNG NÀY — đặt đường dẫn file CSV vào đây
# ============================================================

DUONG_DAN_CSV = r"C:\\Users\\wiisc\\Downloads\\test_emails.csv"

# ============================================================
# (không cần sửa từ đây trở xuống)
# ============================================================

import os
import pandas as pd

# kiểm tra file có tồn tại không
if not os.path.exists(DUONG_DAN_CSV):
    print(f"Khong tim thay file: {DUONG_DAN_CSV}")
    print("  Hay kiem tra lai duong dan")
else:
    # đọc file CSV (thử utf-8 trước, nếu lỗi thử latin-1)
    try:
        df_test = pd.read_csv(DUONG_DAN_CSV, encoding='utf-8-sig', low_memory=False)
    except UnicodeDecodeError:
        df_test = pd.read_csv(DUONG_DAN_CSV, encoding='latin-1', low_memory=False)

    print(f"Doc file thanh cong: {len(df_test)} dong")
    print(f"  Cac cot: {list(df_test.columns)}")
    df_test.fillna("", inplace=True)

    # tự nhận dạng format file
    cols_lower = [str(c).lower().strip() for c in df_test.columns]

    if 'subject' in cols_lower and 'body' in cols_lower:
        # Format B: có cột subject + body
        col_sub  = df_test.columns[cols_lower.index('subject')]
        col_body = df_test.columns[cols_lower.index('body')]
        print(f"  -> Format B: subject='{col_sub}', body='{col_body}'")
        get_sub  = lambda row: str(row[col_sub])
        get_body = lambda row: str(row[col_body])

    elif len(df_test.columns) >= 7:
        # Format A: giống mail1.csv (7 cột không có header rõ ràng)
        df_test.columns = (
            ["subject", "sender", "to", "date", "starred", "size", "body"]
            + [f"col{i}" for i in range(len(df_test.columns) - 7)]
        )
        print(f"  -> Format A: giong mail1.csv (7+ cot)")
        get_sub  = lambda row: str(row['subject'])
        get_body = lambda row: str(row['body'])

    elif 'text' in cols_lower:
        # Format C: chỉ 1 cột text
        col_text = df_test.columns[cols_lower.index('text')]
        print(f"  -> Format C: 1 cot text='{col_text}'")
        get_sub  = lambda row: ""
        get_body = lambda row: str(row[col_text])

    else:
        # không nhận dạng được -> dùng cột 0 và 1
        print(f"  -> Khong nhan dang duoc format, dung cot 0=subject, cot 1=body")
        get_sub  = lambda row: str(row.iloc[0])
        get_body = lambda row: str(row.iloc[1]) if len(row) > 1 else ""

    # chạy dự đoán từng dòng
    print(f"\\nDang phan loai {len(df_test)} email...")
    ds_ket_qua   = []
    ds_spam_pct  = []

    for _, row in df_test.iterrows():
        van_ban  = preprocess(get_sub(row), get_body(row))
        vector   = tfidf.transform([van_ban])
        ket_qua  = model.predict(vector)[0]
        xac_suat = model.predict_proba(vector)[0]
        spam_pct = xac_suat[list(model.classes_).index('spam')] * 100
        ds_ket_qua.append(ket_qua)
        ds_spam_pct.append(round(spam_pct, 2))

    df_test['ket_qua']  = ds_ket_qua
    df_test['spam_pct'] = ds_spam_pct

    # tóm tắt kết quả
    n_spam = (df_test['ket_qua'] == 'spam').sum()
    n_ham  = len(df_test) - n_spam
    print(f"\\nKET QUA PHAN LOAI:")
    print(f"  SPAM : {n_spam} email ({n_spam / len(df_test) * 100:.1f}%)")
    print(f"  HAM  : {n_ham}  email ({n_ham  / len(df_test) * 100:.1f}%)")

    # in mẫu 10 dòng đầu
    print(f"\\nMau 10 dong dau:")
    print(f"{'STT':<5} {'KET QUA':<10} {'SPAM%':>7}   TIEU DE / NOI DUNG")
    print("-" * 65)
    for i, row in df_test.head(10).iterrows():
        nhan = "SPAM" if row['ket_qua'] == 'spam' else "ham "
        tieu = str(get_sub(row))[:42] if str(get_sub(row)) else str(get_body(row))[:42]
        print(f"[{i+1:<3}]  {nhan:<10} {row['spam_pct']:>6.1f}%   {tieu}")
    print("-" * 65)

    # lưu kết quả đầy đủ ra CSV
    output_csv = os.path.join(os.getcwd(), "output", "reports", "test_result.csv")
    df_test.to_csv(output_csv, index=False, encoding='utf-8-sig')
    print(f"\\nDa luu ket qua day du: {output_csv}")
    print(f"  Mo file bang Excel de xem toan bo {len(df_test)} dong")
"""),

    # --- Hướng dẫn tạo CSV test ---
    make_md("""\
### 📝 Hướng dẫn tạo file CSV để test

Nếu bạn chưa có file CSV, có thể tự tạo nhanh bằng Excel hoặc Notepad:

**Format B (dễ nhất)** — lưu file `.csv` với nội dung:
```
subject,body
"Khuyến mãi đặc biệt hôm nay","Click vào link để nhận ưu đãi"
"Meeting reminder tomorrow","Hi, don't forget our 3pm meeting"
"Trúng thưởng 1 triệu","Bạn là người may mắn, nhận ngay!"
```

**Cách tạo trong Excel:**
1. Mở Excel → tạo 2 cột: `subject` và `body`
2. Điền email cần test vào từng dòng
3. File → Save As → chọn **CSV UTF-8** → lưu
4. Copy đường dẫn file → paste vào `DUONG_DAN_CSV` ở cell 9.4
"""),
]

# thêm tất cả cells Phase 9 vào cuối notebook
nb.cells.extend(phase9_cells)
print(f"\nĐã thêm {len(phase9_cells)} cells Phase 9")

# ─────────────────────────────────────────────
# LƯU NOTEBOOK
# ─────────────────────────────────────────────
with open(NB_PATH, "w", encoding="utf-8") as f:
    nbformat.write(nb, f)

size_kb = os.path.getsize(NB_PATH) // 1024
print(f"\nDone! Notebook da luu: {NB_PATH}")
print(f"Kich thuoc: {size_kb} KB | Tong cells: {len(nb.cells)}")
print(f"\nMo lai spam_classifier.ipynb trong VSCode de xem ket qua!")
