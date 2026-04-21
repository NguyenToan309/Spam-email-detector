# =============================================================================
# config.py — Tập tin cấu hình toàn bộ dự án Spam Classifier
# Mọi đường dẫn, tham số, hằng số đều đặt ở đây
# → Không được hardcode ở các file khác, chỉ import từ đây
# =============================================================================

import os

# -----------------------------------------------------------------------------
# 1. ĐƯỜNG DẪN GỐC DỰ ÁN
#    os.path.dirname(__file__)  →  thư mục chứa file config.py này
#    Dùng cách này để code chạy đúng dù bạn đặt thư mục ở đâu trên máy
# -----------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# -----------------------------------------------------------------------------
# 2. ĐƯỜNG DẪN FILE DỮ LIỆU ĐẦU VÀO
#    mail1.csv     → email từ Gmail 1 (wiisch3009@gmail.com)
#    messages.csv  → email từ Gmail 2 (nguyenkhanhtoan.309@gmail.com)
#    MBOX_PATH     → file .mbox gốc từ Google Takeout — dùng để lấy nhãn thật
# -----------------------------------------------------------------------------
MAIL1_CSV    = os.path.join(BASE_DIR, "mail1.csv")
MESSAGES_CSV = os.path.join(BASE_DIR, "messages.csv")
MBOX_PATH    = r"D:\UTH\AI\Spam\data\Mail\Tất cả thư bao gồm spam và thư rác.mbox"

# -----------------------------------------------------------------------------
# 3. ĐƯỜNG DẪN THƯ MỤC ĐẦU RA
#    OUTPUT_DIR   → thư mục lưu tất cả kết quả
#    REPORTS_DIR  → thư mục lưu báo cáo
# -----------------------------------------------------------------------------
OUTPUT_DIR  = os.path.join(BASE_DIR, "output")
REPORTS_DIR = os.path.join(OUTPUT_DIR, "reports")

# -----------------------------------------------------------------------------
# 4. ĐƯỜNG DẪN FILE KẾT QUẢ CÁC BƯỚC
#    Mỗi bước lưu ra 1 file riêng → dễ debug nếu một bước bị lỗi
# -----------------------------------------------------------------------------
LABELED_CSV    = os.path.join(OUTPUT_DIR, "labeled_dataset.csv")    # Sau bước gán nhãn
PROCESSED_CSV  = os.path.join(OUTPUT_DIR, "processed_dataset.csv")  # Sau bước tiền xử lý
MODEL_PKL      = os.path.join(OUTPUT_DIR, "naive_bayes_model.pkl")   # File model đã train
TFIDF_PKL      = os.path.join(OUTPUT_DIR, "tfidf_vectorizer.pkl")    # File TF-IDF vectorizer
REPORT_XLSX    = os.path.join(REPORTS_DIR, "spam_report.xlsx")       # Báo cáo Excel
REPORT_TXT     = os.path.join(REPORTS_DIR, "spam_summary.txt")       # Báo cáo tóm tắt

# -----------------------------------------------------------------------------
# 5. NHÃN (LABELS) — giá trị dùng trong cột "label" của CSV
#    Dùng hằng số thay vì string "spam"/"ham" để tránh typo
# -----------------------------------------------------------------------------
LABEL_SPAM = "spam"
LABEL_HAM  = "ham"

# -----------------------------------------------------------------------------
# 6. CÁC CỘT (COLUMNS) TRONG CSV DỮ LIỆU GỐC
#    Column1 = Tiêu đề (subject)
#    Column2 = Người gửi (from)
#    Column3 = Người nhận (to)
#    Column4 = Ngày gửi (date)
#    Column5 = Đã đánh dấu sao (*) hay không
#    Column6 = Kích thước email (bytes)
#    Column7 = Nội dung đầy đủ (body)
# -----------------------------------------------------------------------------
COL_SUBJECT = "Column1"
COL_FROM    = "Column2"
COL_TO      = "Column3"
COL_DATE    = "Column4"
COL_STARRED = "Column5"
COL_SIZE    = "Column6"
COL_BODY    = "Column7"

# -----------------------------------------------------------------------------
# 7. CHIẾN LƯỢC GÁN NHÃN (LABELING STRATEGY)
#    Gmail 2 (mbox): dùng nhãn thật từ X-Gmail-Labels trong file .mbox
#      → SPAM = nhãn "Spam" hoặc "Danh mục Khuyến mại" (Promotions)
#      → HAM  = tất cả nhãn còn lại
#    Gmail 1 (mail1.csv): dùng heuristics (quy tắc) vì không có file .mbox
#      → Phân tích người gửi + từ khóa tiêu đề
# -----------------------------------------------------------------------------
GMAIL2_SPAM_LABELS = [
    "Spam",
    "Danh mục Khuyến mại",   # Promotions tab — email thương mại/quảng cáo
]

# Những domain/sender được xác định là SPAM marketing cho Gmail 1
SPAM_SENDER_PATTERNS = [
    "ecomm.lenovo.com",          # Lenovo marketing
    "recommendations@ted.com",    # TED recommendation emails
    "recommends@ted.com",
    "discover.pinterest.com",     # Pinterest promotional
    "inspire.pinterest.com",
    "ideas.pinterest.com",
    "explore.pinterest.com",
    "pinterest.com",              # Tất cả Pinterest (đều là gợi ý/quảng cáo)
    "no-reply@grab.com",          # Grab promotional (KHÔNG phải receipt)
    "facebookmail.com",           # Facebook notifications (social spam)
    "news@insideapple.apple.com", # Apple News newsletter
    "marketing",                  # Bất kỳ sender nào chứa "marketing"
    "promo",                      # Sender chứa "promo"
    "newsletter",                 # Sender chứa "newsletter"
    "noreply@autocode.com",       # Autocode platform notifications
]

# Những domain/sender chắc chắn là HAM cho Gmail 1
HAM_SENDER_PATTERNS = [
    "security@facebookmail.com",  # Bảo mật Facebook → quan trọng
    "accounts.google.com",        # Google security/account
    "forms-receipts-noreply@google.com",  # Google Forms receipts
    "appstore@insideapple.apple.com",     # Apple App Store receipts (giao dịch)
    "no_reply@email.apple.com",          # Apple account notifications
]

# Từ khóa trong TIÊU ĐỀ → gán SPAM
SPAM_SUBJECT_KEYWORDS = [
    "ưu đãi", "khuyến mãi", "giảm giá", "voucher", "deal",
    "offer", "sale", "promo", "discount", "coupon",
    "unsubscribe", "newsletter", "miễn phí", "free",
    "win", "winner", "prize", "congratulation",
    "click here", "limited time", "exclusive",
    "recommendations", "recommended for you",
]

# -----------------------------------------------------------------------------
# 8. NGƯỠNG CẢNH BÁO CLASS IMBALANCE
#    Nếu tỉ lệ spam/ham lệch hơn 30/70 → cảnh báo người dùng
# -----------------------------------------------------------------------------
IMBALANCE_THRESHOLD = 0.30  # Nếu tỉ lệ class thiểu số < 30% → cảnh báo

# -----------------------------------------------------------------------------
# 9. THAM SỐ MODEL & TRAIN
# -----------------------------------------------------------------------------
TEST_SIZE    = 0.20    # 20% dữ liệu dùng để test
RANDOM_STATE = 42      # Seed ngẫu nhiên — đặt số cố định để kết quả tái lặp được
MAX_FEATURES = 50000   # TF-IDF lấy tối đa 50,000 từ quan trọng nhất
NGRAM_RANGE  = (1, 2)  # Dùng cả unigram và bigram (1 từ và 2 từ liên tiếp)

# -----------------------------------------------------------------------------
# 10. TẠO CÁC THƯ MỤC NẾU CHƯA CÓ
#     Hàm này được gọi khi import config để đảm bảo thư mục tồn tại
# -----------------------------------------------------------------------------
os.makedirs(OUTPUT_DIR,  exist_ok=True)   # exist_ok=True: không báo lỗi nếu đã có
os.makedirs(REPORTS_DIR, exist_ok=True)
