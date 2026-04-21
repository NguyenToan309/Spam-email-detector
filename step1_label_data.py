# =============================================================================
# step1_label_data.py — BƯỚC 1: Gán nhãn dữ liệu email
#
# Mục tiêu của file này:
#   1. Đọc file .mbox (Gmail 2) → lấy nhãn thật từ Google
#   2. Đọc mail1.csv (Gmail 1) → gán nhãn bằng heuristics (quy tắc)
#   3. Gộp 2 nguồn lại thành 1 dataset
#   4. Kiểm tra class imbalance (độ lệch nhãn)
#   5. Vẽ biểu đồ phân phối nhãn
#   6. Lưu kết quả ra labeled_dataset.csv
#
# Chạy: python step1_label_data.py
# =============================================================================

# --- Thiết lập encoding UTF-8 cho terminal Windows ---
# Windows mặc định dùng cp1252 — không hiển thị được tiếng Việt và ký tự đặc biệt
# Phải đặt TRƯỚC KHI import bất kỳ thư viện nào khác
import sys
sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

# --- Import thư viện ---
import mailbox          # Đọc file .mbox của Gmail (định dạng xuất từ Google Takeout)
import email.header     # Giải mã tiêu đề email bị mã hóa (ví dụ: =?UTF-8?B?...?=)
import csv              # Đọc/ghi file CSV
import os               # Thao tác với đường dẫn file
import re               # Regular expression — dùng để tìm kiếm pattern trong chuỗi
import pandas as pd     # Xử lý dữ liệu dạng bảng (DataFrame)
import matplotlib       # Thư viện vẽ đồ thị
matplotlib.use('Agg')   # Dùng backend không cần cửa sổ (chạy trên Windows VSCode)
import matplotlib.pyplot as plt  # API vẽ đồ thị
from tqdm import tqdm as tqdm   # Thanh tiến trình — hiển thị % khi xử lý nhiều email
from colorama import init, Fore, Style  # Màu sắc trong terminal
init(autoreset=True)    # autoreset=True: sau mỗi lần in màu, tự reset về màu trắng

# --- Import cấu hình từ config.py ---
import config

# =============================================================================
# PHẦN 1: CÁC HÀM TIỆN ÍCH
# =============================================================================

def decode_header_value(raw_value: str) -> str:
    """
    Giải mã giá trị header email bị mã hóa theo chuẩn RFC2047.

    Khi email có ký tự tiếng Việt trong tiêu đề hoặc nhãn,
    Gmail mã hóa thành dạng =?UTF-8?B?...?= hoặc =?UTF-8?Q?...?=
    Hàm này giải mã về chuỗi Unicode bình thường.

    Input:  raw_value (str) — chuỗi có thể chứa mã hóa RFC2047
    Output: (str) — chuỗi đã giải mã, đọc được bình thường
    """
    if not raw_value:                    # Nếu chuỗi rỗng thì trả về rỗng luôn
        return ""
    try:
        # decode_header() trả về list các tuple (bytes/str, encoding)
        parts = email.header.decode_header(raw_value)
        result = ""
        for part, enc in parts:
            if isinstance(part, bytes):
                # Nếu là bytes thì decode với encoding đã chỉ định, mặc định UTF-8
                result += part.decode(enc or "utf-8", errors="replace")
            else:
                result += str(part)      # Nếu đã là string thì giữ nguyên
        return result.strip()
    except Exception:
        return raw_value.strip()         # Nếu lỗi thì trả về chuỗi gốc


def classify_gmail_label(label_str: str) -> str:
    """
    Phân loại email Gmail 2 thành spam/ham dựa trên nhãn Gmail thật.

    Gmail lưu mỗi email với nhiều nhãn, ví dụ: "Hộp thư đến,Danh mục Khuyến mại"
    Hàm này đọc tất cả nhãn và quyết định nhãn cuối cùng.

    Chiến lược:
      → SPAM nếu có nhãn "Spam" hoặc "Danh mục Khuyến mại" (Promotions)
      → HAM nếu không có nhãn spam

    Input:  label_str (str) — chuỗi nhãn Gmail, các nhãn cách nhau bằng dấu phẩy
    Output: (str) — "spam" hoặc "ham"
    """
    if not label_str:
        return config.LABEL_HAM          # Không có nhãn → coi là ham

    # Tách chuỗi nhãn thành list, ví dụ "A,B,C" → ["A", "B", "C"]
    labels = [l.strip() for l in label_str.split(",")]

    # Kiểm tra từng nhãn trong GMAIL2_SPAM_LABELS (định nghĩa trong config.py)
    for spam_label in config.GMAIL2_SPAM_LABELS:
        if spam_label in labels:
            return config.LABEL_SPAM     # Có nhãn spam → trả về "spam"

    return config.LABEL_HAM              # Không có nhãn spam → "ham"


def classify_by_heuristics(sender: str, subject: str) -> str:
    """
    Gán nhãn spam/ham cho Gmail 1 (mail1.csv) bằng quy tắc heuristics.

    Vì Gmail 1 không có file .mbox nên không biết nhãn thật.
    Hàm này dùng 3 quy tắc theo thứ tự ưu tiên:
      1. Nếu sender khớp HAM_SENDER_PATTERNS → HAM ngay (ưu tiên cao nhất)
      2. Nếu sender khớp SPAM_SENDER_PATTERNS → SPAM
      3. Nếu tiêu đề chứa từ khóa spam → SPAM
      4. Mặc định → HAM

    Input:
      sender  (str) — địa chỉ người gửi, ví dụ: "Grab <no-reply@grab.com>"
      subject (str) — tiêu đề email
    Output: (str) — "spam" hoặc "ham"
    """
    sender_lower  = sender.lower()       # Chuyển về chữ thường để so sánh
    subject_lower = subject.lower()

    # Ưu tiên 1: Kiểm tra HAM trước — một số sender trông giống spam nhưng thực ra quan trọng
    for ham_pattern in config.HAM_SENDER_PATTERNS:
        if ham_pattern.lower() in sender_lower:
            return config.LABEL_HAM      # Sender quan trọng → HAM

    # Ưu tiên 2: Kiểm tra spam qua sender
    for spam_pattern in config.SPAM_SENDER_PATTERNS:
        if spam_pattern.lower() in sender_lower:
            return config.LABEL_SPAM     # Sender marketing → SPAM

    # Ưu tiên 3: Kiểm tra spam qua từ khóa trong tiêu đề
    for keyword in config.SPAM_SUBJECT_KEYWORDS:
        if keyword.lower() in subject_lower:
            return config.LABEL_SPAM     # Tiêu đề có từ khóa spam → SPAM

    return config.LABEL_HAM              # Mặc định là ham


# =============================================================================
# PHẦN 2: ĐỌC VÀ GÁN NHÃN GMAIL 2 (từ file .mbox)
# =============================================================================

def load_gmail2_from_mbox() -> list:
    """
    Đọc file .mbox của Gmail 2 và gán nhãn dựa trên X-Gmail-Labels.

    File .mbox là định dạng xuất từ Google Takeout — chứa toàn bộ email
    dưới dạng text thô, mỗi email có đầy đủ header (bao gồm X-Gmail-Labels).

    Input:  (không có — đọc từ config.MBOX_PATH)
    Output: list of dict — mỗi dict là 1 email với các trường đã chuẩn hóa
    """
    print(Fore.YELLOW + f"\n📂 Đang mở file mbox: {config.MBOX_PATH}")
    print(Fore.YELLOW + "   (File ~926MB, có thể mất 1-3 phút...)")

    records = []    # List kết quả — mỗi phần tử là 1 email dưới dạng dict

    try:
        # mailbox.mbox() mở file mbox, tự động parse từng email
        mbox = mailbox.mbox(config.MBOX_PATH)

        # tqdm() bao quanh mbox để hiển thị thanh tiến trình
        # desc= là tên hiển thị, unit= là đơn vị
        for msg in tqdm(mbox, desc="Đọc Gmail 2 (mbox)", unit="email"):
            try:
                # Lấy X-Gmail-Labels — đây là nhãn thật của Gmail
                raw_labels = msg.get("X-Gmail-Labels", "")
                label_str  = decode_header_value(raw_labels)    # Giải mã UTF-8
                label      = classify_gmail_label(label_str)    # Phân loại spam/ham

                # Lấy Subject (tiêu đề) — cũng cần giải mã vì có thể bị encode
                raw_subject = msg.get("Subject", "")
                subject     = decode_header_value(raw_subject)

                # Lấy From (người gửi)
                raw_from = msg.get("From", "")
                sender   = decode_header_value(raw_from)

                # Lấy Date (ngày gửi)
                date = msg.get("Date", "")

                # Lấy nội dung email (body) — email có thể là plain text hoặc HTML
                body = extract_body(msg)

                # Thêm vào list kết quả dưới dạng dict
                records.append({
                    "source":  "gmail2",          # Đánh dấu nguồn dữ liệu
                    "subject": subject,
                    "sender":  sender,
                    "date":    date,
                    "body":    body,
                    "label":   label,             # "spam" hoặc "ham"
                    "label_source": "gmail_api",  # Nhãn từ Gmail API thật
                })
            except Exception as e:
                # Nếu 1 email bị lỗi thì bỏ qua, không dừng toàn bộ chương trình
                continue

        mbox.close()   # Đóng file mbox sau khi đọc xong

    except FileNotFoundError:
        print(Fore.RED + f"❌ Không tìm thấy file mbox: {config.MBOX_PATH}")
        print(Fore.RED + "   Kiểm tra lại đường dẫn trong config.py → MBOX_PATH")
        return []

    # Thống kê kết quả
    spam_count = sum(1 for r in records if r["label"] == config.LABEL_SPAM)
    ham_count  = len(records) - spam_count
    print(Fore.GREEN + f"✅ Gmail 2: {len(records):,} email "
          f"({spam_count:,} spam | {ham_count:,} ham)")

    return records


# =============================================================================
# PHẦN 3: ĐỌC VÀ GÁN NHÃN GMAIL 1 (từ mail1.csv — dùng heuristics)
# =============================================================================

def load_gmail1_from_csv() -> list:
    """
    Đọc file mail1.csv (Gmail 1) và gán nhãn bằng heuristics.

    mail1.csv có 7 cột: Column1..Column7 (xem chi tiết trong config.py)
    Vì không có file .mbox nên phải dùng quy tắc để phân loại.

    Input:  (không có — đọc từ config.MAIL1_CSV)
    Output: list of dict — mỗi dict là 1 email đã được gán nhãn
    """
    print(Fore.YELLOW + f"\n📂 Đang đọc Gmail 1: {config.MAIL1_CSV}")
    records = []

    try:
        # Đọc toàn bộ CSV vào DataFrame để dễ xử lý
        # encoding="utf-8-sig": xử lý BOM (Byte Order Mark) ở đầu file
        df = pd.read_csv(config.MAIL1_CSV, encoding="utf-8-sig",
                         header=0,           # Dòng đầu là tên cột
                         low_memory=False)   # Tắt cảnh báo về kiểu dữ liệu

        # Đổi tên cột về tên chuẩn dễ dùng hơn
        # df.columns sẽ là ["Column1", "Column2", ..., "Column7"]
        df.columns = ["subject", "sender", "to", "date", "starred", "size", "body"]

        # Điền NaN (ô trống) thành chuỗi rỗng để tránh lỗi khi xử lý
        df = df.fillna("")

        # Duyệt từng dòng, hiển thị thanh tiến trình
        for _, row in tqdm(df.iterrows(), total=len(df),
                           desc="Gán nhãn Gmail 1 (heuristics)", unit="email"):
            subject = str(row["subject"])
            sender  = str(row["sender"])
            body    = str(row["body"])
            date    = str(row["date"])

            # Gán nhãn bằng heuristics
            label = classify_by_heuristics(sender, subject)

            records.append({
                "source":  "gmail1",
                "subject": subject,
                "sender":  sender,
                "date":    date,
                "body":    body,
                "label":   label,
                "label_source": "heuristic",  # Nhãn từ quy tắc, không phải Gmail thật
            })

    except FileNotFoundError:
        print(Fore.RED + f"❌ Không tìm thấy: {config.MAIL1_CSV}")
        return []
    except Exception as e:
        print(Fore.RED + f"❌ Lỗi đọc mail1.csv: {e}")
        return []

    spam_count = sum(1 for r in records if r["label"] == config.LABEL_SPAM)
    ham_count  = len(records) - spam_count
    print(Fore.GREEN + f"✅ Gmail 1: {len(records):,} email "
          f"({spam_count:,} spam | {ham_count:,} ham)")

    return records


# =============================================================================
# PHẦN 4: TRÍCH XUẤT NỘI DUNG EMAIL (BODY)
# =============================================================================

def extract_body(msg) -> str:
    """
    Trích xuất nội dung văn bản từ đối tượng email .mbox.

    Email có thể có nhiều phần (multipart): HTML, plain text, attachment...
    Hàm này ưu tiên lấy plain text, nếu không có thì lấy HTML.

    Input:  msg — đối tượng email từ mailbox.mbox
    Output: (str) — nội dung văn bản của email
    """
    body = ""

    if msg.is_multipart():
        # Email nhiều phần — duyệt từng phần
        for part in msg.walk():
            # Lấy loại nội dung: "text/plain" hoặc "text/html"
            content_type = part.get_content_type()

            # Bỏ qua attachment (file đính kèm) — chỉ lấy text
            disposition = str(part.get("Content-Disposition", ""))
            if "attachment" in disposition:
                continue

            if content_type == "text/plain":
                # Lấy nội dung plain text — decode với charset của email
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    body = payload.decode(charset, errors="replace")
                    break   # Ưu tiên plain text → dừng ngay khi tìm thấy

            elif content_type == "text/html" and not body:
                # Lấy HTML nếu chưa có plain text
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    body = payload.decode(charset, errors="replace")
    else:
        # Email đơn giản (không multipart)
        payload = msg.get_payload(decode=True)
        if payload:
            charset = msg.get_content_charset() or "utf-8"
            body = payload.decode(charset, errors="replace")

    # Giới hạn body tối đa 10,000 ký tự để tránh file CSV quá lớn
    return body[:10000]


# =============================================================================
# PHẦN 5: KIỂM TRA CLASS IMBALANCE (ĐỘ LỆCH NHÃN)
# =============================================================================

def check_class_imbalance(df: pd.DataFrame) -> None:
    """
    Kiểm tra và báo cáo tình trạng mất cân bằng nhãn trong dataset.

    Mất cân bằng nhãn (class imbalance) xảy ra khi số lượng spam và ham
    chênh lệch quá nhiều. Ví dụ: 95% ham và 5% spam → model sẽ "lười"
    và luôn đoán ham để đạt accuracy cao mà không thực sự học được.

    Input:  df (DataFrame) — dataset đã gán nhãn, có cột "label"
    Output: (không — in kết quả ra terminal và lưu biểu đồ)
    """
    print(Fore.YELLOW + "\n" + "="*60)
    print(Fore.YELLOW + "  📊 KIỂM TRA CLASS IMBALANCE (ĐỘ LỆCH NHÃN)")
    print(Fore.YELLOW + "="*60)

    # Đếm số lượng mỗi nhãn
    label_counts = df["label"].value_counts()
    total        = len(df)

    # Tính tỉ lệ phần trăm
    spam_count = label_counts.get(config.LABEL_SPAM, 0)
    ham_count  = label_counts.get(config.LABEL_HAM, 0)
    spam_pct   = spam_count / total * 100
    ham_pct    = ham_count  / total * 100

    # In bảng thống kê
    print(f"\n  {'Nhãn':<12} {'Số lượng':>10} {'Tỉ lệ':>10}")
    print(f"  {'-'*34}")
    print(f"  {'SPAM':<12} {spam_count:>10,} {spam_pct:>9.1f}%")
    print(f"  {'HAM':<12} {ham_count:>10,} {ham_pct:>9.1f}%")
    print(f"  {'-'*34}")
    print(f"  {'TỔNG':<12} {total:>10,} {'100.0%':>10}")

    # Phân tích nguồn dữ liệu
    if "source" in df.columns:
        print(Fore.YELLOW + "\n  Phân tích theo nguồn Gmail:")
        for src in df["source"].unique():
            src_df = df[df["source"] == src]
            s = src_df["label"].value_counts().get(config.LABEL_SPAM, 0)
            h = src_df["label"].value_counts().get(config.LABEL_HAM, 0)
            print(f"    {src}: {len(src_df):,} email | {s:,} spam ({s/len(src_df)*100:.1f}%) | {h:,} ham")

    # Kiểm tra ngưỡng cảnh báo
    minority_ratio = min(spam_pct, ham_pct) / 100
    if minority_ratio < config.IMBALANCE_THRESHOLD:
        print(Fore.RED + f"\n  ⚠️  CẢNH BÁO: Dataset bị LỆCH NHÃN NẶNG!")
        print(Fore.RED + f"     Tỉ lệ class thiểu số: {minority_ratio*100:.1f}% < {config.IMBALANCE_THRESHOLD*100:.0f}%")
        print(Fore.RED + f"     → Sẽ áp dụng oversampling ở Bước tiếp theo")
    else:
        print(Fore.GREEN + f"\n  ✅ Dataset cân bằng tốt (tỉ lệ tối thiểu: {minority_ratio*100:.1f}%)")

    # Vẽ biểu đồ phân phối nhãn
    _plot_label_distribution(label_counts, spam_pct, ham_pct)


def _plot_label_distribution(label_counts, spam_pct, ham_pct):
    """
    Vẽ 2 biểu đồ: cột (bar chart) và tròn (pie chart) cho phân phối nhãn.

    Input:
      label_counts — Series với số lượng mỗi nhãn
      spam_pct, ham_pct — tỉ lệ phần trăm spam và ham
    Output: lưu file ảnh vào output/reports/label_distribution.png
    """
    # Tạo figure với 2 subplot cạnh nhau
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))
    fig.suptitle("Phân phối nhãn Spam vs Ham trong Dataset",
                 fontsize=14, fontweight="bold")

    colors = ["#e74c3c", "#2ecc71"]   # Đỏ = spam, Xanh = ham

    # --- Biểu đồ cột (Bar Chart) ---
    labels_list = [config.LABEL_SPAM.upper(), config.LABEL_HAM.upper()]
    counts = [label_counts.get(config.LABEL_SPAM, 0),
              label_counts.get(config.LABEL_HAM, 0)]
    bars = ax1.bar(labels_list, counts, color=colors, edgecolor="white",
                   linewidth=1.5, width=0.5)
    ax1.set_title("Số lượng email theo nhãn")
    ax1.set_ylabel("Số lượng email")
    ax1.set_xlabel("Nhãn")
    # Thêm số lên đầu mỗi cột
    for bar, count in zip(bars, counts):
        ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(counts)*0.01,
                 f"{count:,}", ha="center", va="bottom", fontweight="bold")
    ax1.grid(axis="y", alpha=0.3)

    # --- Biểu đồ tròn (Pie Chart) ---
    wedge_labels = [f"SPAM\n{spam_pct:.1f}%", f"HAM\n{ham_pct:.1f}%"]
    ax2.pie(counts, labels=wedge_labels, colors=colors,
            autopct="%1.1f%%", startangle=90,
            wedgeprops={"edgecolor": "white", "linewidth": 2})
    ax2.set_title("Tỉ lệ phần trăm Spam vs Ham")

    plt.tight_layout()

    # Lưu ảnh vào thư mục reports
    save_path = os.path.join(config.REPORTS_DIR, "label_distribution.png")
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()   # Đóng figure để giải phóng bộ nhớ
    print(Fore.GREEN + f"\n  📈 Đã lưu biểu đồ: {save_path}")


# =============================================================================
# PHẦN 6: KIỂM TRA VÀ SỬA NHÃN NGHI VẤN
# =============================================================================

def detect_suspicious_labels(df: pd.DataFrame) -> pd.DataFrame:
    """
    Phát hiện và sửa các email có nhãn nghi vấn (gán sai).

    Quy tắc phát hiện nhãn sai:
      1. Email được gán HAM nhưng tiêu đề/sender có đặc điểm spam rõ ràng
      2. Email được gán SPAM nhưng từ địa chỉ quan trọng (Google, security)
      3. Email có body quá ngắn (< 10 ký tự) → có thể lỗi dữ liệu

    Input:  df (DataFrame) — dataset đã gán nhãn
    Output: (DataFrame) — dataset đã được sửa nhãn + cột "label_fixed" ghi lý do

    """
    print(Fore.YELLOW + "\n" + "="*60)
    print(Fore.YELLOW + "  🔍 KIỂM TRA VÀ SỬA NHÃN NGHI VẤN")
    print(Fore.YELLOW + "="*60)

    # Sao chép để không làm thay đổi DataFrame gốc
    df = df.copy()

    # Thêm cột ghi lại lý do sửa nhãn (rỗng = không sửa)
    df["label_fixed"] = ""

    fixed_count = 0   # Đếm số email được sửa

    for idx, row in tqdm(df.iterrows(), total=len(df),
                         desc="Kiểm tra nhãn", unit="email"):
        subject = str(row.get("subject", "")).lower()
        sender  = str(row.get("sender",  "")).lower()
        body    = str(row.get("body",    ""))
        label   = row["label"]
        source  = row.get("source", "")

        reason = ""   # Lý do sửa nhãn

        # Quy tắc 1: Email từ địa chỉ security/official bị gán SPAM → sửa thành HAM
        official_domains = [
            "accounts.google.com", "security@facebookmail.com",
            "no_reply@email.apple.com", "apple.com", "paypal.com",
            "linkedin.com", "github.com",
        ]
        if label == config.LABEL_SPAM:
            for domain in official_domains:
                if domain in sender:
                    reason = f"Official sender ({domain}) → sửa thành HAM"
                    break

        # Quy tắc 2: Email body quá ngắn → có thể dữ liệu lỗi, đánh dấu để xem
        if len(body.strip()) < 20 and not reason:
            reason = f"Body quá ngắn ({len(body.strip())} ký tự) → giữ nguyên nhãn, cần xem lại"
            # Không sửa nhãn, chỉ đánh dấu để theo dõi

        # Quy tắc 3: Email gian lận/phishing phổ biến bị gán HAM → sửa thành SPAM
        phishing_keywords = [
            "verify your account", "xác minh tài khoản", "click to verify",
            "your account has been", "tài khoản của bạn bị",
            "won a prize", "you have won", "trúng thưởng",
            "nigerian prince", "send money", "wire transfer",
            "password reset" , "reset your password",
        ]
        if label == config.LABEL_HAM and not reason:
            for kw in phishing_keywords:
                if kw in subject or kw in body[:500].lower():
                    reason = f"Từ khóa phishing ({kw}) → sửa thành SPAM"
                    break

        # Áp dụng sửa nhãn nếu có lý do
        if reason and ("→ sửa thành" in reason):
            if "sửa thành HAM" in reason:
                df.at[idx, "label"] = config.LABEL_HAM
            elif "sửa thành SPAM" in reason:
                df.at[idx, "label"] = config.LABEL_SPAM
            df.at[idx, "label_fixed"] = reason
            fixed_count += 1

    # In kết quả
    print(Fore.CYAN + f"\n  📝 Kết quả kiểm tra nhãn:")
    print(f"     Tổng email đã kiểm tra: {len(df):,}")
    print(f"     Số email bị sửa nhãn:   {fixed_count:,}")

    if fixed_count > 0:
        # Hiển thị một số email đã được sửa để người dùng xem
        fixed_df = df[df["label_fixed"].str.contains("sửa thành", na=False)]
        print(Fore.YELLOW + f"\n  Một số email đã được sửa nhãn (tối đa 10):")
        for i, (_, r) in enumerate(fixed_df.head(10).iterrows()):
            subj = str(r["subject"])[:50]
            print(f"    [{i+1}] {subj:<50} | {r['label_fixed']}")

    return df


# =============================================================================
# PHẦN 7: HÀM CHÍNH (MAIN)
# =============================================================================

def main():
    """
    Hàm chính — chạy toàn bộ pipeline gán nhãn dữ liệu.

    Thứ tự thực hiện:
      1. Đọc Gmail 2 từ mbox (nhãn thật)
      2. Đọc Gmail 1 từ CSV (heuristics)
      3. Gộp 2 dataset
      4. Kiểm tra và sửa nhãn nghi vấn
      5. Kiểm tra class imbalance
      6. Lưu kết quả ra CSV
    """
    print(Fore.CYAN + Style.BRIGHT + """
╔══════════════════════════════════════════════════════════╗
║       BƯỚC 1: GÁN NHÃN DỮ LIỆU EMAIL                   ║
║       Spam Classifier — UTH Project                      ║
╚══════════════════════════════════════════════════════════╝""")

    # --- Bước 1.1: Đọc Gmail 2 từ mbox ---
    print(Fore.CYAN + "\n▶ Bước 1.1: Đọc Gmail 2 từ file .mbox (nhãn thật từ Google)")
    gmail2_records = load_gmail2_from_mbox()

    # --- Bước 1.2: Đọc Gmail 1 từ CSV ---
    print(Fore.CYAN + "\n▶ Bước 1.2: Đọc Gmail 1 từ mail1.csv (gán nhãn bằng heuristics)")
    gmail1_records = load_gmail1_from_csv()

    # --- Bước 1.3: Gộp 2 dataset ---
    print(Fore.CYAN + "\n▶ Bước 1.3: Gộp dữ liệu từ 2 Gmail")
    all_records = gmail2_records + gmail1_records   # Nối 2 list

    if not all_records:
        print(Fore.RED + "❌ Không có dữ liệu nào! Kiểm tra lại đường dẫn file.")
        return

    # Chuyển list of dict thành DataFrame
    df = pd.DataFrame(all_records)
    print(Fore.GREEN + f"✅ Tổng dataset: {len(df):,} email từ 2 Gmail")

    # --- Bước 1.4: Kiểm tra và sửa nhãn nghi vấn ---
    print(Fore.CYAN + "\n▶ Bước 1.4: Kiểm tra và sửa nhãn nghi vấn")
    df = detect_suspicious_labels(df)

    # --- Bước 1.5: Kiểm tra class imbalance ---
    print(Fore.CYAN + "\n▶ Bước 1.5: Kiểm tra class imbalance (độ lệch nhãn)")
    check_class_imbalance(df)

    # --- Bước 1.6: Lưu kết quả ra CSV ---
    print(Fore.CYAN + "\n▶ Bước 1.6: Lưu dữ liệu đã gán nhãn")
    df.to_csv(config.LABELED_CSV, index=False, encoding="utf-8-sig")
    # utf-8-sig: thêm BOM để Excel mở đúng tiếng Việt

    print(Fore.GREEN + f"✅ Đã lưu: {config.LABELED_CSV}")
    print(Fore.GREEN + f"   Kích thước file: {os.path.getsize(config.LABELED_CSV)/1024/1024:.1f} MB")

    # --- Tóm tắt cuối ---
    spam_total = (df["label"] == config.LABEL_SPAM).sum()
    ham_total  = (df["label"] == config.LABEL_HAM).sum()
    print(Fore.CYAN + Style.BRIGHT + f"""
╔══════════════════════════════════════════════════════════╗
║  ✅ BƯỚC 1 HOÀN THÀNH                                   ║
║  📊 Tổng:   {len(df):>7,} email                              ║
║  🔴 Spam:   {spam_total:>7,} ({spam_total/len(df)*100:5.1f}%)                        ║
║  🟢 Ham:    {ham_total:>7,} ({ham_total/len(df)*100:5.1f}%)                        ║
║  📁 Đã lưu: labeled_dataset.csv                         ║
║  📈 Biểu đồ: output/reports/label_distribution.png      ║
╚══════════════════════════════════════════════════════════╝""")
    print(Fore.YELLOW + "\n➡  Bước tiếp theo: chạy python step2_preprocess.py")


# =============================================================================
# ĐIỂM VÀO CHƯƠNG TRÌNH
# =============================================================================

if __name__ == "__main__":
    # Kiểm tra xem các thư viện cần thiết đã được cài chưa
    try:
        import tqdm as _tqdm_mod, colorama, matplotlib, pandas
    except ImportError as e:
        print(f"❌ Thiếu thư viện: {e}")
        print("   Chạy lệnh sau để cài: pip install tqdm colorama matplotlib pandas")
        exit(1)

    main()
