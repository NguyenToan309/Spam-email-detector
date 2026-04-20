# 🔍 Spam Email Classifier — Phân loại Email Rác

**Chủ đề 16** | Môn: Trí tuệ nhân tạo | Trường: Đại học Công nghệ TP.HCM (UTH)

## Mô tả

Hệ thống phân loại email rác (Spam/Ham) sử dụng thuật toán **Multinomial Naive Bayes** kết hợp **TF-IDF**, huấn luyện trên dữ liệu thực từ 2 tài khoản Gmail cá nhân. Hỗ trợ email **tiếng Việt** (underthesea) và **tiếng Anh**.

## Kết quả

| Chỉ số | Giá trị |
|--------|---------|
| Accuracy | ~95%+ |
| Model | Multinomial Naive Bayes |
| Dataset | ~39,000 email thực |

## Cấu trúc dự án

```
spam-classifier/
├── spam_classifier.ipynb   # Notebook chính — chạy từ đầu đến cuối
├── requirements.txt        # Thư viện cần cài
├── README.md               # File này
├── mail1.csv               # Dữ liệu Gmail 1 (cần tự thêm)
├── messages.csv            # Dữ liệu Gmail 2 (cần tự thêm)
└── output/                 # Tự động tạo khi chạy notebook
    ├── labeled_dataset.csv
    ├── processed_dataset.csv
    ├── models/
    │   ├── naive_bayes_model.pkl
    │   └── tfidf_vectorizer.pkl
    └── reports/
        ├── spam_report.xlsx
        ├── spam_summary.txt
        ├── eda_analysis.png
        ├── class_balance.png
        ├── evaluation.png
        └── feature_importance.png
```

## Cài đặt & Chạy

```bash
# 1. Clone repo
git clone <repo-url>
cd spam-classifier

# 2. Cài thư viện
pip install -r requirements.txt

# 3. Mở notebook trong VSCode
# Mở file spam_classifier.ipynb
# Chọn kernel Python 3.13
# Run All Cells (Ctrl+Shift+P → "Run All")
```

## Lưu ý

- Đặt file `mail1.csv` và `messages.csv` vào cùng thư mục với notebook
- Nếu có file `.mbox`, cập nhật đường dẫn `MBOX_PATH` trong cell **Cấu hình**
- Lần đầu chạy sẽ mất ~2-5 phút để xử lý dữ liệu (các lần sau nhanh hơn nhờ cache)

## Công nghệ

- **Python 3.13** | **Jupyter Notebook** | **VSCode**
- `scikit-learn` — Naive Bayes, TF-IDF, metrics
- `underthesea` — Tách từ tiếng Việt
- `imbalanced-learn` — Xử lý mất cân bằng nhãn
- `matplotlib` / `seaborn` — Biểu đồ trực quan
- `openpyxl` — Xuất báo cáo Excel
