# 暗号資産 確定申告ツール

複数の取引所のCSVデータを統合し、国税庁の「暗号資産の計算書（総平均法用）」Excelを自動作成するツール。

## 対応取引所

- GMOコイン
- BitPOINT
- bitFlyer
- BitLending
- PBRレンディング
- コインチェック
- Binance（API経由）

## 対応暗号資産

取引CSVに存在する通貨を自動検出（国税庁テンプレートの上限9通貨まで）

## ファイル構成

```
├── crypto_tax.py                   # 取引CSV統合スクリプト
├── fill_nta_excel.py               # 国税庁Excel自動記入スクリプト
├── {YEAR}_crypto_transactions.csv  # 統合取引データ（出力）
├── {YEAR}_crypto_summary.csv       # カテゴリ別サマリー（出力）
├── {YEAR}_暗号資産の計算書.xlsx     # 国税庁計算書（出力）
└── data/
    ├── GMOCoin/
    ├── BitPOINT/
    ├── bitFlyer/
    ├── BitLending/
    ├── PBRLending/
    ├── Coincheck/
    ├── Binance/                    # APIキャッシュ（自動生成）
    └── NTA/                        # テンプレート (002.xlsx)
```

## 使い方

### 0. 事前準備

```bash
pip install openpyxl
```

Binance APIを使う場合は `.env` を作成：

```
API_KEY=your_api_key
SECRET=your_secret_key
FETCH_BINANCE=1
```

### 1. 取引データの統合

```bash
python crypto_tax.py --year 2025
```

各取引所のCSVを読み込み、統一フォーマットの `{YEAR}_crypto_transactions.csv` と `{YEAR}_crypto_summary.csv` を出力する。

`--year` を省略すると前年が対象になる。Binanceを使わない場合は `--no-binance` を指定。

### 2. 国税庁Excelの作成

```bash
python fill_nta_excel.py --year 2025
```

統合CSVと国税庁テンプレート（`data/NTA/002.xlsx`）をもとに、取引のあった通貨のシートを記入した `{YEAR}_暗号資産の計算書.xlsx` を出力する。

## 注意事項

- 計算方式は **総平均法** を使用
- ステーキング・レンディング報酬は **受取時に課税対象**（セクション3に記載）
- セクション4（取得価額の計算）・セクション5（所得金額の計算）はExcel内の数式で自動計算される
- 各取引所のCSVエンコーディング差異（UTF-8 BOM / Shift_JIS）は自動判別
- Binanceのステーキング報酬は受取日の終値で円換算（Klines API使用）
- 元データ（`{YEAR}_crypto_transactions.csv`、`data/Binance/{YEAR}_records.json`）は税務調査に備えて保管すること
