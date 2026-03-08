# 暗号資産 確定申告ツール（2025年分）

複数の取引所のCSVデータを統合し、国税庁の「暗号資産の計算書（総平均法用）」Excelを自動作成するツール。

## 対応取引所

- GMOコイン
- BitPOINT
- bitFlyer
- BitLending
- PBRレンディング
- コインチェック

## 対応暗号資産

BTC, ETH, XRP, DOGE, BAT, ATOM, TRX, SOL, BNB

## ファイル構成

```
├── crypto_tax.py                  # 取引CSV統合スクリプト
├── fill_nta_excel.py              # 国税庁Excel自動記入スクリプト
├── 2025_crypto_transactions.csv   # 統合取引データ（出力）
├── 2025_crypto_summary.csv        # カテゴリ別サマリー（出力）
├── 2025_暗号資産の計算書.xlsx      # 国税庁計算書（出力）
└── data/
    ├── GMOCoin/
    ├── BitPOINT/
    ├── bitFlyer/
    ├── BitLending/
    ├── PBRLending/
    ├── Coincheck/
    └── NTA/                       # テンプレート (002.xlsx)
```

## 使い方

### 1. 取引データの統合

```bash
python crypto_tax.py
```

各取引所のCSVを読み込み、統一フォーマットの `2025_crypto_transactions.csv` と `2025_crypto_summary.csv` を出力する。

### 2. 国税庁Excelの作成

```bash
python fill_nta_excel.py
```

統合CSVと国税庁テンプレート(`data/NTA/002.xlsx`)をもとに、暗号資産ごとのシート（計算書①〜⑨）を記入した `2025_暗号資産の計算書.xlsx` を出力する。

## 注意事項

- 計算方式は **総平均法** を使用
- ステーキング・レンディング報酬は **受取時に課税対象**（セクション3に記載）
- セクション4（取得価額の計算）・セクション5（所得金額の計算）はExcel内の数式で自動計算される
- 各取引所のCSVエンコーディング差異（UTF-8 BOM / Shift_JIS）は自動判別
