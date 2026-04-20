# SALE一括 (Browser版)

楽天RMS × ネクストエンジン（NE）のデータを組み合わせて、セール価格計算表Excelを自動生成するWebツールです。

**完全ブラウザ処理**（Pyodide使用） → サーバー不要、CSVデータが外部送信されないため安心、GitHub Pagesに置くだけで公開できる。

---

## 🚀 今すぐ使う

👉 **[https://あなたのID.github.io/sale-ikkatsu/](https://あなたのID.github.io/sale-ikkatsu/)**

初回は30〜60秒ほど読み込みに時間がかかります（Pyodideランタイムのダウンロード）。2回目以降はキャッシュが効いて数秒で起動します。

---

## 使い方

1. 4つのCSVをアップロード:
   - **① RMSデータ** - 楽天RMSから一括DLした商品情報（Shift_JIS）
   - **② NE-セット商品コード** - ネクストエンジンのセット商品検索結果
   - **③ NE-セット構成物コード** - セット構成商品の原価情報
   - **④ NE-セット無しコード** - 単品商品の原価情報
2. 「プレビュー」で統計・警告を確認（任意）
3. 「Excelを生成してダウンロード」クリック

---

## 出力Excel仕様

| 列 | 内容 | 由来 / 計算 |
|---|---|---|
| A | 商品名 | RMSの商品名 |
| B | セット商品コード | RMS「システム連携用SKU番号」 |
| C | 商品管理番号 | RMS「商品管理番号」 |
| D | SS用タイトル | 空欄（手入力） |
| E | 通常価格 | RMS「通常購入販売価格」 |
| F | 原価 | セット商品は Σ(構成物原価 × 数量)、単品はNE-セット無しコード |
| G | 送料 | 空欄（手入力） |
| H〜K | バリエ内容1〜4 | RMSバリエーション選択肢 |
| L/M/N/O | 楽天手数料/楽天ペイ/ポイント/NE | 固定値 6/3.5/1/10 |
| P/Q/R | 10% OFF 価格/利益額/利益率 | 数式 |
| S/T/U | 20% OFF 〃 | 数式 |
| V/W/X | 30% OFF 〃 | 数式 |
| Y/Z/AA | 50% OFF 〃 | 数式 |
| AB | セット商品コード（B列コピー） | |

書式: 游ゴシック11pt、セル色分け、H〜Oグループ化、全セル罫線、1の位切り捨て。

---

## 仕組み

```
ブラウザ
  └─ index.html (UI)
      └─ Pyodide (WebAssembly版Python)
          ├─ pandas (CSV読み込み・結合)
          ├─ openpyxl (Excel書き出し)
          └─ pyapp/
              ├─ merger.py (結合ロジック)
              ├─ excel_builder.py (Excel生成)
              └─ runner.py (JSとの橋渡し)
```

**サーバーは一切使いません。** CSVデータはあなたのブラウザから一歩も外に出ません。

---

## プロジェクト構成

```
sale-ikkatsu/
├── index.html            # ブラウザアプリ本体（GitHub Pagesのルート）
├── pyapp/                # ブラウザで動くPythonコード
│   ├── merger.py
│   ├── excel_builder.py
│   └── runner.py
├── tests/                # ローカルpytest用
├── sample_data/          # サンプルCSV
├── docs/
│   └── PUBLISH_GUIDE.md  # GitHub Pages公開手順
├── README.md
└── LICENSE
```

---

## ローカル開発

```bash
# 1. 静的サーバー起動（fileスキームだとfetchが効かないので必須）
python3 -m http.server 8000

# 2. ブラウザで http://localhost:8000 を開く
```

### ロジック部分だけユニットテスト（ローカル）

```bash
pip install pandas openpyxl pytest
pytest -v
```

---

## 既知の制約

- **初回ロード**: 約10MB のダウンロード。2回目以降はブラウザキャッシュで高速
- **メモリ**: ブラウザタブあたり2GB程度まで。CSVが数万行を超える場合は注意
- **ブラウザ**: Chrome / Firefox / Safari 最新版推奨

---

## ライセンス

MIT License
