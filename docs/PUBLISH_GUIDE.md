# 🚀 GitHub Pagesで公開する手順

所要時間の目安: **15〜20分**（Render使用版より簡単）

## 全体の流れ

1. ZIPを展開
2. GitHubで新規リポジトリを作る
3. ターミナルでpush
4. GitHub PagesをON（Web操作）
5. 公開URL確認

---

## STEP 1. ZIPを展開してターミナルに移動

ダウンロードした `sale-ikkatsu.zip` を展開し、好きな場所に置く。

```bash
cd ~/Projects/sale-ikkatsu   # 実際に置いた場所
ls   # index.html, pyapp/, README.md などが見えればOK
```

---

## STEP 2. GitHubで新規リポジトリを作る

1. ブラウザで https://github.com/new を開く
2. 以下のように入力:
   - **Repository name**: `sale-ikkatsu`
   - **Public** を選択 ← 📌 GitHub Pagesを無料で使うにはPublic必須
   - ⚠️ 下のチェックボックス（README / .gitignore / license）は **全部OFF**
3. 「Create repository」クリック

→ 空のリポジトリ画面が表示される

---

## STEP 3. ターミナルで push

以下を **上から順にコピペ** 実行。`あなたのID` は自分のGitHubユーザー名に変えてください。

```bash
git init
git add .
git commit -m "initial commit: SALE一括 Browser版 v0.1.0"
git branch -M main
git remote add origin https://github.com/あなたのID/sale-ikkatsu.git
git push -u origin main
```

認証を求められた場合：
- macOS/Windows: 初回はブラウザが開いて GitHub ログイン
- うまくいかない場合は Personal Access Token を使用（下記トラブルシュート参照）

pushが成功したら、GitHubのリポジトリページをリロードして全ファイルが並んでいることを確認。

---

## STEP 4. GitHub Pagesを有効化

1. GitHubのリポジトリページで、**Settings**（上部タブ）をクリック
2. 左サイドバーの **「Pages」** をクリック
3. **Build and deployment** セクションで:
   - **Source**: `Deploy from a branch` を選択
   - **Branch**: `main` / `/ (root)` を選択
   - 「Save」クリック
4. 数秒〜1分待つ

→ ページ上部に以下のような緑のボックスが表示されます:

```
Your site is live at https://あなたのID.github.io/sale-ikkatsu/
```

---

## STEP 5. 公開URLで動作確認

表示されたURLをブラウザで開く:

```
https://あなたのID.github.io/sale-ikkatsu/
```

- 初回は **30〜60秒** かかる（Pyodide約10MBをDL中）
- 「Pythonランタイムを読み込んでいます…」の表示が消えれば準備完了
- サンプルCSV4つ（ZIP内 `sample_data/` にあり）をアップロードしてExcel生成→DLできればOK 🎉

---

## 💡 使い始めてから

- **2回目以降はブラウザキャッシュ**が効いて数秒で起動
- CSVは**ブラウザ内で処理**されるのでデータが外部に漏れない
- 他の人にURLを共有すれば誰でも使える
- 無料で**スリープ無し・常時公開**

---

## 🔧 困ったとき

### Q. `git push` で「Permission denied」

Personal Access Tokenを使う：
1. GitHubで https://github.com/settings/tokens → 「Generate new token (classic)」
2. Scope は「repo」にチェックしてTokenを生成・コピー
3. push時にパスワードを聞かれたらTokenを貼る

### Q. GitHub Pagesを有効化したのに404

- Pagesの設定画面で、正しく `main / root` が選ばれているか確認
- リポジトリが **Public** になっているか確認（Privateだと無料枠では使えない）
- 1〜5分待つ（デプロイに少し時間がかかることあり）

### Q. ページは開くがPyodideの読み込みでエラー

- ブラウザのDevTools（F12）のConsoleでエラーメッセージ確認
- 広告ブロッカーやコンテンツブロッカーを一時的にOFF
- シークレットウィンドウで開いてみる

### Q. CSVをアップロードしてもExcelが出ない

- Consoleでエラーを確認
- CSVの文字コードがShift_JIS / UTF-8 どちらかであることを確認
- 必須列（`商品管理番号`、`システム連携用SKU番号`など）がCSVにあるか確認
- エラー内容をチャットに貼ってください

---

## ✅ 公開後のチェックリスト

- [ ] `https://あなたのID.github.io/sale-ikkatsu/` が開ける
- [ ] サンプルCSV4つをアップロードしてExcelが落ちてくる
- [ ] 生成Excelを開いて書式が正しい
- [ ] READMEの「あなたのID」を実際のIDに書き換えて再push（任意）

---

## 🔄 後から更新したいとき

コードを修正したら、以下のコマンドで再デプロイ:

```bash
git add .
git commit -m "機能追加/バグ修正のメッセージ"
git push
```

→ GitHub Pagesが自動的に更新を反映（1〜2分）
