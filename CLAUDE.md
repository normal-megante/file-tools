# ファイルツール箱（file-tools）

PDF24 のような「ファイルのちょっとした作業」を、自分用に無料・回数無制限で使えるようにするツール集。
PC（Windows / Edge・Chrome）とスマホ（Android / Chrome）の両方で、アプリのように使う（PWA）。

## Claude への作業ルール

- ユーザーはプログラミング初心者。説明は日本語で、専門用語はかみ砕いて短く説明する。
- いきなり大きく作らない。各ステップの前に「何を作るか」を示し、確認を取ってから進める。
- 1ステップごとに「動くもの」を作り、ユーザーが自分で試せる状態にしてから次へ進む。
- 作業フォルダ：この CLAUDE.md があるフォルダ（`file-tools`）。開発 PC は Windows（Python / Node.js あり）。
- GitHub で公開する予定。個人の情報（本名・個人のメールアドレス・PC 内のフルパスなど）をファイルやコミットに書かない。
- 公開（GitHub への push）はユーザーが GitHub Desktop で行う。Claude はコミットまで。
- 動作確認でブラウザからファイルを保存するときは、ユーザーのダウンロードフォルダに出さない（`saveBlob` / `downloadUrl` を差し替えて中身だけ確かめる）。

## 方針

- ファイルはすべて端末の中で処理する。サーバーには送らない。
- HTML / CSS / JavaScript だけで作る（ビルド作業なし・外部ライブラリなし）。
- 1つの大きなアプリにせず、作業ごとに1ページ。トップの `index.html` がメニュー。
- PC でしかできないもの（Web ページ → PDF、Office → PDF、PDF → Word など）は、あとで PC 用の小さな Python ツールで補う予定。
- むずかしいもの・本格的な PDF 編集は、PC に入れた PDF24 Creator を使う（自作しない）。

## ファイル構成

| ファイル | 内容 |
| --- | --- |
| `index.html` | メニュー。「アプリとして追加」ボタン |
| `image-to-pdf.html` | 画像 → PDF（JPEG はそのまま埋め込み、用紙・余白・向き・画質・白黒・1つ／別々） |
| `image.html` | 画像の変換・圧縮（JPEG/PNG/WebP、画質、ファイルサイズ上限、縮小、回転・反転・切り抜き、ZIP 保存） |
| `qr.html` | QR コード作成（URL・文章、Wi-Fi、連絡先。PNG / SVG 保存、コピー、共有） |
| `css/common.css` | 全ページ共通の見た目 |
| `js/common.js` | 共通の部品（ファイル受け取り、画像一覧 `ImageList`、縮小 `downscale`、ZIP 作成、保存・共有など） |
| `js/qrcode.js` | QR コードの作成（自作。Node でもテストできる形） |
| `manifest.webmanifest` / `sw.js` / `icons/` | アプリ化（PWA）。`sw.js` は「ネット優先・つながらなければ保存分」 |
| `dev-server.py` | PC で確かめるための簡易サーバー（`python dev-server.py` → http://localhost:8140） |

## 作り方のメモ（ハマりどころ）

- 画像の書き出しは `canvas.toBlob` ではなく `toDataURL`（`encodeCanvas`）。`toBlob` は1枚1秒近く待たされることがある。
- 画像は `createImageBitmap(file, { imageOrientation: 'from-image' })` で読み、スマホ写真の向き（EXIF）を反映する。
- 縮小は半分ずつ段階的に（`downscale`）。一気に縮めると線がギザギザになる。
- 画像の編集情報は `{ rot, flip, crop }`。適用順は「回転 → 反転 → 切り抜き」。`crop` は 0〜1 の割合。
- `sw.js` の `CACHE` 名と `PRECACHE` 一覧は、ファイルを増やしたら更新する。
- 設定は `localStorage` に保存（`img2pdf-settings-v1`、`image-convert-settings-v1`、`qr-settings-v1`）。QR の入力内容は保存しない。

## 進み具合

1. ✅ **土台＋画像ツール＋QR**（2026-11-05）
   - 共通部品、画像 → PDF の移植、画像の変換・圧縮、QR コード作成、メニュー、アプリ化。
   - 確認済み：QR はバージョン1〜40×誤り訂正4段階すべて読み取り OK。画像 → PDF は PyMuPDF で正常に開ける。
2. ⬜ **スキャン補正**（ユーザーの希望が強い）
   - 書類の四隅を自動で見つけて、まっすぐに（台形補正）。四隅は手で直せる。
   - 影を消す（背景の明るさをならす）、手書きの線をくっきり、白黒書類モード。
   - ボタン1つでおまかせ。スマホのカメラから直接撮れるように。
3. ⬜ **PDF ツール**：圧縮、PDF → 画像、結合・分割・並べ替え・回転・ページ削除など。
4. ⬜ **変換（PC のみ）**：Web ページ → PDF、Office → PDF、PDF → Word など（Python ツール）。

## 公開

- まだ未公開。公開するなら GitHub Pages（例：`normal-megante.github.io/file-tools/`）。
- 公開すればスマホからも開けて、「ホーム画面に追加」でアプリのように使える。
