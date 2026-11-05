# pdf.js（PDF の表示部品）

ファイルツール箱で PDF のページを表示する（見本の絵を作る）ために使っている部品です。

- 作者：Mozilla（Firefox の PDF 表示と同じもの） https://mozilla.github.io/pdf.js/
- 版：pdfjs-dist 6.4.299（npm から取得。中身は書き換えていません）
- ライセンス：Apache License 2.0（`LICENSE`）。商用でも無料で使えて、許諾は不要。配るときはライセンス文を同梱する。
- 同梱した部品の一部は別のライセンス（`wasm/LICENSE_*`、`standard_fonts/LICENSE_*`、`iccs/LICENSE`）

| ファイル | 内容 |
| --- | --- |
| `pdf.min.mjs` | 本体 |
| `pdf.worker.min.mjs` | 裏で PDF を読む部分 |
| `cmaps/` | 日本語などの文字の対応表（フォントが埋め込まれていない PDF 用） |
| `standard_fonts/` | PDF の標準フォントの代わり |
| `wasm/` | JPEG 2000・JBIG2 画像の読み込み、色の管理 |
| `iccs/` | CMYK の色の変換 |

更新するときは、新しい版の `pdfjs-dist` から同じファイルを置き換える。
