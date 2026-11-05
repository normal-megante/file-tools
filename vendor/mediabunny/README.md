# Mediabunny（動画・音声の読み書き部品）

「動画・音声の変換」（`media.html`）で使っています。中身は書き換えていません（ファイル名の拡張子だけ `.cjs` → `.js` にしました）。

| ファイル | 元 | ライセンス |
| --- | --- | --- |
| `mediabunny.min.js` | npm `mediabunny` 1.62.1 の `dist/bundles/mediabunny.min.cjs` | MPL-2.0（`LICENSE`） |
| `mediabunny-mp3-encoder.min.js` | npm `@mediabunny/mp3-encoder` 1.62.1 の `dist/bundles/mediabunny-mp3-encoder.min.js` | MPL-2.0（`LICENSE`）。中の MP3 作成部分は LAME 3.100（LGPL、`LICENSE-LAME-LGPL.txt`） |
| `mediabunny-flac-encoder.min.js` | npm `@mediabunny/flac-encoder` 1.62.1 の `dist/bundles/mediabunny-flac-encoder.min.js` | MPL-2.0（`LICENSE`）。中の FLAC 作成部分は libFLAC（BSD、`LICENSE-libFLAC.txt`） |

- Mediabunny：https://github.com/Vanilagy/mediabunny （作者 Vanilagy and contributors）
- LAME MP3 Encoder：https://lame.sourceforge.io/ （このアプリは MP3 を作るときに LAME を使っています。LAME の部分は変更していません。LAME のソースは同サイトから、WASM への組み立て方は Mediabunny のリポジトリの `packages/mp3-encoder` にあります。MP3 を読む機能（デコーダー）は入っていない版です）
- libFLAC：https://github.com/xiph/flac （Copyright (C) 2000-2009 Josh Coalson, Copyright (C) 2011-2025 Xiph.Org Foundation）

更新するときは、npm の同じ名前のパッケージの `dist/bundles/` から同じファイルを取り、ここと `media.html`・`sw.js` の版を書きかえる。
