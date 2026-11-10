// =========================================================
// サービスワーカー：ネットにつながっていなくても開けるようにするしくみ
//
// ツールのファイル（画面の部品）を、端末の中にしまっておきます。
//   ・ネットにつながるとき … いつも最新のファイルを取ってくる（取ったら、しまい直す）
//   ・ネットが遅い／つながらないとき … しまっておいたファイルで開く
// だから、ツールを更新して公開すると、次に開いたときから新しい版になります。
//
// ※ 変換する画像やPDFは、ここを通りません（端末の中だけで処理します）。
// ※ ツールのファイルを新しく足したら、下の PRECACHE にも書き足してください。
// =========================================================

const CACHE = 'file-tools-v15';   // しまっておく入れ物の名前
const SHARE_INBOX = 'share-inbox'; // スマホの「共有」で届いたファイルを、ページが受け取るまで一時的に置く所（古い入れ物の片付けでは消さない）
const NETWORK_TIMEOUT = 4000;    // ネットの返事をこれだけ（ミリ秒）待ち、来なければしまってある分を使う

// 最初にまとめてしまっておくファイル
const PRECACHE = [
  './',
  'index.html',
  'image-to-pdf.html',
  'image.html',
  'qr.html',
  'qr-read.html',
  'scan.html',
  'pdf-pages.html',
  'pdf-to-image.html',
  'pdf-compress.html',
  'pdf-write.html',
  'media.html',
  'gif.html',
  'view.html',
  'manifest.webmanifest',
  'css/common.css',
  'js/common.js',
  'js/qrcode.js',
  'js/qrread.js',
  'js/pdf.js',
  'js/pdfdoc.js',
  'js/pdfview.js',
  'js/pdfshrink.js',
  'js/pdfwrite.js',
  'js/stamp.js',
  'js/scan.js',
  'js/scancam.js',
  'js/gif.js',
  'js/zip.js',
  'js/drawingml.js',
  'js/pptx.js',
  'js/docx.js',
  'js/numfmt.js',
  'js/xlsx.js',
  'js/xlgrid.js',
  // PDF の表示部品（日本語の文字表などは、使ったときにしまう）
  'vendor/pdfjs/pdf.min.mjs',
  'vendor/pdfjs/pdf.worker.min.mjs',
  // 動画・音声の部品（MP3・FLAC の部品は、その形式を作るときだけ使う）
  'vendor/mediabunny/mediabunny.min.js',
  'vendor/mediabunny/mediabunny-mp3-encoder.min.js',
  'vendor/mediabunny/mediabunny-flac-encoder.min.js',
  'icons/icon-192.png',
  'icons/icon-512.png',
  'icons/icon-maskable-192.png',
  'icons/icon-maskable-512.png',
];

// ---- はじめて入ったとき：ファイルをまとめてしまう ----
self.addEventListener('install', (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE);
    // 1つ取れなくても、ほかはしまえるように1つずつ
    await Promise.allSettled(PRECACHE.map((path) => cache.add(new Request(path, { cache: 'no-cache' }))));
    await self.skipWaiting(); // 新しい版のサービスワーカーに、すぐ交代する
  })());
});

// ---- 交代したとき：古い入れ物を片付ける ----
self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    const names = await caches.keys();
    await Promise.all(names
      .filter((name) => name.startsWith('file-tools-') && name !== CACHE)
      .map((name) => caches.delete(name)));
    await self.clients.claim(); // 開いている画面も、すぐ受け持つ
  })());
});

// ---- ファイルを読み込むたびに ----
self.addEventListener('fetch', (event) => {
  const { request } = event;
  // スマホの「共有」→「ツール箱」でファイルが届いたとき（manifest の share_target）。
  // ファイルをしまってから、画像だけなら QR コード読み取り、それ以外は「ファイルを見る」のページを開き直す（ページがしまったファイルを読む）。
  // 古い manifest のままの端末は qr-read.html に送ってくるので、そちらも受ける
  if (request.method === 'POST' && /\/(qr-read|view)\.html$/.test(new URL(request.url).pathname)) {
    event.respondWith(receiveShare(request));
    return;
  }
  // このツール箱のファイルを「見る」ときだけ受け持つ（それ以外はブラウザにおまかせ）
  if (request.method !== 'GET' || !request.url.startsWith(self.registration.scope)) return;

  const key = cacheKey(request.url);

  // ネットから最新を取ってくる。取れたら、次のためにしまい直す。
  let saving = Promise.resolve();
  const network = fetch(key, { cache: 'no-cache' }).then((response) => {
    if (response.ok) saving = saveCopy(key, response.clone());
    return response;
  });
  // 画面に返したあとも、しまう作業を最後までやり切る
  event.waitUntil(network.then(() => saving, () => {}));

  event.respondWith((async () => {
    // ネットの返事を少しだけ待つ
    const response = await Promise.race([network, wait(NETWORK_TIMEOUT)]).catch(() => null);
    if (response?.ok) return response;
    // 遅い・つながらない・エラーのときは、しまってある分を使う
    const cached = await findInCache(key, request.mode === 'navigate');
    if (cached) return cached;
    // しまってもいなければ、ネットの返事をそのまま待つ
    return response ?? network;
  })());
});

// 「?」から後ろは無視して、同じファイルとして扱う
function cacheKey(url) {
  const { origin, pathname } = new URL(url);
  return origin + pathname;
}

async function saveCopy(key, response) {
  try {
    const cache = await caches.open(CACHE);
    await cache.put(key, response);
  } catch {
    // 容量がいっぱいなど。しまえなくても、今の表示には困らない
  }
}

async function findInCache(key, isPage) {
  const cache = await caches.open(CACHE);
  const options = { ignoreVary: true };
  const found = await cache.match(key, options);
  if (found || !isPage) return found;
  // ページそのものなら、トップページ（index.html）で代わりにする
  return (await cache.match(self.registration.scope, options))
    ?? cache.match(`${self.registration.scope}index.html`, options);
}

const IMAGE_NAME = /\.(jpe?g|jfif|png|webp|gif|bmp|avif|heic|heif)$/i;

async function receiveShare(request) {
  let page = new URL(request.url).pathname.endsWith('/qr-read.html') ? 'qr-read.html' : 'view.html';
  try {
    const form = await request.formData();
    const files = [...form.getAll('file'), ...form.getAll('image')].filter((f) => typeof f === 'object' && f.size > 0).slice(0, 20);
    const cache = await caches.open(SHARE_INBOX);
    for (const old of await cache.keys()) await cache.delete(old); // 前の残りは消す
    let n = 0;
    for (const f of files) {
      await cache.put(new Request(`${self.registration.scope}__shared/${n++}`), new Response(f, {
        headers: { 'Content-Type': f.type || 'application/octet-stream', 'X-File-Name': encodeURIComponent(f.name || '共有されたファイル') },
      }));
    }
    // 画像だけなら、これまでどおり QR コード読み取りへ（スクショの QR を読むのに使う）
    if (files.length) page = files.every((f) => (f.type || '').startsWith('image/') || IMAGE_NAME.test(f.name || '')) ? 'qr-read.html' : 'view.html';
  } catch {
    // 受け取れなくても、ページは開く（ページが「受け取れませんでした」と出す）
  }
  return Response.redirect(`${self.registration.scope}${page}?shared`, 303);
}

function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms, null));
}
