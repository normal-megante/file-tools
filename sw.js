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

const CACHE = 'file-tools-v8';   // しまっておく入れ物の名前
const NETWORK_TIMEOUT = 4000;    // ネットの返事をこれだけ（ミリ秒）待ち、来なければしまってある分を使う

// 最初にまとめてしまっておくファイル
const PRECACHE = [
  './',
  'index.html',
  'image-to-pdf.html',
  'image.html',
  'qr.html',
  'scan.html',
  'pdf-pages.html',
  'pdf-to-image.html',
  'pdf-compress.html',
  'pdf-write.html',
  'media.html',
  'manifest.webmanifest',
  'css/common.css',
  'js/common.js',
  'js/qrcode.js',
  'js/pdf.js',
  'js/pdfdoc.js',
  'js/pdfview.js',
  'js/pdfshrink.js',
  'js/pdfwrite.js',
  'js/stamp.js',
  'js/scan.js',
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

function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms, null));
}
