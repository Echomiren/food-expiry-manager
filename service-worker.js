const MODEL_CACHE = 'food-expiry-ocr-model-ppocrv6-tiny-v1';
const RUNTIME_CACHE = 'food-expiry-ocr-runtime-v1';

const MODEL_FILES = [
  'assets/ocr/PP-OCRv6_tiny_det_onnx_infer.tar',
  'assets/ocr/PP-OCRv6_tiny_rec_onnx_infer.tar'
];

self.addEventListener('install', event => {
  // 不在 install 階段預抓 OCR 模型，避免使用者第一次進站就吃行動流量。
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    await self.clients.claim();

    // 只清理舊 runtime cache；模型 cache 版本不變就保留。
    const names = await caches.keys();

    await Promise.all(
      names
        .filter(name =>
          name.startsWith('food-expiry-ocr-runtime-') &&
          name !== RUNTIME_CACHE
        )
        .map(name => caches.delete(name))
    );
  })());
});

function isModelRequest(url) {
  if (url.origin !== self.location.origin) return false;

  return MODEL_FILES.some(path =>
    url.pathname.endsWith('/' + path)
  );
}

function isPinnedOcrRuntimeRequest(url) {
  if (url.hostname !== 'cdn.jsdelivr.net') return false;

  return (
    url.pathname.includes('/@paddleocr/paddleocr-js@0.4.2/') ||
    url.pathname.includes('/onnxruntime-web@1.22.0/')
  );
}

async function cacheFirst(request,cacheName,{
  cacheNetworkResponse=true
}={}) {
  const cache = await caches.open(cacheName);
  const cached = await cache.match(request);

  if (cached) return cached;

  const response = await fetch(request);

  if (
    cacheNetworkResponse &&
    (response.ok || response.type === 'opaque')
  ) {
    try {
      await cache.put(
        request,
        response.clone()
      );
    } catch(e) {
      console.warn('OCR runtime cache put failed:',e);
    }
  }

  return response;
}

self.addEventListener('fetch', event => {
  const request = event.request;

  if (request.method !== 'GET') return;

  const url = new URL(request.url);

  if (isModelRequest(url)) {
    // 模型只有使用者按「下載 OCR 離線包」時才會主動放入 MODEL_CACHE。
    // 正常 OCR 使用時先從 Cache Storage 取；沒有才回網路。
    // index.html 會在呼叫 PP-OCR 前先檢查 ready marker，因此正常情況不會偷跑網路。
    event.respondWith(
      cacheFirst(
        request,
        MODEL_CACHE,
        {cacheNetworkResponse:false}
      )
    );
    return;
  }

  if (isPinnedOcrRuntimeRequest(url)) {
    // SDK / ONNX Runtime 的固定版本採 cache-first。
    event.respondWith(
      cacheFirst(
        request,
        RUNTIME_CACHE
      )
    );
  }
});
