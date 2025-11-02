const cacheName = 'customwordstyling-cache';
const filestoCache = [
    './',
    './index.html',
    './downloader.html',
    './logo_dark.svg',
    './logo_light.svg',
    './assets/ChartAddLabelEveryX.jpg',
    './assets/ChartAngle0.jpg',
    './assets/ChartAngle45.jpg',
    './assets/ChartAxis2D.jpg',
    './assets/ChartAxis3D.jpg',
    './assets/ChartDataLabel.jpg',
    './assets/ChartDistanceLength1.jpg',
    './assets/ChartDistanceLength2.jpg',
    './assets/ChartExplosion0.jpg',
    './assets/ChartExplosion15.jpg',
    './assets/ChartLegendShadow.jpg',
    './assets/ChartRotateLabelAxis.jpg',
    './assets/ChartTitleShadow.jpg',
    './assets/ListTabPosition.jpg',
    './assets/ListNumberFormat.jpg',
    './assets/IndentPositive.jpg',
    './assets/CellSpacing.jpg',
    './assets/ListSecondLineTextPosition.jpg',
    './assets/ListSecondLineTextPosition2.jpg',
    './assets/IndentNegative.jpg',
    './assets/ParagraphSameLine.gif',
    './assets/ParagraphSamePage.gif',
    './assets/Shading.jpg',
    './assets/index.js',
    './assets/index.css',
    'https://appsforoffice.microsoft.com/lib/1/hosted/office.js'
];
self.addEventListener('install', e => {
    e.waitUntil(
        caches.open(cacheName)
            .then(cache => cache.addAll(filestoCache))
    );
    e.skipWaiting();
});
self.addEventListener('activate', e => self.clients.claim());
self.addEventListener('fetch', event => {
    const req = event.request;
    if (req.url.indexOf("updatecode") !== -1 || req.url.indexOf(".mp4") !== -1) event.respondWith(fetch(req)); else event.respondWith(networkFirst(req));
});

async function networkFirst(req) {
    try {
        const networkResponse = await fetch(req);
        const cache = await caches.open(cacheName);
        await cache.delete(req);
        await cache.put(req, networkResponse.clone());
        return networkResponse;
    } catch (error) {
        const cachedResponse = await caches.match(req);
        return cachedResponse;
    }
}