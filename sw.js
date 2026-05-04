const CACHE_NAME = 'rsamio-v1.1.0';
const PREFIX = '/Monitoreo-Cadena-de-Fr-o---Farmacia';
const urlsToCache = [
  `${PREFIX}/`,
  `${PREFIX}/index.html`,
  `${PREFIX}/Vicus_farmacia.html`,
  `${PREFIX}/reporte_individual_farmacia.html`,
  `${PREFIX}/css/pwa-styles.css`,
  `${PREFIX}/js/pwa-utils.js`,
  `${PREFIX}/logos/logo_hnb_oscuro.jpg`,
  `${PREFIX}/logos/logo_rih.jpg`,
  `${PREFIX}/logos/logo_hnb_blanco.jpg`,
  `${PREFIX}/logos/footer.jpg`,
  `${PREFIX}/logos/portal-vicus.png`
];

// Instalación del Service Worker
self.addEventListener('install', event => {
  console.log('[SW] Instalando rsamio...');
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        console.log('[SW] Cacheando archivos...');
        return cache.addAll(urlsToCache);
      })
      .then(() => self.skipWaiting())
  );
});

// Activación
self.addEventListener('activate', event => {
  console.log('[SW] Activando rsamio...');
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.map(cacheName => {
          if (cacheName !== CACHE_NAME) {
            console.log('[SW] Eliminando cache antigua:', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    })
  );
  return self.clients.claim();
});

// Interceptación de peticiones
self.addEventListener('fetch', event => {
  event.respondWith(
    caches.match(event.request)
      .then(response => {
        // Si está en cache, devolverlo
        if (response) {
          return response;
        }

        // Si no está en cache, hacer la petición real
        return fetch(event.request)
          .then(response => {
            // Clonar la respuesta para guardar en cache
            const responseClone = response.clone();

            // Solo cachear si es exitoso y es GET
            if (response.status === 200 && event.request.method === 'GET') {
              caches.open(CACHE_NAME)
                .then(cache => {
                  cache.put(event.request, responseClone);
                });
            }

            return response;
          })
          .catch(() => {
            // Si falla la conexión, devolver página offline
            if (event.request.url.includes('.html')) {
              return caches.match(`${PREFIX}/index.html`);
            }
          });
      })
  );
});
