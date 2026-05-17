// Servly Expenses — Service Worker
// Caches the app for full offline use after first load

var CACHE_NAME = 'servly-expenses-v1';

// Files to cache on install
var CACHE_FILES = [
  './Servly_Expenses.html',
  'https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,500;0,700;1,500&family=DM+Sans:wght@300;400;500&family=DM+Mono:wght@400;500&display=swap',
  'https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@latest/tabler-icons.min.css',
  'https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js'
];

// Install — cache all core files
self.addEventListener('install', function(e) {
  self.skipWaiting();
  e.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      // Cache the app itself (required)
      return cache.add('./Servly_Expenses.html').then(function() {
        // Cache external resources separately — don't fail if they miss
        return Promise.allSettled(
          CACHE_FILES.slice(1).map(function(url) {
            return cache.add(url).catch(function() {
              console.log('Could not cache:', url);
            });
          })
        );
      });
    })
  );
});

// Activate — clean up old caches
self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(key) {
          return key !== CACHE_NAME;
        }).map(function(key) {
          return caches.delete(key);
        })
      );
    }).then(function() {
      return self.clients.claim();
    })
  );
});

// Fetch — serve from cache first, fall back to network
self.addEventListener('fetch', function(e) {
  // Skip non-GET requests and GAS/API calls
  if (e.request.method !== 'GET') return;
  if (e.request.url.includes('script.google.com')) return;
  if (e.request.url.includes('googleapis.com/macros')) return;

  e.respondWith(
    caches.match(e.request).then(function(cached) {
      if (cached) return cached;

      // Not in cache — fetch from network and cache it
      return fetch(e.request).then(function(response) {
        // Only cache valid responses
        if (!response || response.status !== 200 || response.type === 'error') {
          return response;
        }
        var toCache = response.clone();
        caches.open(CACHE_NAME).then(function(cache) {
          cache.put(e.request, toCache);
        });
        return response;
      }).catch(function() {
        // Offline and not in cache — return the app shell
        return caches.match('./Servly_Expenses.html');
      });
    })
  );
});
