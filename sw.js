// ═══════════════════════════════════════════════
//  sw.js — Service Worker Markaz Belajar Invoice
//  Simpan file ini sebagai: sw.js
//  Upload ke folder yang SAMA dengan invoice.html
// ═══════════════════════════════════════════════

const CACHE_NAME = 'markaz-invoice-v1';

// Install
self.addEventListener('install', e => {
  self.skipWaiting();
});

// Activate
self.addEventListener('activate', e => {
  e.waitUntil(clients.claim());
});

// Push event — menerima notifikasi dari server
self.addEventListener('push', e => {
  if (!e.data) return;
  
  const data = e.data.json();
  
  const options = {
    body:    data.body    || 'Ada invoice baru!',
    icon:    data.icon    || '/Open_MB_Icon.png',
    badge:   data.badge   || '/Open_MB_Icon.png',
    vibrate: [200, 100, 200],
    data:    { url: data.url || '/' },
    actions: [
      { action: 'view',    title: '📋 Lihat Invoice' },
      { action: 'dismiss', title: 'Tutup' }
    ],
    requireInteraction: true,
    tag: data.tag || 'invoice-notif'
  };
  
  e.waitUntil(
    self.registration.showNotification(data.title || 'Markaz Belajar', options)
  );
});

// Klik notifikasi
self.addEventListener('notificationclick', e => {
  e.notification.close();
  
  if (e.action === 'dismiss') return;
  
  // Buka atau fokus ke tab app
  e.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then(windowClients => {
      for (const client of windowClients) {
        if (client.url.includes(self.location.origin) && 'focus' in client) {
          client.focus();
          client.postMessage({ type: 'OPEN_TRACKING' });
          return;
        }
      }
      if (clients.openWindow) return clients.openWindow('/');
    })
  );
});

// Message dari halaman (untuk trigger notif lokal)
self.addEventListener('message', e => {
  if (e.data && e.data.type === 'SHOW_NOTIF') {
    const data = e.data.payload;
    self.registration.showNotification(data.title, {
      body:    data.body,
      icon:    '/Open_MB_Icon.png',
      badge:   '/Open_MB_Icon.png',
      vibrate: [200, 100, 200],
      tag:     'invoice-local',
      requireInteraction: false
    });
  }
});
