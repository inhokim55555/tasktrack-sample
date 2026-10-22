self.addEventListener('push', event => {
  let payload = {};
  try {
    payload = event.data ? event.data.json() : {};
  } catch (_error) {
    payload = { title: 'Friday notification', body: event.data ? event.data.text() : '' };
  }

  const title = payload.title || 'Friday notification';
  const options = {
    body: payload.body || '',
    icon: './favicon.png',
    badge: './favicon.png',
    tag: payload.tag || payload.id || 'friday-notification',
    data: {
      url: payload.url || './',
      notificationId: payload.id || '',
      taskId: payload.taskId || '',
    },
  };

  event.waitUntil(self.registration.showNotification(title, options));
});

self.addEventListener('notificationclick', event => {
  event.notification.close();
  const targetUrl = new URL(event.notification.data?.url || './', self.location.origin).href;

  event.waitUntil((async () => {
    const windows = await clients.matchAll({ type: 'window', includeUncontrolled: true });
    const existing = windows.find(client => client.url.startsWith(self.location.origin));
    if (existing) {
      await existing.focus();
      existing.postMessage({
        type: 'friday-notification-click',
        notificationId: event.notification.data?.notificationId || '',
        taskId: event.notification.data?.taskId || '',
      });
      return;
    }
    await clients.openWindow(targetUrl);
  })());
});
