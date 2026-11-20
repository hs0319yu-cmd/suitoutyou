/*
 * GitHub Pages <-> Google Apps Script bridge
 * Uses a one-request-per-hidden-iframe POST so browser CORS restrictions
 * do not prevent the static GitHub Pages app from calling Apps Script.
 */
(() => {
  'use strict';
  const API_URL = 'https://script.google.com/macros/s/AKfycbxIBqXqVy03xdhrS-BM0Nzo8AZ3fqbDQ0MqYy1wGeyGgi6Qi8sLwYuENu0j926_WtRmCg/exec';
  const APP_MARK = 'suitoutyou-api';
  let sequence = 0;

  function makeId() {
    return 'req-' + Date.now().toString(36) + '-' + (++sequence).toString(36) + '-' + Math.random().toString(36).slice(2, 9);
  }

  function call(method, ...args) {
    return new Promise((resolve, reject) => {
      if (!API_URL || !/^https:\/\/script\.google\.com\/macros\/s\//.test(API_URL)) {
        reject(new Error('GASのWebアプリURLが設定されていません。api.jsを確認してください。'));
        return;
      }
      const requestId = makeId();
      const frameName = 'gas_' + requestId.replace(/[^a-zA-Z0-9_]/g, '_');
      const iframe = document.createElement('iframe');
      iframe.name = frameName;
      iframe.title = 'データ通信';
      iframe.setAttribute('aria-hidden', 'true');
      iframe.style.cssText = 'position:absolute;width:0;height:0;border:0;left:-9999px;top:-9999px';

      const form = document.createElement('form');
      form.method = 'POST';
      form.action = API_URL;
      form.target = frameName;
      form.acceptCharset = 'UTF-8';
      form.style.display = 'none';

      const payload = {
        app: APP_MARK,
        requestId,
        method: String(method || ''),
        args
      };
      const input = document.createElement('input');
      input.type = 'hidden';
      input.name = 'payload';
      input.value = JSON.stringify(payload);
      form.appendChild(input);

      let settled = false;
      const cleanup = () => {
        window.removeEventListener('message', onMessage);
        clearTimeout(timer);
        form.remove();
        iframe.remove();
      };
      const finish = (fn, value) => {
        if (settled) return;
        settled = true;
        cleanup();
        fn(value);
      };
      const onMessage = event => {
        // The iframe is created per request, so its WindowProxy is a strong
        // correlation check even though the response comes from Google.
       
        const data = event.data;
        if (!data || data.app !== APP_MARK || data.requestId !== requestId) return;
        if (data.ok) finish(resolve, data.result);
        else finish(reject, new Error(data.error || 'GASで処理に失敗しました。'));
      };
      window.addEventListener('message', onMessage);
      const timer = setTimeout(() => finish(reject, new Error('GASとの通信がタイムアウトしました。通信状態を確認して再度お試しください。')), 90000);
      document.body.appendChild(iframe);
      document.body.appendChild(form);
      try {
        form.submit();
      } catch (error) {
        finish(reject, error instanceof Error ? error : new Error(String(error)));
      }
    });
  }

  window.InventoryApi = Object.freeze({ call });

  // PWA shell only. API POSTs go to script.google.com and are not cached.
  if ('serviceWorker' in navigator && location.protocol === 'https:') {
    window.addEventListener('load', () => {
      navigator.serviceWorker.register('./service-worker.js').catch(err => {
        console.warn('Service Worker registration failed:', err);
      });
    });
  }
})();
