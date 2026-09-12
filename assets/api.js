/* Слой обращений к Google Apps Script */
(function () {
  'use strict';

  function isConfigured() {
    var u = (window.CONFIG && window.CONFIG.API_URL) || '';
    return /^https?:\/\/.+/.test(u) && u.indexOf('ВСТАВЬТЕ') === -1;
  }

  function base() { return (window.CONFIG && window.CONFIG.API_URL) || ''; }

  function get(params) {
    var qs = Object.keys(params)
      .map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]); })
      .join('&');
    return fetch(base() + '?' + qs, { method: 'GET', redirect: 'follow' })
      .then(function (r) { return r.json(); });
  }

  // text/plain — чтобы браузер не делал preflight-запрос (Apps Script его не поддерживает)
  function post(body) {
    return fetch(base(), {
      method: 'POST',
      redirect: 'follow',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(body)
    }).then(function (r) { return r.json(); });
  }

  window.API = { get: get, post: post, isConfigured: isConfigured };
})();
