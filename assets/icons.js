/* Иконки (Lucide-подобные, 24×24, stroke) */
(function () {
  var s = function (d) {
    return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" ' +
           'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + d + '</svg>';
  };
  window.ICON = {
    menu:   s('<path d="M3 6h18M3 12h18M3 18h18"/>'),
    close:  s('<path d="M18 6 6 18M6 6l12 12"/>'),
    check:  s('<path d="M20 6 9 17l-5-5"/>'),
    lock:   s('<rect x="4" y="10" width="16" height="10" rx="2"/><path d="M8 10V7a4 4 0 0 1 8 0v3"/>'),
    book:   s('<path d="M4 19.5V5a2 2 0 0 1 2-2h13v18H6.5A2.5 2.5 0 0 0 4 19.5Z"/><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H19"/>'),
    sun:    s('<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>'),
    moon:   s('<path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8Z"/>'),
    arrow:  s('<path d="M5 12h14M13 6l6 6-6 6"/>'),
    back:   s('<path d="M19 12H5M11 18l-6-6 6-6"/>'),
    search: s('<circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/>'),
    copy:   s('<rect x="9" y="9" width="12" height="12" rx="2"/><path d="M5 15V5a2 2 0 0 1 2-2h10"/>'),
    users:  s('<path d="M16 20v-1.5a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4V20"/><circle cx="9" cy="7" r="3.5"/><path d="M22 20v-1.5a4 4 0 0 0-3-3.87"/><path d="M16 3.6a4 4 0 0 1 0 7"/>'),
    refresh:s('<path d="M20 11A8 8 0 0 0 6.3 6.3L3 9"/><path d="M3 4v5h5"/><path d="M4 13a8 8 0 0 0 13.7 4.7L21 15"/><path d="M21 20v-5h-5"/>'),
    print:  s('<path d="M6 9V3h12v6"/><rect x="6" y="14" width="12" height="7"/><path d="M6 17H4a2 2 0 0 1-2-2v-3a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v3a2 2 0 0 1-2 2h-2"/>'),
    info:   s('<circle cx="12" cy="12" r="9"/><path d="M12 11v5M12 8h.01"/>')
  };
})();
