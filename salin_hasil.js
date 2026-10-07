/* Bookmarklet "📋 Salin hasil": salin seluruh tumpukan panen minggu ini
 * ke clipboard. Klik di tab tiket.com mana pun (data tersimpan per situs). */
(async () => {
  const UI = (html, warna) => {
    let el = document.getElementById('__panen_ui');
    if (!el) {
      el = document.createElement('div');
      el.id = '__panen_ui';
      el.style.cssText = 'position:fixed;z-index:2147483647;right:16px;bottom:16px;'
        + 'max-width:380px;padding:12px 16px;border-radius:10px;'
        + 'font:14px/1.5 -apple-system,system-ui,sans-serif;background:#fff;color:#222;'
        + 'box-shadow:0 4px 18px rgba(0,0,0,.25);';
      document.body.appendChild(el);
    }
    el.style.border = `2px solid ${warna || '#06c'}`;
    el.innerHTML = html;
    return el;
  };
  if (!/(^|\.)tiket\.com$/.test(location.hostname)) {
    UI('⚠ Klik tombol ini di salah satu <b>tab tiket.com</b>.<br>'
     + '<small>Data panen tersimpan di situs tiket.com, bukan di halaman ini.</small>', '#c00');
    return;
  }
  let data = [];
  try { data = JSON.parse(localStorage.getItem('PANEN_TIKET') || '[]'); } catch (e) {}
  if (!data.length) { UI('⚠ Belum ada data terkumpul minggu ini.', '#c00'); return; }
  const teks = data.join('\n');
  const halaman = new Set(data.map((b) => b.split('|').slice(0, 2).join('|'))).size;
  try {
    await navigator.clipboard.writeText(teks);
    UI(`✅ <b>${data.length}</b> penerbangan dari <b>${halaman}</b> halaman tersalin.<br>`
     + '<small>Tempel ke kumpulan.txt (Cmd+V), lalu simpan (Cmd+S).</small>', '#0a7d3c');
  } catch (e) {
    const el = UI(`<b>${data.length}</b> penerbangan dari <b>${halaman}</b> halaman.<br>`
     + 'Teks di bawah sudah terpilih — tekan <b>Cmd+C</b>, lalu tempel ke kumpulan.txt.<br>', '#06c');
    const ta = document.createElement('textarea');
    ta.value = teks;
    ta.style.cssText = 'width:100%;height:140px;margin-top:8px;font:11px monospace;';
    el.appendChild(ta);
    ta.focus(); ta.select();
  }
})();
void 0;
