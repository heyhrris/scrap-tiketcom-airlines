/* Panen daftar penerbangan dari halaman hasil tiket.com yang SEDANG Anda buka.
 *
 * Tidak membuat request baru — hanya membaca kartu yang sudah ter-render,
 * sambil scroll (halaman tiket.com memakai virtual scrolling, jadi DOM cuma
 * menyimpan ~12 kartu yang terlihat).
 *
 * Dipakai sebagai BOOKMARKLET "📦 Panen" (lihat worksheet.html), atau
 * ditempel di Console (Cmd+Option+J). Status tampil di pojok kanan bawah.
 *
 * Hasil menumpuk di localStorage lintas halaman; salin sekali di akhir
 * lewat bookmarklet "📋 Salin hasil" (atau ketik copy(HASIL) di Console).
 * Tumpukan direset OTOMATIS tiap ganti minggu (acuan: Senin terakhir).
 */
(async () => {
  const UI = (html, warna) => {
    let el = document.getElementById('__panen_ui');
    if (!el) {
      el = document.createElement('div');
      el.id = '__panen_ui';
      el.style.cssText = 'position:fixed;z-index:2147483647;right:16px;bottom:16px;'
        + 'max-width:360px;padding:12px 16px;border-radius:10px;'
        + 'font:14px/1.5 -apple-system,system-ui,sans-serif;background:#fff;color:#222;'
        + 'box-shadow:0 4px 18px rgba(0,0,0,.25);';
      document.body.appendChild(el);
    }
    el.style.border = `2px solid ${warna || '#06c'}`;
    el.innerHTML = html;
  };

  if (window.__PANEN_JALAN) { UI('⏳ Masih berjalan — tunggu sebentar…'); return; }
  window.__PANEN_JALAN = true;
  try {
    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
    const p = new URLSearchParams(location.search);
    const dest = (p.get('a') || '').replace(/C$/, '');   // BPNC -> BPN
    const tgl  = p.get('date') || '';
    if (!dest || !tgl) {
      UI('❌ Ini bukan halaman hasil pencarian tiket.com.<br>'
       + '<small>Buka link dari worksheet dulu, lalu klik 📦 Panen.</small>', '#c00');
      return;
    }

    // Tunggu daftar penerbangan muncul (kalau tombol diklik saat halaman masih memuat)
    UI(`⏳ <b>${dest} · ${tgl}</b><br>Menunggu daftar penerbangan…`);
    for (let i = 0; i < 30 && !document.querySelector('[class*="FlightCard_card__"]'); i++) {
      await sleep(500);
    }
    if (!document.querySelector('[class*="FlightCard_card__"]')) {
      UI('⚠ Daftar penerbangan belum muncul.<br>'
       + '<small>Tunggu halaman selesai dimuat, lalu klik 📦 Panen lagi.</small>', '#c00');
      return;
    }

    const reTime = /^\d{1,2}:\d{2}$/, reCode = /^[A-Z]{3}$/;
    const seen = new Map();

    const panen = () => {
      document.querySelectorAll('[class*="FlightCard_card__"]').forEach((card) => {
        const leaves = [];
        card.querySelectorAll('*').forEach((e) => {
          if (e.children.length === 0) {
            const t = (e.textContent || '').trim();
            if (t) leaves.push({ c: typeof e.className === 'string' ? e.className : '', t });
          }
        });
        const has = (l, s) => l.c.includes(s);
        let airline = card.querySelector('img[alt]')?.alt?.trim() || '';
        if (!airline) {
          const a = leaves.find((l) => has(l, 'Text_size_b2__') && !has(l, 'Text_variant_lowEmphasis__')
            && !has(l, 'Text_weight_regular__') && !reCode.test(l.t) && l.t.length > 2);
          airline = a ? a.t : '';
        }
        const times = leaves.filter((l) => has(l, 'Text_size_h3__') && has(l, 'Text_weight_bold__')
          && reTime.test(l.t)).map((l) => l.t);
        const codes = leaves.filter((l) => has(l, 'Text_size_b2__') && has(l, 'Text_weight_regular__')
          && !has(l, 'Text_variant_lowEmphasis__') && reCode.test(l.t)).map((l) => l.t);
        const durEl  = leaves.find((l) => has(l, 'Text_variant_lowEmphasis__') && /^\d+j/.test(l.t));
        const stopEl = leaves.find((l) => has(l, 'Text_variant_lowEmphasis__') && /langsung|transit/i.test(l.t));
        const durasi = [durEl ? durEl.t : '', stopEl ? stopEl.t : ''].filter(Boolean).join(' ').trim();
        const priceEl = leaves.find((l) => has(l, 'Text_variant_price__'));
        if (times[0] && priceEl) {
          const harga = parseInt(priceEl.t.replace(/[^\d]/g, ''), 10);
          if (!harga) return;
          const key = [airline, times[0], times[1], harga].join('|');
          seen.set(key, [dest, tgl, airline, times[0] || '', times[1] || '',
                         codes[0] || '', codes[1] || '', durasi, harga].join('|'));
        }
      });
    };

    const yAwal = window.scrollY;
    let pos = 0, diam = 0, sampaiDasar = false;
    for (let i = 0; i < 400; i++) {
      const sebelum = seen.size;
      panen();
      diam = seen.size === sebelum ? diam + 1 : 0;

      const tinggi = document.body.scrollHeight;
      sampaiDasar = (window.scrollY + window.innerHeight) >= (tinggi - 80);

      // Berhenti HANYA bila sudah benar-benar di dasar halaman DAN tidak ada
      // tambahan baru lagi. Stagnasi di tengah (halaman lambat memuat) tidak
      // dianggap selesai — jika tidak, bagian termahal di bawah bisa terlewat.
      if (sampaiDasar && diam >= 6) break;
      if (diam >= 40) break;                       // jaring pengaman

      pos = Math.min(pos + 400, tinggi);
      window.scrollTo(0, pos);
      await sleep(450);
      if (i % 6 === 0) UI(`⏳ <b>${dest} · ${tgl}</b><br>Memanen… ${seen.size} penerbangan`);
    }
    window.scrollTo(0, yAwal);

    // Tumpuk hasil semua halaman di localStorage, direset tiap ganti minggu.
    // Penanda minggu = tanggal Senin terakhir menurut jam LOKAL (bukan UTC,
    // supaya Senin dini hari WIB tidak terbaca masih hari Minggu).
    const KUNCI = 'PANEN_TIKET', KUNCI_MINGGU = 'PANEN_TIKET_MINGGU';
    const senin = new Date();
    senin.setDate(senin.getDate() - ((senin.getDay() + 6) % 7));
    const dua = (n) => String(n).padStart(2, '0');
    const mingguIni = `${senin.getFullYear()}-${dua(senin.getMonth() + 1)}-${dua(senin.getDate())}`;

    let kumpulan = [], mingguBaru = false;
    try {
      if (localStorage.getItem(KUNCI_MINGGU) !== mingguIni) {
        localStorage.removeItem(KUNCI);
        localStorage.setItem(KUNCI_MINGGU, mingguIni);
        mingguBaru = true;
      }
      kumpulan = JSON.parse(localStorage.getItem(KUNCI) || '[]');
    } catch (e) {}
    const gabung = [...new Set([...kumpulan, ...seen.values()])];
    let tersimpan = true;
    try { localStorage.setItem(KUNCI, JSON.stringify(gabung)); } catch (e) { tersimpan = false; }

    window.HASIL = gabung.join('\n');          // untuk copy(HASIL) di Console
    const halaman = new Set(gabung.map((b) => b.split('|').slice(0, 2).join('|'))).size;

    const baris = [];
    if (mingguBaru) baris.push(`🔄 Minggu baru (acuan Senin ${mingguIni}) — tumpukan lama dikosongkan`);
    baris.push(`✅ <b>${seen.size}</b> penerbangan · ${dest} ${tgl}`);
    baris.push(sampaiDasar
      ? '✔ sudah sampai dasar daftar'
      : '<b style="color:#c00">⚠ BELUM sampai dasar — klik 📦 Panen lagi</b>');
    if (tersimpan) {
      baris.push(`📦 Total terkumpul: <b>${gabung.length}</b> penerbangan · <b>${halaman}</b> halaman`);
      baris.push('<small>Tutup tab ini (Cmd+W), lanjut ke link berikutnya.</small>');
    } else {
      baris.push('<b style="color:#c00">⚠ Penyimpanan penuh — klik 📋 Salin hasil sekarang</b>');
    }
    UI(baris.join('<br>'), sampaiDasar && tersimpan ? '#0a7d3c' : '#c00');
    console.log(`✅ ${seen.size} penerbangan | total ${gabung.length} dari ${halaman} halaman`
      + (sampaiDasar ? '' : ' | ⚠ BELUM sampai dasar'));
  } finally {
    window.__PANEN_JALAN = false;
  }
})();
void 0;
