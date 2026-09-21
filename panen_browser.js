/* Panen daftar penerbangan dari halaman hasil tiket.com yang SEDANG Anda buka.
 *
 * Tidak membuat request baru — hanya membaca kartu yang sudah ter-render,
 * sambil scroll (halaman tiket.com memakai virtual scrolling, jadi DOM cuma
 * menyimpan ~12 kartu yang terlihat).
 *
 * CARA PAKAI
 *   1. Buka halaman hasil pencarian tiket.com di Chrome seperti biasa
 *   2. Buka DevTools Console  (Cmd+Option+J)
 *   3. Tempel seluruh isi file ini, tekan Enter
 *   4. Tunggu sampai muncul "SELESAI — N penerbangan (tersalin)"
 *   5. Hasil sudah ada di clipboard; tempel ke Claude atau ke file .txt
 */
(async () => {
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const p = new URLSearchParams(location.search);
  const dest = (p.get('a') || '').replace(/C$/, '');   // BPNC -> BPN
  const tgl  = p.get('date') || '';
  if (!dest || !tgl) { console.error('❌ Bukan halaman hasil pencarian tiket.com'); return; }

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
    if (i % 12 === 0) console.log(`  ...${seen.size} penerbangan`);
  }
  window.scrollTo(0, yAwal);

  const teks = [...seen.values()].join('\n');
  window.HASIL = teks;                       // simpan agar bisa disalin dengan copy(HASIL)
  let tersalin = false;
  try { await navigator.clipboard.writeText(teks); tersalin = true; } catch (e) {}
  console.log(`✅ SELESAI — ${seen.size} penerbangan`);
  console.log(sampaiDasar
    ? '   ✔ sudah mencapai dasar daftar (tidak ada yang terlewat)'
    : '%c   ⚠ BELUM sampai dasar — ulangi halaman ini (tekan ↑ lalu Enter)',
    sampaiDasar ? '' : 'color:#c00;font-weight:bold');
  if (tersalin) {
    console.log('   Sudah tersalin ke clipboard — tinggal tempel (Cmd+V).');
  } else {
    console.log('%c   Ketik:  copy(HASIL)   lalu Enter  → data tersalin ke clipboard',
                'color:#0a0;font-weight:bold');
  }
})();
