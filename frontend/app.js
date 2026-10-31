/* ZAIN.NET Multi Template Artikel - koneksi mesin otomatis, tanpa API key. */
(() => {
  'use strict';

  const status = document.getElementById('status');
  const connection = document.getElementById('connection');
  const retry = document.getElementById('retry');
  const btn = document.getElementById('process');
  const sel = document.getElementById('template');
  const profile = document.getElementById('profile');
  const titleCount = document.getElementById('titleCount');
  const fileInput = document.getElementById('file');
  const config = window.ZAINNET_CONFIG || {};
  const configuredBase = typeof config.apiBaseUrl === 'string' ? config.apiBaseUrl.trim().replace(/\/+$/, '') : '';
  const apiBase = configuredBase || window.location.origin;
  const isGithubPages = /\.github\.io$/i.test(window.location.hostname);
  const originalProfiles = profile.innerHTML;
  let connected = false;

  function setConnection(message, state) {
    connection.textContent = message;
    connection.className = 'connection ' + state;
  }

  function updateOptions() {
    const template = sel.value;
    const options = document.getElementById('rabbaniOptions');
    options.hidden = template === 'karsa';
    const current = profile.value;
    profile.innerHTML = template === 'alihkam'
      ? '<option value="template2026">File template Al-Ihkam 2026</option>'
      : originalProfiles;
    if ([...profile.options].some(o => o.value === current)) profile.value = current;
    document.getElementById('layoutInfo').textContent = template === 'alihkam'
      ? 'Letter, margin atas/kiri 5 cm, bawah/kanan 4,5 cm. Book Antiqua. Abstract maksimal 250 kata; 3–5 kata kunci.'
      : 'A4, margin 3 cm, Book Antiqua. Gunakan judul bagian pada paragraf tersendiri.';
  }
  sel.addEventListener('change', updateOptions);

  async function readJSON(response) {
    const type = response.headers.get('content-type') || '';
    if (!response.ok || !type.includes('application/json')) throw new Error('Server tidak merespons sebagai mesin pemroses ZAIN.NET.');
    return response.json();
  }

  async function connect() {
    btn.disabled = true;
    connected = false;
    retry.hidden = true;
    status.textContent = '';
    setConnection('Menghubungkan mesin pemroses...', 'checking');
    try {
      if (window.location.protocol === 'file:') throw new Error('Jalankan web melalui server, bukan dengan membuka file HTML langsung.');
      if (configuredBase && !/^https?:\/\//i.test(configuredBase)) throw new Error('Alamat backend di config.js harus berawalan https:// atau http://.');
      if (window.location.protocol === 'https:' && apiBase.startsWith('http://')) throw new Error('Alamat backend harus menggunakan HTTPS.');

      const health = await readJSON(await fetch(apiBase + '/health', { cache: 'no-store' }));
      if (health.ok !== true) throw new Error('Mesin pemroses belum siap.');
      const templates = await readJSON(await fetch(apiBase + '/templates', { cache: 'no-store' }));
      if (!Array.isArray(templates) || !templates.length) throw new Error('Daftar template dari backend kosong.');
      const selected = sel.value;
      sel.replaceChildren();
      for (const item of templates) {
        if (!item || typeof item.id !== 'string' || typeof item.name !== 'string') continue;
        const option = document.createElement('option');
        option.value = item.id;
        option.textContent = item.name;
        sel.append(option);
      }
      if (!sel.options.length) throw new Error('Data template dari backend tidak valid.');
      if ([...sel.options].some(o => o.value === selected)) sel.value = selected;
      updateOptions();
      connected = true;
      btn.disabled = false;
      setConnection('Mesin pemroses terhubung. Siap memformat dokumen.', 'online');
    } catch (error) {
      setConnection('Mesin pemroses belum terhubung.', 'offline');
      retry.hidden = false;
      const setupHint = isGithubPages && !configuredBase
        ? 'Website ini di GitHub Pages. Admin perlu mengisi apiBaseUrl pada file config.js dengan URL hosting backend Python (satu kali saja).'
        : !configuredBase
          ? 'Pastikan aplikasi dijalankan dari server FastAPI ZAIN.NET, atau atur apiBaseUrl pada config.js.'
          : 'Pastikan server backend aktif, alamat di config.js benar, dan CORS mengizinkan domain web ini.';
      status.textContent = setupHint + '\nRincian: ' + (error.message || 'Koneksi gagal.');
    }
  }
  retry.addEventListener('click', connect);

  btn.addEventListener('click', async () => {
    if (!connected) { status.textContent = 'Mesin pemroses belum terhubung. Klik Coba Sambungkan Lagi.'; return; }
    const file = fileInput.files[0];
    if (!file || !/\.docx$/i.test(file.name)) { status.textContent = 'Pilih artikel berformat DOCX terlebih dahulu.'; return; }
    if (file.size > 15_000_000) { status.textContent = 'Ukuran DOCX maksimal 15 MB.'; return; }
    btn.disabled = true;
    status.textContent = 'Sedang memproses dokumen. Jangan tutup halaman ini.';
    try {
      const form = new FormData();
      const chosen = sel.value;
      form.append('file', file);
      form.append('template_id', chosen);
      form.append('profile', profile.value);
      form.append('title_paragraphs', titleCount.value);
      form.append('metadata', JSON.stringify(Object.fromEntries([...document.querySelectorAll('[data-metadata]')].map(el => [el.dataset.metadata, el.value]))));
      const response = await fetch(apiBase + '/format', { method: 'POST', body: form });
      if (!response.ok) {
        let detail = 'Kesalahan server (' + response.status + ').';
        try { const err = await response.json(); if (typeof err.detail === 'string') detail = err.detail; } catch (_ignored) {}
        throw new Error(detail);
      }
      let warnings = [];
      try { warnings = JSON.parse(decodeURIComponent(response.headers.get('X-Format-Warnings') || '%5B%5D')); } catch (_ignored) {}
      const blob = await response.blob();
      if (!blob.size) throw new Error('Hasil pemrosesan kosong.');
      const blobUrl = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = blobUrl;
      link.download = 'ZAINNET_' + chosen.toUpperCase() + '_BETA.docx';
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(blobUrl), 60000);
      status.textContent = 'Dokumen selesai diproses. Unduhan telah dimulai. Periksa hasil Word sebelum digunakan.' +
        (Array.isArray(warnings) && warnings.length ? '\n\nCatatan pemeriksaan:\n• ' + warnings.join('\n• ') : '');
    } catch (error) {
      status.textContent = 'Gagal memproses dokumen: ' + (error.message || 'Terjadi kesalahan jaringan.');
    } finally {
      btn.disabled = !connected;
    }
  });

  updateOptions();
  connect();
})();
