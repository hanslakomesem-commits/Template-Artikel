/* User interface for the static GitHub Pages version */
(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const status=$('status'), engine=$('engine'), button=$('process'), formFile=$('file');
  const sel=$('template'), profile=$('profile'), titleCount=$('titleCount');
  const download=$('download'), warns=$('warnings'), warningBox=$('warningBox');
  let oldUrl=null;
  function options(){
    const t=sel.value;
    $('profile-wrap').hidden=t!=='rabbani';
    $('layoutInfo').textContent= t==='alihkam'
      ? 'AL-IHKAM: ukuran Letter, margin dan header mengikuti file template 2026, Book Antiqua 11 pt, abstrak maksimal 250 kata.'
      :t==='rabbani'
       ? 'RABBANI: ukuran A4, Book Antiqua 12 pt, pilihan kotak abstrak 2023 atau panduan web; daftar pustaka gantung.'
       :'KARSA: dimensi halaman dan header asli dari template. Hasil perlu diperiksa karena tata letak jurnal bisa kompleks.';
  }
  function resetDownload(){
    if(oldUrl){URL.revokeObjectURL(oldUrl);oldUrl=null;}
    download.hidden=true;warningBox.hidden=true;warns.replaceChildren();
  }
  function showError(e){status.textContent='Gagal: '+(e?.message||String(e));status.style.color='#ae2c20';}
  function showStatus(text){status.textContent=text;status.style.color='#235a8f';}
  sel.addEventListener('change',()=>{options();resetDownload();});
  profile.addEventListener('change',resetDownload);
  formFile.addEventListener('change',resetDownload);
  options();
  if(!window.JSZip || !window.ZainFormatter){engine.textContent='Mesin dokumen gagal dimuat. Pastikan file vendor/jszip.min.js dan formatter.js ikut diunggah.';engine.className='warn';return;}
  engine.textContent='✓ Mesin Word lokal siap — tidak membutuhkan koneksi API / server Python';
  engine.className='good';button.disabled=false;
  button.addEventListener('click',async()=>{
    if(!formFile.files || !formFile.files[0]){showError(new Error('Pilih file artikel DOCX dahulu.'));return;}
    resetDownload();button.disabled=true;showStatus('Memulai pemformatan…');
    try {
      const meta={};document.querySelectorAll('[data-meta]').forEach(el=>{if(el.value.trim())meta[el.dataset.meta]=el.value.trim();});
      const result=await window.ZainFormatter.formatFile(formFile.files[0],{
        templateId:sel.value, profile:profile.value,titleParagraphs:Number(titleCount.value),metadata:meta
      },showStatus);
      oldUrl=URL.createObjectURL(result.blob);
      download.href=oldUrl;download.download=result.filename;download.hidden=false;
      status.textContent='Selesai. File Word siap diunduh ('+(result.blob.size/1024/1024).toFixed(2)+' MB).';
      status.style.color='#137746';
      if(result.warnings.length){warningBox.hidden=false;for(const msg of result.warnings){const li=document.createElement('li');li.textContent=msg;warns.append(li);}}
      // Automatic save from user-initiated processing. Keep download link as fallback for strict browsers.
      download.click();
    } catch(e) {showError(e);}
    finally{button.disabled=false;}
  });
  window.addEventListener('beforeunload',()=>{if(oldUrl)URL.revokeObjectURL(oldUrl);});
})();
