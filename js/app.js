(function(){
'use strict';
const $=n=>document.getElementById(n),state={templates:[],rules:{},busy:false};
function status(s,kind='info'){$('progress').style.display='block';$('progress').dataset.kind=kind;$('progress').textContent=s;}
function safe(s){return (s||'artikel').normalize('NFD').replace(/[^a-zA-Z0-9-]+/g,'_').replace(/^_+|_+$/g,'').slice(0,48);}
function download(blob,name){const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);}
function path(p){return p.split('/').map(x=>encodeURIComponent(x)).join('/');}
function apiUrl(){return (window.ZAINNET_CONFIG?.apiBaseUrl || '').trim().replace(/\/+$/,'');}
function assertZipHeader(buffer,label){
 if(!buffer || buffer.byteLength<1000){
   throw Error(label+' kosong atau terlalu kecil ('+(buffer?.byteLength||0)+' byte). Pilih file DOCX yang benar, unduh ulang, atau buka dan Save As DOCX di Microsoft Word.');
 }
 const x=new Uint8Array(buffer,0,4);
 if(x[0]!==0x50||x[1]!==0x4b||x[2]!==0x03||x[3]!==0x04){
   throw Error(label+' bukan ZIP/DOCX asli (tidak memiliki tanda PK). Jangan mengganti ekstensi PDF/DOC menjadi .docx; buka file lalu Save As → Word Document (.docx).');
 }
}
async function diagnoseDocxError(err,article,template,t,file){
 // Run the more expensive ZIP inspection only on failure; spare old PCs unnecessary work.
 const message=String(err?.message||err||'');
 if(!/zip|central directory|end of data|corrupt|encrypted|document\.xml/i.test(message))return err;
 const checks=[['Artikel mahasiswa '+file.name,article],['Template '+t.name,template]];
 for(const [label,buffer] of checks){
   try{
     const zip=await JSZip.loadAsync(buffer);
     if(!zip.file('word/document.xml'))throw Error('word/document.xml tidak ditemukan');
   }catch(zErr){return Error(label+' tidak bisa dibaca sebagai DOCX valid: '+String(zErr.message||zErr)+'. Coba unduh ulang/simpan ulang file ini.');}
 }
 return Error('Artikel dan template lolos pemeriksaan ZIP, tetapi mesin pemformatan gagal: '+message+'. Kirim pesan lengkap ini untuk perbaikan mesin.');
}
function element(tag,className,text){const e=document.createElement(tag);if(className)e.className=className;if(text!==undefined)e.textContent=text;return e;}
function select(){const t=state.templates.find(x=>x.id===$('journal').value),r=t&&state.rules[t.id];if(!r)return;
 $('rules').textContent='Aturan utama: '+r.bodyFont+' '+r.bodySize+' pt, spasi '+r.line+'. '+r.note;
 const english=document.getElementById('englishFields');if(english)english.hidden=(t.id==='rabbani');
 const note=document.getElementById('rabbaniNote');if(note)note.hidden=(t.id!=='rabbani');
 $('originalLink').href=path(t.original);$('originalLink').download='';list();}
function list(){const q=$('filter').value.toLocaleLowerCase('id');$('list').replaceChildren();let n=0;
 for(const t of state.templates.filter(t=>t.name.toLocaleLowerCase('id').includes(q))){n++;
  const article=element('article','journal'+(t.id===$('journal').value?' selected':''));
  const heading=element('div','jtop'),h=element('h3','',t.name),badge=element('span','tag',t.year||'DOCX');heading.append(h,badge);
  const row=element('div','btnrow'),use=element('button','small','Pilih jurnal');use.type='button';use.addEventListener('click',()=>{$('journal').value=t.id;select();window.scrollTo({top:0,behavior:'smooth'})});
  const a=element('a','button subtle small','Unduh template rujukan');a.href=path(t.original);a.download='';row.append(use,a);
  if(t.docx&&t.docx!==t.original){const b=element('a','button subtle small','DOCX konversi');b.href=path(t.docx);b.download='';row.append(b)}
  article.append(heading,row);$('list').append(article);
 }
 $('count').textContent=n+' dari '+state.templates.length+' jurnal';}
function metadata(){return Object.fromEntries(['title','author','affiliation','email','abstractId','abstractEn','keywordsId','keywordsEn','phone','doi','volume','issue','publicationDate','received','revised','accepted','published'].map(k=>[k,$(k)?.value.trim()||'']));}
$('start').addEventListener('click',async()=>{
 if(state.busy)return;
 const file=$('file').files?.[0],t=state.templates.find(x=>x.id===$('journal').value);
 if(!t){status('Pilih jurnal terlebih dahulu.','error');return}
 if(!file||!/\.docx$/i.test(file.name)){status('Pilih artikel Word .docx (bukan PDF/.doc).','error');return}
 if(file.size===0){status('Artikel '+file.name+' memiliki ukuran 0 byte (kosong). Pilih file DOCX lain atau simpan ulang di Word terlebih dahulu.','error');return}
 if(file.size>25*1024*1024){status('File lebih dari 25 MB. Untuk stabilitas browser, kecilkan ukuran Word terlebih dahulu.','error');return}
 const base=apiUrl();
 if(!/^https:\/\//.test(base)&&!/^http:\/\/(localhost|127\.0\.0\.1)(?::\d+)?$/.test(base)){status('Server LibreOffice belum dikonfigurasi. Edit config.js dan isi apiBaseUrl dengan URL HTTPS backend Anda. File belum diproses.','error');return}
 state.busy=true;$('start').disabled=true;$('details').style.display='none';
 try{
  status('Membaca artikel dan template jurnal di browser…');
  if(!t.docx)throw Error('Path template '+t.id+' tidak tersedia dalam templates.json.');
  const templateUrl=path(t.docx)+'?v=10.1';
  const response=await fetch(templateUrl,{cache:'no-store'});if(!response.ok)throw Error('Template '+t.id+' gagal dimuat (HTTP '+response.status+').');
  const [article,template]=await Promise.all([file.arrayBuffer(),response.arrayBuffer()]);
  assertZipHeader(article,'Artikel '+file.name);
  assertZipHeader(template,'Template '+t.name+' ('+t.docx+')');
  status('Artikel '+Math.round(article.byteLength/1024)+' KB dan template '+Math.round(template.byteLength/1024)+' KB berhasil dimuat. Menerapkan aturan jurnal…');
  const userMeta=metadata();if(t.id==='rabbani'){userMeta.abstractEn='';userMeta.keywordsEn='';}
  let output;
  try{output=await ZainDocxEngine.build(article,template,state.rules[t.id],userMeta);}
  catch(buildError){throw await diagnoseDocxError(buildError,article,template,t,file);}
  // PENTING: jangan pernah mendownload hasil JSZip langsung. Output ini
  // hanya DRAFT INTERNAL. LibreOffice server HARUS menyimpan ulang Word-nya.
  status('Mengirim DOCX sementara ke server LibreOffice untuk disimpan ulang…');
  const form=new FormData();form.append('file',output.blob,'zainnet_intermediate.docx');form.append('template_id',t.id);
  const normalizedResponse=await fetch(base+'/api/normalize',{method:'POST',body:form,cache:'no-store'});
  if(!normalizedResponse.ok){
    let explanation='HTTP '+normalizedResponse.status;
    try{const err=await normalizedResponse.json();explanation=err.detail||explanation;}catch{}
    throw Error('Server LibreOffice gagal: '+explanation+'. Tidak ada file mentah yang diunduh.');
  }
  const finalDocx=await normalizedResponse.blob();
  if(finalDocx.size<1024)throw Error('Hasil server terlalu kecil/invalid; Word tidak diunduh.');
  download(finalDocx,'ZAINNET_'+safe(t.id)+'_'+safe(file.name.replace(/\.docx$/i,''))+'_V10_LO_WORD2010.docx');
  const d=output.details;status('Berhasil: file DOCX telah DISIMPAN ULANG OLEH LIBREOFFICE SERVER ('+(finalDocx.size/1024/1024).toFixed(2)+' MB). Ini bukan hasil Word langsung dari browser. Tetap periksa di Word 2010.','success');
  $('stat').textContent='Identitas diisi: '+d.frontFields+' slot · Bagian isi: '+d.contentBlocks+' blok · Tabel sumber: '+d.sourceTables+' · Relasi aset sumber: '+d.sourceAssets+' · Catatan kaki asli: '+(d.sourceFootnotes?'terdeteksi':'tidak ditemukan')+'.';
  $('warnings').replaceChildren();for(const warning of d.warnings){const li=element('li','',warning);$('warnings').appendChild(li)}$('details').style.display='block';
 }catch(e){console.error(e);status('Gagal memproses: '+e.message,'error')}finally{state.busy=false;$('start').disabled=false}
});
$('journal').addEventListener('change',select);$('filter').addEventListener('input',list);
Promise.all([fetch('templates.json?v=10.1').then(r=>{if(!r.ok)throw Error('templates.json gagal');return r.json()}),fetch('rules.json?v=10.1').then(r=>{if(!r.ok)throw Error('rules.json gagal');return r.json()})])
.then(([items,rules])=>{state.templates=items;state.rules=rules;$('journal').replaceChildren();for(const i of items)$('journal').add(new Option(i.name,i.id));$('journal').value='karsa';select()})
.catch(e=>status('Website gagal memuat daftar template. Pastikan semua isi ZIP diunggah dengan struktur folder yang benar: '+e.message,'error'));
})();
