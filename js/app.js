(function(){
'use strict';
const $=n=>document.getElementById(n),state={templates:[],rules:{},busy:false};
function status(s,kind='info'){$('progress').style.display='block';$('progress').dataset.kind=kind;$('progress').textContent=s;}
function safe(s){return (s||'artikel').normalize('NFD').replace(/[^a-zA-Z0-9-]+/g,'_').replace(/^_+|_+$/g,'').slice(0,48);}
function download(blob,name){const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);}
function path(p){return p.split('/').map(x=>encodeURIComponent(x)).join('/');}
function apiUrl(){return (window.ZAINNET_CONFIG?.apiBaseUrl || '').trim().replace(/\/+$/,'');}
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
 if(file.size>25*1024*1024){status('File lebih dari 25 MB. Untuk stabilitas browser, kecilkan ukuran Word terlebih dahulu.','error');return}
 const base=apiUrl();
 if(!/^https:\/\//.test(base)&&!/^http:\/\/(localhost|127\.0\.0\.1)(?::\d+)?$/.test(base)){status('Server LibreOffice belum dikonfigurasi. Edit config.js dan isi apiBaseUrl dengan URL HTTPS backend Anda. File belum diproses.','error');return}
 state.busy=true;$('start').disabled=true;$('details').style.display='none';
 try{
  status('Membaca artikel dan template jurnal di browser…');
  const response=await fetch(path(t.docx));if(!response.ok)throw Error('Template '+t.id+' gagal dimuat (HTTP '+response.status+').');
  const [article,template]=await Promise.all([file.arrayBuffer(),response.arrayBuffer()]);
  status('Menerapkan isi artikel dan aturan jurnal, lalu memeriksa struktur DOCX…');
  const userMeta=metadata();if(t.id==='rabbani'){userMeta.abstractEn='';userMeta.keywordsEn='';}
  const output=await ZainDocxEngine.build(article,template,state.rules[t.id],userMeta);
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
Promise.all([fetch('templates.json?v=10.0').then(r=>{if(!r.ok)throw Error('templates.json gagal');return r.json()}),fetch('rules.json?v=10.0').then(r=>{if(!r.ok)throw Error('rules.json gagal');return r.json()})])
.then(([items,rules])=>{state.templates=items;state.rules=rules;$('journal').replaceChildren();for(const i of items)$('journal').add(new Option(i.name,i.id));$('journal').value='karsa';select()})
.catch(e=>status('Website gagal memuat daftar template. Pastikan semua isi ZIP diunggah dengan struktur folder yang benar: '+e.message,'error'));
})();
