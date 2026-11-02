(function(){
'use strict';
const $=n=>document.getElementById(n),state={templates:[],rules:{},busy:false};
function status(s,kind='info'){$('progress').style.display='block';$('progress').dataset.kind=kind;$('progress').textContent=s;}
function safe(s){return (s||'artikel').normalize('NFD').replace(/[^a-zA-Z0-9-]+/g,'_').replace(/^_+|_+$/g,'').slice(0,48);}
function download(blob,name){const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);}
function path(p){return p.split('/').map(x=>encodeURIComponent(x)).join('/');}
function element(tag,className,text){const e=document.createElement(tag);if(className)e.className=className;if(text!==undefined)e.textContent=text;return e;}
function select(){const t=state.templates.find(x=>x.id===$('journal').value),r=t&&state.rules[t.id];if(!r)return;
 $('rules').textContent='Aturan utama: '+r.bodyFont+' '+r.bodySize+' pt, spasi '+r.line+'. '+r.note;
 $('originalLink').href=path(t.original);$('originalLink').download='';list();}
function list(){const q=$('filter').value.toLocaleLowerCase('id');$('list').replaceChildren();let n=0;
 for(const t of state.templates.filter(t=>t.name.toLocaleLowerCase('id').includes(q))){n++;
  const article=element('article','journal'+(t.id===$('journal').value?' selected':''));
  const heading=element('div','jtop'),h=element('h3','',t.name),badge=element('span','tag',t.year||'DOCX');heading.append(h,badge);
  const row=element('div','btnrow'),use=element('button','small','Pilih jurnal');use.type='button';use.addEventListener('click',()=>{$('journal').value=t.id;select();window.scrollTo({top:0,behavior:'smooth'})});
  const a=element('a','button subtle small','Unduh asli');a.href=path(t.original);a.download='';row.append(use,a);
  if(t.docx&&t.docx!==t.original){const b=element('a','button subtle small','DOCX konversi');b.href=path(t.docx);b.download='';row.append(b)}
  article.append(heading,row);$('list').append(article);
 }
 $('count').textContent=n+' dari '+state.templates.length+' jurnal';}
function metadata(){return Object.fromEntries(['title','author','affiliation','email','abstractId','abstractEn','keywordsId','keywordsEn','received','revised','accepted','published'].map(k=>[k,$(k)?.value.trim()||'']));}
$('start').addEventListener('click',async()=>{
 if(state.busy)return;
 const file=$('file').files?.[0],t=state.templates.find(x=>x.id===$('journal').value);
 if(!t){status('Pilih jurnal terlebih dahulu.','error');return}
 if(!file||!/\.docx$/i.test(file.name)){status('Pilih artikel Word .docx (bukan PDF/.doc).','error');return}
 if(file.size>25*1024*1024){status('File lebih dari 25 MB. Untuk stabilitas browser, kecilkan ukuran Word terlebih dahulu.','error');return}
 state.busy=true;$('start').disabled=true;$('details').style.display='none';
 try{
  status('Membaca artikel dan template jurnal asli di browser…');
  const response=await fetch(path(t.docx));if(!response.ok)throw Error('Template '+t.id+' gagal dimuat (HTTP '+response.status+').');
  const [article,template]=await Promise.all([file.arrayBuffer(),response.arrayBuffer()]);
  status('Mengisi identitas, abstrak, dan isi artikel pada struktur Word jurnal…');
  const output=await ZainDocxEngine.build(article,template,state.rules[t.id],metadata());
  download(output.blob,'ZAINNET_'+safe(t.id)+'_'+safe(file.name.replace(/\.docx$/i,''))+'_TEMPLATE_ASLI.docx');
  const d=output.details;status('Berhasil membuat DOCX ('+(output.blob.size/1024/1024).toFixed(2)+' MB) dari template '+t.name+'. Periksa hasil di Word sebelum digunakan.','success');
  $('stat').textContent='Identitas diisi: '+d.frontFields+' slot · Bagian isi: '+d.contentBlocks+' blok · Tabel sumber: '+d.sourceTables+' · Relasi aset sumber: '+d.sourceAssets+' · Catatan kaki asli: '+(d.sourceFootnotes?'terdeteksi':'tidak ditemukan')+'.';
  $('warnings').replaceChildren();for(const warning of d.warnings){const li=element('li','',warning);$('warnings').appendChild(li)}$('details').style.display='block';
 }catch(e){console.error(e);status('Gagal memproses: '+e.message,'error')}finally{state.busy=false;$('start').disabled=false}
});
$('journal').addEventListener('change',select);$('filter').addEventListener('input',list);
Promise.all([fetch('templates.json').then(r=>{if(!r.ok)throw Error('templates.json gagal');return r.json()}),fetch('rules.json').then(r=>{if(!r.ok)throw Error('rules.json gagal');return r.json()})])
.then(([items,rules])=>{state.templates=items;state.rules=rules;$('journal').replaceChildren();for(const i of items)$('journal').add(new Option(i.name,i.id));$('journal').value='karsa';select()})
.catch(e=>status('Website gagal memuat daftar template. Pastikan semua isi ZIP diunggah dengan struktur folder yang benar: '+e.message,'error'));
})();
