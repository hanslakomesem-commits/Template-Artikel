(function(){
'use strict';
const $=id=>document.getElementById(id);
const state={items:[],profiles:{},busy:false};
function fmt(n){return (n/1024/1024).toFixed(2)+' MB'}
function msg(text,kind='info'){$('status').textContent=text;$('status').dataset.kind=kind}
function safeFilename(text){return text.normalize('NFD').replace(/[^a-zA-Z0-9_-]+/g,'_').slice(0,50)}
function save(blob,name){const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000)}
function href(path){return path.split('/').map(s=>encodeURIComponent(s)).join('/')}
function make(tag,cls,text){const e=document.createElement(tag);if(cls)e.className=cls;if(text!==undefined)e.textContent=text;return e}
function updateFilter(){const q=$('filter').value.toLowerCase().trim();$('catalog').replaceChildren();let count=0;
 for(const p of state.items.filter(i=>i.name.toLowerCase().includes(q))){count++;
  const card=make('article','journal-card'),badge=make('span','small-tag',p.year||''),title=make('h3','',p.name),subtitle=make('p','mini',p.docx&&p.docx!==p.original?'Asli .doc + konversi .docx':'Template Word asli');
  const row=make('div','card-actions'),a=make('a','small-link','Download asli');a.href=href(p.original);a.download='';row.appendChild(a);
  if(p.docx && p.docx!==p.original){const b=make('a','small-link subtle','DOCX konversi');b.href=href(p.docx);b.download='';row.appendChild(b)}
  const choose=make('button','small-link outline','Pilih AUTO dasar');choose.type='button';choose.addEventListener('click',()=>{$('journal').value=p.id;window.scrollTo({top:0,behavior:'smooth'});$('file').focus()});row.appendChild(choose);
  card.append(title,badge,subtitle,row);$('catalog').appendChild(card);
 }
 $('count').textContent=count+' dari '+state.items.length+' jurnal';
}
function updateSelection(){const i=state.items.find(i=>i.id===$('journal').value);if(!i)return;
 $('templateInfo').textContent='Acuan: '+i.name+' · '+(i.year||'')+'. AUTO dasar menyesuaikan ukuran kertas, margin, dan tipografi. Tidak menjamin tata letak sama persis.';
 $('chosenOriginal').href=href(i.original);$('chosenOriginal').download='';
}
async function start(){
 try{
  const [items,profiles]=await Promise.all([fetch('templates.json').then(r=>{if(!r.ok)throw Error('templates.json tidak ditemukan');return r.json()}),fetch('profiles.json').then(r=>{if(!r.ok)throw Error('profiles.json tidak ditemukan');return r.json()})]);
  state.items=items;state.profiles=profiles;
  $('journal').replaceChildren();for(const item of items){const op=new Option(item.name,item.id);$('journal').append(op)}
  $('journal').value='karsa';updateSelection();updateFilter();
  $('catalogIntro').textContent=items.length+' jurnal tersedia · template asli dapat diunduh';
 }catch(e){msg('Gagal memuat daftar template: '+e.message,'error')}
}
$('filter').addEventListener('input',updateFilter);
$('journal').addEventListener('change',updateSelection);
$('start').addEventListener('click',async()=>{
 const input=$('file').files[0],selected=state.items.find(i=>i.id===$('journal').value);
 if(state.busy)return;
 if(!input||!/\.docx$/i.test(input.name)){msg('Pilih file artikel dalam format .docx (bukan .doc atau PDF).','error');return}
 if(!selected||!selected.docx){msg('Template DOCX tidak ditemukan.','error');return}
 if(input.size>35*1024*1024){msg('Ukuran artikel melebihi batas 35 MB untuk proses browser.','error');return}
 state.busy=true;$('start').disabled=true;
 try{
  msg('Membaca artikel dan template Word di browser…','info');
  const res=await fetch(href(selected.docx));if(!res.ok)throw Error('Template tidak dapat diunduh (HTTP '+res.status+').');
  const tpl=await res.arrayBuffer();const article=await input.arrayBuffer();
  const opts={justify:$('justify').checked,indent:$('indent').checked,headerFooter:$('branding').checked,pageNumbers:$('pageNumber').checked};
  msg('Menerapkan ukuran halaman, margin, font dan aturan dasar…','info');
  const output=await window.ZainDocxEngine.build(article,tpl,state.profiles[selected.id],opts);
  save(output.blob,'ZAINNET_'+safeFilename(selected.id)+'_'+safeFilename(input.name.replace(/\.docx$/i,''))+'_BETA.docx');
  const d=output.details;
  msg('Berhasil: DOCX dasar sudah dibuat dan diunduh ('+fmt(output.blob.size)+'). '+d.formattedRuns+' bagian teks disesuaikan, '+d.headersFooters+' komponen header/footer disalin. Periksa hasil di Microsoft Word sebelum dicetak/dikirim jurnal.','success');
 }catch(e){msg('Tidak dapat memproses artikel: '+e.message,'error');console.error(e)}finally{state.busy=false;$('start').disabled=false}
});
start();
})();
