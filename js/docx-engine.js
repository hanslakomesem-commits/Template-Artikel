/* ZAIN.NET template-first DOCX compositor — client-side only.
 * LGPL? Application code: MIT. JSZip separately licensed (MIT/GPL).
 * Retains journal OOXML package and replaces front-matter placeholders in situ.
 * Imports source document body, drawing dependencies, and footnotes.
 * Intended for review before journal submission; not a guarantee of compliance.
 */
(function (global) {
'use strict';
const NS={
 w:'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
 r:'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
 rel:'http://schemas.openxmlformats.org/package/2006/relationships',
 ct:'http://schemas.openxmlformats.org/package/2006/content-types',
 xml:'http://www.w3.org/XML/1998/namespace'
};
const DOCRELPATH='word/_rels/document.xml.rels';
const CTYPES='[Content_Types].xml';
const formatMime={png:'image/png',jpg:'image/jpeg',jpeg:'image/jpeg',gif:'image/gif',bmp:'image/bmp',svg:'image/svg+xml',emf:'image/x-emf',wmf:'image/x-wmf',tif:'image/tiff',tiff:'image/tiff',bin:'application/octet-stream'};
function parse(x){const d=new DOMParser().parseFromString(x,'application/xml');if(d.getElementsByTagName('parsererror').length)throw Error('XML DOCX rusak');return d;}
function serialize(x){return new XMLSerializer().serializeToString(x);}
function children(e,n,ns=NS.w){return Array.from(e.childNodes).filter(x=>x.nodeType===1&&x.localName===n&&x.namespaceURI===ns);}
function first(e,n,ns=NS.w){return children(e,n,ns)[0]||null;}
function get(e,n,ns=NS.w){return Array.from(e.getElementsByTagNameNS(ns,n));}
function make(d,n){return d.createElementNS(NS.w,'w:'+n);}
function aval(n,key,ns=NS.w){return n?n.getAttributeNS(ns,key):null;}
function attr(n,key,v,ns=NS.w,prefix='w'){n.setAttributeNS(ns,prefix+':'+key,String(v));}
function textOf(p){return get(p,'t').map(x=>x.textContent||'').join('').replace(/\s+/g,' ').trim();}
function replaceText(p,value){
 if(!p)return;
 const tnodes=get(p,'t');
 if(!tnodes.length){const r=make(p.ownerDocument,'r'),t=make(p.ownerDocument,'t');t.setAttributeNS(NS.xml,'xml:space','preserve');t.textContent=String(value||'');r.appendChild(t);p.appendChild(r);return;}
 tnodes[0].textContent=String(value||'');tnodes[0].setAttributeNS(NS.xml,'xml:space','preserve');
 for(const t of tnodes.slice(1))t.textContent='';
 // Remove instruction-only red/highlight character formatting from filled data.
 for(const r of get(p,'r')){
  const rp=first(r,'rPr');if(!rp)continue;
  for(const h of children(rp,'highlight'))rp.removeChild(h);
  for(const c of children(rp,'color')){
   const v=(aval(c,'val')||'').toLowerCase();
   if(['ff0000','f00','ff0001','ed1c24'].includes(v))attr(c,'val','000000');
  }
 }
}

function titleFriendly(s){
 const original=(s||'').replace(/\s+/g,' ').trim();
 if(!original)return '';
 const letters=(original.match(/[A-Za-zÀ-ÿ]/g)||[]).length;
 const uppers=(original.match(/[A-ZÀ-Þ]/g)||[]).length;
 // Respect mixed-case source, but normalize all-caps university paper titles.
 if(letters && uppers/letters > .85){
  return original.toLocaleLowerCase('id').replace(/(^|[\s(\[{:—–-])([a-zà-ÿ])/g,(_,a,b)=>a+b.toLocaleUpperCase('id'))
   .replace(/\b(uin|paud|spss|ict|pai|umkm|doi)\b/gi,x=>x.toUpperCase());
 }
 return original;
}
function shortTitle(title,max=52){
 const s=titleFriendly(title).replace(/[\s,.…:;!]+$/g,'');
 if(s.length<=max)return s;
 const words=s.split(/\s+/),out=[];
 for(const word of words){if((out.join(' ')+(out.length?' ':'')+word).length>max-2)break;out.push(word)}
 return (out.length?out.join(' '):s.slice(0,Math.max(18,max-3)).replace(/\s+\S*$/,''))+'…';
}
function displayAuthor(text){
 const s=(text||'').replace(/\s+/g,' ').trim();
 if(!s)return '';
 const letters=(s.match(/[A-Za-zÀ-ÿ]/g)||[]).length,upper=(s.match(/[A-ZÀ-Þ]/g)||[]).length;
 return letters&&upper/letters>.85?s.toLocaleLowerCase('id').replace(/\b[a-zà-ÿ]/g,c=>c.toLocaleUpperCase('id')):s;
}
function cleanHeading(s,rule){
 s=s.replace(/^\s*(?:[A-G]|\d{1,2})\s*[.)]\s*/i,'').trim();
 if(rule.headingCase==='upper')return s.toLocaleUpperCase('id');
 if(rule.headingCase==='title')return s.replace(/\w[^\s]*/g,m=>m[0].toLocaleUpperCase('id')+m.slice(1).toLocaleLowerCase('id'));
 return s;
}
function getBody(doc){const b=get(doc,'body')[0];if(!b)throw Error('DOCX tidak memiliki badan dokumen');return b;}
function bodyEls(doc){return Array.from(getBody(doc).childNodes).filter(x=>x.nodeType===1);}
function locateP(doc,i){const n=bodyEls(doc)[Number(i)];return n&&n.namespaceURI===NS.w&&n.localName==='p'?n:null;}
function locateCellP(doc,bi,ri,ci,pi){const tbl=bodyEls(doc)[Number(bi)];if(!tbl||tbl.localName!=='tbl')return null;
 const row=children(tbl,'tr')[Number(ri)],cell=row&&children(row,'tc')[Number(ci)];return cell?children(cell,'p')[Number(pi)]||null:null;}
function sliceCount(str){return (str||'').trim().split(/\s+/).filter(Boolean).length;}
function fromSource(srcDoc){
 const nodes=bodyEls(srcDoc),lines=nodes.map(n=>n.localName==='p'?textOf(n):'');
 const intro=/^(?:(?:[A-H]|\d{1,2})[.)]\s*)?(?:PENDAHULUAN|INTRODUCTION|LATAR BELAKANG|BACKGROUND)(?:\b|\s|\W)/i;
 let start=lines.findIndex((s,i)=>i>0&&intro.test(s)&&s.length<110);
 const abstractHead=/^(?:ABSTRAK|ABSTRACT)\s*:?(?:\s+|$)/i;
 const keywordHead=/^(?:KATA\s*KUNCI|KEYWORDS?)\s*:/i;
 if(start<0)start=lines.findIndex((s,i)=>i>0&&/^(?:A\.|1\.)\s+(?:PENDAHULUAN|INTRODUCTION)/i.test(s));
 if(start<0)throw Error('Bagian Pendahuluan/Introduction tidak ditemukan pada artikel. Tambahkan judul bagian tersebut di Word.');
 const front=nodes.slice(0,start).filter(n=>n.localName==='p').map(n=>textOf(n));
 const title=front.find(s=>s.length>12&&!/^\d+$/.test(s))||'';
 const abs={id:'',en:'',ki:'',ke:''};let mode='';
 for(const line of front){if(!line)continue;
  const match=abstractHead.exec(line);
  if(match){mode=/^ABSTRAK\b/i.test(line)?'id':'en';const rest=line.slice(match[0].length).trim();if(rest&&!/^[:;]+$/.test(rest))abs[mode]+=(abs[mode]?' ':'')+rest;continue;}
  const k=keywordHead.exec(line);if(k){const target=/^KATA/i.test(line)?'ki':'ke';abs[target]=line.slice(k[0].length).trim();mode='';continue;}
  if(mode){abs[mode]+=(abs[mode]?' ':'')+line;}
 }
 const authorLines=[];let afterTitle=false;
 for(const s of front){if(!afterTitle){if(s===title)afterTitle=true;continue;}
  if(abstractHead.test(s)||keywordHead.test(s))break;
  if(s)authorLines.push(s);
 }
 const emails=authorLines.filter(x=>/\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/i.test(x));
 const affiliations=authorLines.filter(x=>/^(?:program studi|prodi|fakultas|universitas|university|department|institut|sekolah tinggi|institut)/i.test(x));
 const authors=authorLines.filter(x=>!emails.includes(x)&&!affiliations.includes(x)&&!/^\b(?:telp|no\. hp|phone)\b/i.test(x));
 const content=nodes.slice(start).filter(n=>!(n.localName==='sectPr'));
 const hasNotes=content.some(n=>get(n,'footnoteReference').length>0);
 return {title,author:authors[0]||'',affiliation:affiliations[0]||'',email:emails[0]||'',
 abstractId:abs.id,abstractEn:abs.en,keywordsId:abs.ki,keywordsEn:abs.ke,
 content,frontLines:front,start,paragraphs:content.filter(n=>n.localName==='p').length,
 tables:content.filter(n=>n.localName==='tbl').length,hasNotes};
}
function unifyMeta(info,meta){const out={...info};
 for(const k of ['title','author','affiliation','email','abstractId','abstractEn','keywordsId','keywordsEn'])if((meta||{})[k]?.trim())out[k]=meta[k].trim();
 return out;
}
function valueOf(slot,info,meta,warnings,rule){
 const en=info.abstractEn||'',id=info.abstractId||'';
 const ke=info.keywordsEn||info.keywordsId||'',ki=info.keywordsId||info.keywordsEn||'';
 const vals={title:titleFriendly(info.title),shortTitle:shortTitle(info.title,52),author:displayAuthor(info.author),affiliation:info.affiliation,email:info.email,
 abstractEn:en,abstractId:id,abstractIdBrackets:id?'['+id+']':'',abstractEnLabel:en?'Abstract: '+en:'Abstract: —',
 abstractEnPrefix:en?'ABSTRACT  '+en:'ABSTRACT',keywordsEn:ke,keywordsId:ki,
 keywordsEnLabel:ke?'Keywords: '+ke:'Keywords: —',dateReceived:'Received: '+(meta?.received||'—'),
 dateRevised:'Revised: '+(meta?.revised||'—'),dateAccepted:'Accepted: '+(meta?.accepted||'—'),
 datePublished:'Published: '+(meta?.published||'—'),
 dateHistory:'Article history: Received: '+(meta?.received||'—')+'; Revised: '+(meta?.revised||'—')+'; Accepted: '+(meta?.accepted||'—')+'; Published: '+(meta?.published||'—')};
 if(slot==='')return '';
 if(slot.startsWith('literal:'))return slot.slice(8);
 if(slot==='abstractEn'&&!en){warnings.push('Abstrak Bahasa Inggris tidak ditemukan; bagian template dikosongkan.');return '—';}
 if(slot==='abstractId'&&!id){warnings.push('Abstrak Bahasa Indonesia tidak ditemukan; bagian template dikosongkan.');return '—';}
 if(slot==='abstractIdBrackets'&&!id){warnings.push('Abstrak Bahasa Indonesia belum tersedia.');return '';}
 return vals[slot]===undefined?String(meta?.[slot]||''):String(vals[slot]||'');
}
function replaceFront(tDoc,rule,info,meta,warnings){
 let filled=0,missing=[];
 for(const [idx,slot] of Object.entries(rule.fields||{})){
  const p=locateP(tDoc,idx);
  if(!p){missing.push('p'+idx);continue;}
  replaceText(p,valueOf(slot,info,meta,warnings,rule));filled++;
 }
 for(const [bi,ri,ci,pi,slot] of rule.table||[]){
  const p=locateCellP(tDoc,bi,ri,ci,pi);
  if(!p){missing.push('t'+bi+'/'+ri+'/'+ci+'/'+pi);continue;}
  replaceText(p,valueOf(slot,info,meta,warnings,rule));filled++;
 }
 if(missing.length)warnings.push('Posisi template berubah: '+missing.join(', '));
 return filled;
}
function fitNuansaLongTitle(tDoc,info,warnings){
 const p=locateP(tDoc,2);if(!p)return;
 const texts=get(p,'t');if(texts.length<3)return;
 // NUANSA has a fixed-size Word drawing and AlternateContent fallback with TWO title copies.
 // Fill both representations, otherwise Word 2010 and modern renderers may show different text.
 texts[0].textContent=info.title;texts[0].setAttributeNS(NS.xml,'xml:space','preserve');
 texts[2].textContent=info.title;texts[2].setAttributeNS(NS.xml,'xml:space','preserve');
 for(let i=0;i<texts.length;i++)if(i!==0&&i!==2)texts[i].textContent='';
 if(info.title.length>72){
  const pt=info.title.length>180?10:info.title.length>130?11:12;
  for(const size of get(p,'sz'))attr(size,'val',pt*2);
  for(const size of get(p,'szCs'))attr(size,'val',pt*2);
  warnings.push('NUANSA: judul panjang disesuaikan menjadi '+pt+' pt agar tidak terpotong di dalam kotak asli. Sesuaikan judul dengan batas maksimal jurnal.');
 }
}
function findPrototype(doc,rule){
 const a=bodyEls(doc);const hi=Number(rule.cut);
 const heading=a[hi];let body=null;
 for(let i=hi+1;i<a.length;i++){if(a[i].localName!=='p')continue;if(textOf(a[i]).length>60){body=a[i];break}}
 return {heading,body};
}
function isHeading(s){return /^(?:(?:[A-H]|\d{1,2})[.)]\s*)?(?:PENDAHULUAN|INTRODUCTION|METODE(?:\s+PENELITIAN)?|METHODS?|RESULTS?(?:\s+AND\s+DISCUSSION)?|HASIL(?:\s+PENELITIAN)?(?:\s+DAN\s+PEMBAHASAN)?|PEMBAHASAN|DISCUSSION|KESIMPULAN(?:\s+DAN\s+SARAN)?|CONCLUSION(?:S)?|DAFTAR\s+PUSTAKA|REFERENCES|LITERATURE\s+REVIEW|TINJAUAN\s+PUSTAKA|SARAN|PENUTUP|ACKNOWLEDGEMENTS?)(?:\b|\s*\W)/i.test(s.trim()) &&s.length<100;}
function setFont(r,rule,heading=false){
 const doc=r.ownerDocument,rpr=first(r,'rPr')||make(doc,'rPr');if(!rpr.parentNode)r.insertBefore(rpr,r.firstChild);
 let rf=first(rpr,'rFonts');if(!rf){rf=make(doc,'rFonts');rpr.insertBefore(rf,rpr.firstChild)}
 for(const k of ['ascii','hAnsi','cs','eastAsia'])attr(rf,k,rule.bodyFont);
 let sz=first(rpr,'sz');if(!sz){sz=make(doc,'sz');rpr.appendChild(sz)}attr(sz,'val',Math.round((heading?(rule.headingSize||rule.bodySize):rule.bodySize)*2));
 let szcs=first(rpr,'szCs');if(!szcs){szcs=make(doc,'szCs');rpr.appendChild(szcs)}attr(szcs,'val',Math.round((heading?(rule.headingSize||rule.bodySize):rule.bodySize)*2));
 if(heading&&!first(rpr,'b'))rpr.appendChild(make(doc,'b'));
}
function applyParaFormatting(p,kind,rule,proto){
 const doc=p.ownerDocument,old=first(p,'pPr');
 if(old)for(const v of children(old,'sectPr'))old.removeChild(v);
 // Clone journal's actual paragraph alignment, borders and spacing, without its section break.
 if(proto){const src=first(proto,'pPr');if(src){const repl=doc.importNode(src,true);for(const v of children(repl,'sectPr'))repl.removeChild(v);if(old)p.replaceChild(repl,old);else p.insertBefore(repl,p.firstChild);}}
 const ppr=first(p,'pPr')||make(doc,'pPr');if(!ppr.parentNode)p.insertBefore(ppr,p.firstChild);
 for(const ps of children(ppr,'pStyle'))ppr.removeChild(ps); // source style ID may not exist in journal.
 // Imported paragraphs often carry the source alignment; the journal body must
 // override it. Figures are centered separately in normalizeArticleNode().
 let jc=first(ppr,'jc');if(!jc){jc=make(doc,'jc');ppr.appendChild(jc)}
 attr(jc,'val',kind==='heading'?'left':'both');
 const sp=first(ppr,'spacing')||make(doc,'spacing');if(!sp.parentNode)ppr.appendChild(sp);
 if(rule.exactLinePt){attr(sp,'line',Math.round(rule.exactLinePt*20));attr(sp,'lineRule','exact');}
 else if(rule.line){attr(sp,'line',Math.round(rule.line*240));attr(sp,'lineRule','auto');}
 if(kind==='heading'){
  const k=first(ppr,'keepNext')||make(doc,'keepNext');if(!k.parentNode)ppr.appendChild(k);
 }else if(kind==='body' && rule.indentTwips){
  const ind=first(ppr,'ind')||make(doc,'ind');if(!ind.parentNode)ppr.appendChild(ind);
  if(!ind.hasAttributeNS(NS.w,'firstLine')&&!ind.hasAttributeNS(NS.w,'hanging'))attr(ind,'firstLine',rule.indentTwips);
 }
 for(const r of get(p,'r'))if(get(r,'t').length)setFont(r,rule,kind==='heading');
}
function availablePageWidth(doc){
 const sec=get(doc,'sectPr').find(x=>first(x,'pgSz')&&first(x,'pgMar'));
 const sz=sec&&first(sec,'pgSz'),mg=sec&&first(sec,'pgMar');
 const width=Number(aval(sz,'w'))||11906,left=Number(aval(mg,'left'))||1701,right=Number(aval(mg,'right'))||1701;
 return Math.max(3500,width-left-right-200);
}
function fitSourceDrawing(p,rule){
 const wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
 const a='http://schemas.openxmlformats.org/drawingml/2006/main';
 const maxWidth=(rule.pageWidthTwips||6500)*635-165000;
 for(const d of get(p,'drawing')){
  for(const x of get(d,'extent',wp)){
   const width=Number(x.getAttribute('cx'))||0,height=Number(x.getAttribute('cy'))||0;
   if(width<=maxWidth||width===0)continue;
   const scale=maxWidth/width,newW=Math.floor(width*scale),newH=Math.floor(height*scale);
   x.setAttribute('cx',String(newW));x.setAttribute('cy',String(newH));
   // Image's picture transform uses the same EMU size and must be scaled too.
   for(const ext of get(d,'ext',a)){
    if(!ext.hasAttribute('cx'))continue;
    const ew=Number(ext.getAttribute('cx')),eh=Number(ext.getAttribute('cy'));
    if(Math.abs(ew-width)<1000){ext.setAttribute('cx',String(newW));ext.setAttribute('cy',String(newH));}
    else if(ew>maxWidth){ext.setAttribute('cx',String(Math.floor(ew*scale)));ext.setAttribute('cy',String(Math.floor(eh*scale)));}
   }
  }
 }
}
function fitSourceTable(tbl,rule){
 const max=rule.pageWidthTwips||6500;
 const grid=first(tbl,'tblGrid'),cols=grid?children(grid,'gridCol'):[];
 const total=cols.reduce((a,x)=>a+(Number(aval(x,'w'))||0),0);
 if(!total||total<=max)return;
 const factor=max/total;
 for(const col of cols)attr(col,'w',Math.floor((Number(aval(col,'w'))||0)*factor));
 for(const x of get(tbl,'tcW'))if(aval(x,'type')==='dxa')attr(x,'w',Math.floor((Number(aval(x,'w'))||0)*factor));
 const props=first(tbl,'tblPr');if(props){
  const width=first(props,'tblW');if(width&&aval(width,'type')==='dxa')attr(width,'w',Math.min(max,Math.floor((Number(aval(width,'w'))||max)*factor)));
 }
}
function normalizeArticleNode(node,rule,prototypes){
 if(node.localName==='p'){
  const txt=textOf(node),heading=isHeading(txt),hasDrawing=get(node,'drawing').length>0;
  if(heading){ // change existing text but preserve footnote markers (headings rarely have notes).
   replaceText(node,cleanHeading(txt,rule));
  }
   applyParaFormatting(node,heading?'heading':'body',rule,heading?prototypes.heading:prototypes.body);
  if(hasDrawing){
   const pp=first(node,'pPr');
   if(pp){for(const ind of children(pp,'ind'))pp.removeChild(ind);
    let jc=first(pp,'jc');if(!jc){jc=make(node.ownerDocument,'jc');pp.appendChild(jc)}attr(jc,'val','center');}
   fitSourceDrawing(node,rule);
  }
 }else if(node.localName==='tbl'){
  fitSourceTable(node,rule);
  // Source figure/table structure is preserved; normalize fonts but never flatten table.
  for(const p of get(node,'p')){
   for(const r of get(p,'r'))if(get(r,'t').length)setFont(r,rule,false);
  }
 }
 // Avoid source section breaks that would override imported journal's running headers.
 for(const p of get(node,'p'))for(const pp of children(p,'pPr'))for(const s of children(pp,'sectPr'))pp.removeChild(s);
 return node;
}
function appendMain(tDoc,rule,srcContent,prototypes,warnings){
 const body=getBody(tDoc),els=bodyEls(tDoc),cut=Number(rule.cut);
 if(!els[cut]||els[cut].localName==='sectPr')throw Error('Penanda badan utama jurnal tidak valid: '+rule.id);
 const originalSec=els.findLast(n=>n.localName==='sectPr')||null;
 for(let i=els.length-1;i>=cut;i--){if(els[i]!==originalSec)body.removeChild(els[i]);}
 let count=0;
 for(const sourceEl of srcContent){if(sourceEl.localName==='sectPr')continue;
  const node=tDoc.importNode(sourceEl,true);
  normalizeArticleNode(node,rule,prototypes);
  if(originalSec)body.insertBefore(node,originalSec);else body.appendChild(node);
  count++;
 }
 if(!originalSec){warnings.push('Dokumen jurnal tidak memiliki sectPr akhir; header/footer dapat berbeda.');}
 return count;
}
function normal(p){const b=[];for(const x of p.split('/')){if(!x||x==='.')continue;if(x==='..')b.pop();else b.push(x);}return b.join('/');}
function dir(p){return p.slice(0,p.lastIndexOf('/')+1);}
function bas(p){return p.split('/').pop();}
function relpath(from,to){const a=from.split('/').filter(Boolean),b=to.split('/').filter(Boolean);while(a.length&&b.length&&a[0]===b[0]){a.shift();b.shift()}return '../'.repeat(a.length)+b.join('/');}
function relXml(x){return x?parse(x):parse('<Relationships xmlns="'+NS.rel+'"/>');}
function relNodes(doc){return get(doc,'Relationship',NS.rel);}
function xmlRelFile(part){return dir(part)+'_rels/'+bas(part)+'.rels';}
function ctOverride(ct,part,type){if(get(ct,'Override',NS.ct).some(n=>n.getAttribute('PartName')==='/'+part))return;
 const node=ct.createElementNS(NS.ct,'Override');node.setAttribute('PartName','/'+part);node.setAttribute('ContentType',type);ct.documentElement.appendChild(node);}
function ctDefault(ct,ext,type){if(get(ct,'Default',NS.ct).some(n=>n.getAttribute('Extension')===ext))return;
 const node=ct.createElementNS(NS.ct,'Default');node.setAttribute('Extension',ext);node.setAttribute('ContentType',type);ct.documentElement.appendChild(node);}
function rIdDoc(existing){let n=500;const used=new Set(relNodes(existing).map(x=>x.getAttribute('Id')));while(used.has('rIdZN'+n))n++;return 'rIdZN'+n;}
async function importSourceParts(srcZip,tplZip,tDoc,srcDoc,notes,warnings){
 const srcCt=parse(await srcZip.file(CTYPES).async('string'));
 const tplCt=parse(await tplZip.file(CTYPES).async('string'));
 const templateRels=relXml(await tplZip.file(DOCRELPATH).async('string'));
 const sourceRels=relXml(await srcZip.file(DOCRELPATH)?.async('string'));
 const originals=new Map(),idMap=new Map();let assetCount=0;
 const ignoredDocTypes=new Set(['styles','settings','fontTable','theme','webSettings','numbering','footnotes','endnotes','comments','header','footer','glossaryDocument']);
 async function copyPart(oldPath,newPath){
  oldPath=normal(oldPath);if(originals.has(oldPath))return originals.get(oldPath);
  const file=srcZip.file(oldPath);if(!file)return null;
  newPath=newPath||('word/zain_src/'+(oldPath.startsWith('word/')?oldPath.slice(5):oldPath.replace(/\//g,'_')));
  originals.set(oldPath,newPath);tplZip.file(newPath,await file.async('uint8array'));assetCount++;
  const ov=get(srcCt,'Override',NS.ct).find(n=>n.getAttribute('PartName')==='/'+oldPath);
  const extension=(bas(newPath).split('.').pop()||'bin').toLowerCase();
  if(ov)ctOverride(tplCt,newPath,ov.getAttribute('ContentType'));
  else{
   const df=get(srcCt,'Default',NS.ct).find(n=>n.getAttribute('Extension')===extension);
   ctDefault(tplCt,extension,df?.getAttribute('ContentType')||formatMime[extension]||'application/octet-stream');
  }
  const from=xmlRelFile(oldPath);if(srcZip.file(from)){
   const rels=parse(await srcZip.file(from).async('string'));
   for(const rel of relNodes(rels)){
    if(rel.getAttribute('TargetMode')==='External')continue;
    const oldTarget=rel.getAttribute('Target')||'';
    if(!oldTarget||oldTarget[0]==='#')continue;
    const sourceDest=normal(dir(oldPath)+oldTarget);
    const copied=await copyPart(sourceDest);
    if(copied)rel.setAttribute('Target',relpath(dir(newPath),copied));
   }
   tplZip.file(xmlRelFile(newPath),serialize(rels));
  }
  return newPath;
 }
 const referencedIds=new Set();
 for(const n of [tDoc.documentElement])for(const z of Array.from(n.getElementsByTagName('*'))){
  if(z.hasAttributeNS(NS.r,'id'))referencedIds.add(z.getAttributeNS(NS.r,'id'));
  if(z.hasAttributeNS(NS.r,'embed'))referencedIds.add(z.getAttributeNS(NS.r,'embed'));
  if(z.hasAttributeNS(NS.r,'link'))referencedIds.add(z.getAttributeNS(NS.r,'link'));
 }
 // Ignore template document's preexisting IDs: source nodes carry IDs, but they may overlap!
 // Caller marks imported nodes with data-scope via an explicit list of appended node references.
 async function mappingFor(ids){
  for(const srcRel of relNodes(sourceRels)){
   const old=srcRel.getAttribute('Id');if(!ids.has(old))continue;
   const typ=(srcRel.getAttribute('Type')||'').split('/').pop();if(ignoredDocTypes.has(typ))continue;
   const clone=templateRels.createElementNS(NS.rel,'Relationship');
   const newid=rIdDoc(templateRels);clone.setAttribute('Id',newid);clone.setAttribute('Type',srcRel.getAttribute('Type'));
   if(srcRel.getAttribute('TargetMode')==='External'){
    clone.setAttribute('TargetMode','External');clone.setAttribute('Target',srcRel.getAttribute('Target'));
   }else{
    const oldPart=normal('word/'+srcRel.getAttribute('Target'));
    const dest=await copyPart(oldPart);
    if(!dest){warnings.push('Aset yang dirujuk tidak ada: '+oldPart);continue;}
    clone.setAttribute('Target',relpath('word/',dest));
   }
   templateRels.documentElement.appendChild(clone);idMap.set(old,newid);
  }
 }
 await mappingFor(notes.sourceRelationshipIds);
 for(const node of notes.inserted){for(const z of Array.from(node.getElementsByTagName('*'))){
  for(const attName of ['id','embed','link']){if(z.hasAttributeNS(NS.r,attName)){
   const old=z.getAttributeNS(NS.r,attName);if(idMap.has(old))z.setAttributeNS(NS.r,'r:'+attName,idMap.get(old));
  }}
 }}
 // Source notes are imported into the STANDARD Word note parts so the native Word footnotes survive.
 for(const kind of ['footnotes','endnotes']){
  const part='word/'+kind+'.xml';if(!srcZip.file(part))continue;
  const original=await srcZip.file(part).async('string');
  // Templates are cut before their sample body; source notes become authoritative.
  tplZip.file(part,original);
  const srcO=get(srcCt,'Override',NS.ct).find(n=>n.getAttribute('PartName')==='/'+part);
  ctOverride(tplCt,part,srcO?.getAttribute('ContentType')||'application/vnd.openxmlformats-officedocument.wordprocessingml.'+kind+'+xml');
  let tRel=relNodes(templateRels).find(x=>(x.getAttribute('Type')||'').endsWith('/'+kind));
  if(!tRel){tRel=templateRels.createElementNS(NS.rel,'Relationship');tRel.setAttribute('Id',rIdDoc(templateRels));tRel.setAttribute('Type','http://schemas.openxmlformats.org/officeDocument/2006/relationships/'+kind);tRel.setAttribute('Target',kind+'.xml');templateRels.documentElement.appendChild(tRel)}
  const noteRels=xmlRelFile(part);
  if(srcZip.file(noteRels)){
   const rels=parse(await srcZip.file(noteRels).async('string'));
   for(const rel of relNodes(rels)){
    if(rel.getAttribute('TargetMode')==='External')continue;
    const oldPart=normal('word/'+(rel.getAttribute('Target')||''));
    const dest=await copyPart(oldPart);
    if(dest)rel.setAttribute('Target',relpath('word/',dest));
   }
   tplZip.file(noteRels,serialize(rels));
  }else tplZip.remove(noteRels);
 }
 tplZip.file(DOCRELPATH,serialize(templateRels));tplZip.file(CTYPES,serialize(tplCt));
 return {assets:assetCount,relationships:idMap.size};
}
function sourceRefIds(nodes){const ids=new Set();for(const n of nodes){for(const z of [n,...Array.from(n.getElementsByTagName('*'))])for(const a of ['id','embed','link'])if(z.hasAttributeNS?.(NS.r,a))ids.add(z.getAttributeNS(NS.r,a));}return ids;}
function editHeaderParagraph(p,value){
 // Replace only textual running-header content; keep PAGE fields / cached digits.
 const nodes=get(p,'t');const candidates=nodes.filter(x=>{
  const r=x.parentNode;
  return /[A-Za-zÀ-ÿ]/.test(x.textContent||'') && !(r&&get(r,'instrText').length);
 });
 if(!candidates.length)return false;
 // In a mixed header such as "2   Author Surnames" choose the longest
 // human-readable string and leave the PAGE field intact.
 candidates.sort((a,b)=>(b.textContent||'').length-(a.textContent||'').length);
 const target=candidates[0];target.textContent=value;target.setAttributeNS(NS.xml,'xml:space','preserve');
 for(const t of candidates.slice(1))t.textContent='';
 return true;
}
function headerTextMatchesRule(s,key,rule){
 if(!s||s.length>260)return false;
 const blacklist=/^(?:DOI|ISSN|Vol\.|Journal Homepage|E-ISSN|P-ISSN)/i;
 if(blacklist.test(s))return false;
 if(key==='author'||key==='authorShort')return /author|penulis|nama lengkap|indah susilowati|dewi astin|name|hazim|sulistowati|sulisowati|mohammad/i.test(s);
 if(key==='shortTitle')return /title|judul|historical and religious|article tittle|article title|persepsi|artikel/i.test(s);
 return true;
}
async function editJournalHeader(tplZip,templateDoc,info,rule,warnings){
 const rules=rule.header||{},sections=get(templateDoc,'sectPr'),relsFile=tplZip.file(DOCRELPATH);
 if(!sections.length||!Object.keys(rules).length||!relsFile)return 0;
 const rels=parse(await relsFile.async('string'));let count=0;const done=new Set();
 const valueFor=(key)=>key==='shortTitle'?shortTitle(info.title,rule.headerMaxChars||52):
   key==='author'?displayAuthor(info.author):
   key==='authorShort'?(displayAuthor(info.author)+(info.title?' — '+shortTitle(info.title,30):'')):
   info[key]||'';
 for(const sec of sections)for(const header of children(sec,'headerReference')){
  const type=aval(header,'type')||'default',key=rules[type];if(!key)continue;
  const rid=aval(header,'id',NS.r),rel=relNodes(rels).find(x=>x.getAttribute('Id')===rid);if(!rel)continue;
  const part=normal('word/'+rel.getAttribute('Target'));if(!tplZip.file(part)||done.has(part))continue;
  done.add(part);
  const h=parse(await tplZip.file(part).async('string'));
  const eligible=get(h,'p').filter(p=>headerTextMatchesRule(textOf(p),key,rule));
  // Do not touch the journal's masthead merely because it contains 'Title'.
  const p=eligible.find(p=>textOf(p).length>5&& !/\b(?:Journal Homepage|ISSN|DOI)\b/i.test(textOf(p)));
  if(!p){warnings.push('Header '+rule.id+'/'+type+' tidak dapat dipetakan aman; periksa manual.');continue;}
  const content=valueFor(key);
  if(content&&editHeaderParagraph(p,content)){tplZip.file(part,serialize(h));count++;}
 }
 return count;
}
function safeEditorDates(tplZip,tDoc,rule,meta){
 // Scrub the author-guideline demonstration dates (NEVER copy them to a new student article).
 function cleanDate(p){
  const s=textOf(p).replace(/\s+/g,' ').trim();
  if(!s||s.length>220)return false;
  if(!/(?:Received|Accepted|Revised|Published|Submitted|Terkirim|Diterima|Available online)\b/i.test(s))return false;
  if(!/(?:\b(?:19|20)\d{2}\b|\bDD MM YYYY\b|\bmm dd, yyyy\b|\byyyy-mm-dd\b|[.…._]{3,})/i.test(s))return false;
  // Do not alter a descriptive paragraph, table metadata description or the journal name.
  if(!/^(?:Article History|Article history|Received|Accepted|Published|Revised|Submitted|Terkirim|Diterima|Available online)/i.test(s))return false;
  const keys=[['Received','received'],['Submitted','received'],['Revised','revised'],['Accepted','accepted'],['Published','published'],['Terkirim','received'],['Diterima','accepted'],['Available Online','published']];
  let fixed=s;
  for(const [label,k] of keys){
   const pat=new RegExp('\\b'+label.replace(' ','\\s+')+'\\s*:?\\s*(?:\\d{1,2}\\s+[A-Za-z]+\\s+\\d{4}|\\d{1,2}\\s+[A-Za-z]+\\s*,?\\s*\\d{4}|\\d{4}-\\d{2}-\\d{2}|DD MM YYYY|mm dd, yyyy|yyyy-mm-dd|[.…._-]{3,})','gi');
   fixed=fixed.replace(pat,label+': '+(meta?.[k]||'—'));
  }
  if(fixed!==s){replaceText(p,fixed);return true;}
  return false;
 }
 for(const p of get(getBody(tDoc),'p'))cleanDate(p);
 return cleanDate;
}
async function safeguardHeaderFooter(tplZip,rule,meta){
 let changed=0;
 const dateReplacement='Received: '+(meta.received||'—')+', Accepted: '+(meta.accepted||'—')+', Published: '+(meta.published||'—');
 for(const path of Object.keys(tplZip.files).filter(k=>/^word\/(?:header|footer)[^/]*\.xml$/.test(k))){
  const d=parse(await tplZip.file(path).async('string'));let edited=false;
  // Some Word 2010 templates split a single date across many runs, including
  // a legacy WordArt fallback. Replace ALL representations, not just one w:t.
  for(const paragraph of get(d,'p')){
   const s=textOf(paragraph);
   if(rule.id==='karsa' && s.includes('Received:')&&s.includes('Accepted')&&s.includes('Published')){
    // Top paragraph can wrap the inner textboxes: leave nested p elements alone
    // and patch the innermost actual text-bearing p only.
    if(get(paragraph,'p').length)continue;
    const ts=get(paragraph,'t');if(!ts.length)continue;
    ts[0].textContent=dateReplacement;ts[0].setAttributeNS(NS.xml,'xml:space','preserve');
    for(const t of ts.slice(1))t.textContent='';edited=true;
   }
   if(rule.id==='sociale' && /Vol\.\s*99.*Month\s*2099/i.test(s)){
    // Keep "SOCIALE" and the paragraph style but clear sample volume/year.
    const ts=get(paragraph,'t');if(!ts.length)continue;
    ts[0].textContent='SOCIALE, Vol. —, No. —, Tahun —';
    for(const t of ts.slice(1))t.textContent='';edited=true;
   }
  }
  for(const t of get(d,'t')){
   const s=t.textContent||'';
   if(rule.id==='sociale'&&/2099|paperID/.test(s)){
    t.textContent=s.replace(/Vol\.\s*99,\s*No\.\s*1,\s*Month\s*2099,\s*pp\.\s*1~1x/gi,'Vol. —, No. —, Tahun —, pp. —').replace(/Vol\.\s*99,\s*No\.\s*1,\s*Month\s*2099/gi,'Vol. —, No. —, Tahun —').replace(/10\.19105\/sociale\.v2i1\.paperID/gi,'—');edited=true;
   }
  }
  if(edited){tplZip.file(path,serialize(d));changed++;}
 }
 return changed;
}
function sourceStylesUsed(nodes){const styles=new Set();for(const n of nodes)for(const z of get(n,'pStyle'))if(aval(z,'val'))styles.add(aval(z,'val'));return [...styles];}
async function build(articleBuffer,templateBuffer,rule,metadata={}){
 if(typeof global.JSZip==='undefined')throw Error('JSZip tidak dimuat; pastikan folder lib/ ikut diunggah ke GitHub.');
 const [srcZip,tplZip]=await Promise.all([JSZip.loadAsync(articleBuffer),JSZip.loadAsync(templateBuffer)]);
 if(!srcZip.file('word/document.xml')||!tplZip.file('word/document.xml'))throw Error('File Word (.docx) tidak valid.');
 const [srcXml,tplXml]=await Promise.all([srcZip.file('word/document.xml').async('string'),tplZip.file('word/document.xml').async('string')]);
 const srcDoc=parse(srcXml),tplDoc=parse(tplXml),warnings=[];
 const extracted=fromSource(srcDoc),info=unifyMeta(extracted,metadata);
 if(!info.title)throw Error('Judul artikel tidak terdeteksi. Isi Judul manual pada formulir.');
 const headingP=bodyEls(tplDoc)[Number(rule.cut)];if(!headingP)throw Error('Letak paragraf isi template berubah.');
 const prototype=findPrototype(tplDoc,rule);
 const actualRule={...rule,pageWidthTwips:availablePageWidth(tplDoc)};
 const frontCount=replaceFront(tplDoc,rule,info,metadata,warnings);
 if(rule.id==='nuansa')fitNuansaLongTitle(tplDoc,info,warnings);
 safeEditorDates(tplZip,tplDoc,rule,metadata);
 const imported=appendMain(tplDoc,actualRule,extracted.content,prototype,warnings);
 const body=bodyEls(tplDoc);
 // newly inserted source elements appear immediately after the original preface
 const appended=body.slice(Number(rule.cut),Number(rule.cut)+imported);
 const deps=await importSourceParts(srcZip,tplZip,tplDoc,srcDoc,{inserted:appended,sourceRelationshipIds:sourceRefIds(appended)},warnings);
 const headers=await editJournalHeader(tplZip,tplDoc,info,rule,warnings);
 const footerFixes=await safeguardHeaderFooter(tplZip,rule,metadata);
 tplZip.file('word/document.xml',serialize(tplDoc));
 const blob=await tplZip.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:6}});
 const absEnWords=sliceCount(info.abstractEn),absIdWords=sliceCount(info.abstractId);
 if(rule.abstractWords){for(const [lang,n] of [['Inggris',absEnWords],['Indonesia',absIdWords]]){
  if(n && (n<rule.abstractWords[0]||n>rule.abstractWords[1]))warnings.push('Abstrak '+lang+' '+n+' kata, pedoman '+rule.abstractWords[0]+'–'+rule.abstractWords[1]+' kata.');
 }}
 if(rule.titleMax&&sliceCount(info.title)>rule.titleMax)warnings.push('Judul '+sliceCount(info.title)+' kata; batas yang tercatat untuk '+rule.id+' '+rule.titleMax+' kata.');
 if(!info.author)warnings.push('Nama penulis belum terdeteksi; isi manual pada metadata.');
 if(!info.affiliation)warnings.push('Afiliasi belum terdeteksi; isi manual pada metadata.');
 if(!info.abstractEn)warnings.push('Tidak ada Abstract Inggris di sumber; tidak membuat terjemahan otomatis.');
 if(!info.abstractId)warnings.push('Tidak ada Abstrak Indonesia di sumber; tidak membuat terjemahan otomatis.');
 if(/mabny/.test(rule.id))warnings.push('MABNY membutuhkan abstrak terstruktur; periksa Purpose/Methodology/Findings/Implications/Originality.');
 warnings.push('Volume/nomor/tahun/DOI contoh dari template mungkin belum diterbitkan; editor wajib memverifikasi dan melengkapinya sebelum pengiriman.');
 if(headers<Object.keys(rule.header||{}).length)warnings.push('Sebagian running header tidak terpetakan secara otomatis; cek semua halaman di Word.');
 return {blob,details:{journal:rule.id,templateBased:true,frontFields:frontCount,contentBlocks:imported,sourceTables:extracted.tables,sourceFootnotes:extracted.hasNotes,sourceAssets:deps.assets,sourceRelationships:deps.relationships,headersModified:headers,footerFixes,warnings:[...new Set(warnings)],extracted:{title:info.title,author:info.author,affiliation:info.affiliation,email:info.email,abstractEnWords:absEnWords,abstractIdWords:absIdWords}}};
}
global.ZainDocxEngine={build};
})(typeof window!=='undefined'?window:globalThis);
