/* ZAIN.NET — DOCX format dasar di browser, tanpa server.
 * Source document remains the base to preserve original body, tables, images,
 * text annotations, relationships and native Word footnotes.
 * This is NOT an exact implementation of publisher journal typesetting.
 */
(function(global){
'use strict';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL='http://schemas.openxmlformats.org/package/2006/relationships';
const CT='http://schemas.openxmlformats.org/package/2006/content-types';
function parse(x){const doc=new DOMParser().parseFromString(x,'application/xml');if(doc.getElementsByTagName('parsererror').length)throw Error('XML DOCX tidak valid');return doc}
function serial(x){return new XMLSerializer().serializeToString(x)}
function get(e,n,ns=W){return Array.from(e.getElementsByTagNameNS(ns,n))}
function own(e,n,ns=W){return Array.from(e.childNodes).filter(x=>x.nodeType===1&&x.localName===n&&x.namespaceURI===ns)}
function first(e,n){return own(e,n)[0]||null}
function el(doc,n){return doc.createElementNS(W,'w:'+n)}
function val(node,name,ns=W){return node&&node.getAttributeNS(ns,name)}
function setAttr(node,name,value,ns=W,prefix='w'){node.setAttributeNS(ns,prefix+':'+name,String(value))}
function ensure(e,name,before){let z=first(e,name);if(!z){z=el(e.ownerDocument,name);const candidate=before?first(e,before):null;if(candidate)e.insertBefore(z,candidate);else e.appendChild(z)}return z}
function replaceOwned(parent,name,source){for(const n of own(parent,name))parent.removeChild(n);if(source)parent.appendChild(parent.ownerDocument.importNode(source,true))}
function paraText(p){return get(p,'t').map(x=>x.textContent).join('').trim()}
function setChildValue(parent,name,value){const c=ensure(parent,name);setAttr(c,'val',value);return c}
function normalizeStyle(sourceXml,profile,flags){
const doc=parse(sourceXml), body=get(doc,'body')[0]; if(!body)throw Error('Word body tidak ditemukan');
const paragraphs=get(body,'p').filter(p=>!p.closest||true); let inAbstract=false;let titleFound=false;let authorFound=false;let updated=0;
for(const p of paragraphs){
 const text=paraText(p);if(!text)continue;
 // Never rewrite paragraphs inside a table: these may be layout tables or figures.
 let q=p.parentNode;let table=false;while(q&&q!==body){if(q.nodeType===1 && q.namespaceURI===W && q.localName==='tbl'){table=true;break}q=q.parentNode}
 if(table)continue;
 const raw=text.replace(/\s+/g,' ').trim(), upper=raw.toLocaleUpperCase('id');
 const abstractHeading=/^(?:ABSTRAK|ABSTRACT)(?:\s*[:.]?)$/i.test(raw);
 const keywords=/^(?:KATA\s*KUNCI|KEYWORDS?)(?:\s*[:.]?)/i.test(raw);
 const sectionHeading=/^(?:(?:[A-H]|\d{1,2})[.)]\s*)?(?:PENDAHULUAN|INTRODUCTION|METODE(?:\s+PENELITIAN)?|METHODS?|HASIL(?:\s+PENELITIAN)?|RESULTS?(?:\s+AND\s+DISCUSSION)?|PEMBAHASAN|DISCUSSION|KESIMPULAN(?:\s+DAN\s+SARAN)?|CONCLUSIONS?(?:\s+AND\s+RECOMMENDATIONS?)?|DAFTAR\s+PUSTAKA|REFERENCES|TINJAUAN\s+PUSTAKA|LITERATURE\s+REVIEW)$/i.test(raw);
 let role='body';
 if(!titleFound && raw.length>14 && raw.length<420 && !abstractHeading && !keywords){role='title';titleFound=true}
 else if(titleFound&&!authorFound&&!abstractHeading&&raw.length<180 && !/^(?:program studi|fakultas|universitas|email|e-mail|afiliasi|affiliation|department)/i.test(raw)){role='author';authorFound=true}
 if(abstractHeading){role='abstractTitle';inAbstract=true}
 else if(keywords){role='keywords';inAbstract=false}
 else if(sectionHeading){role='heading';inAbstract=false}
 else if(inAbstract && role==='body'){role='abstract'}
 const pp=ensure(p,'pPr','r');
 // Avoid overriding manually formatted bibliography/list paragraph indents.
 if(role==='heading' || role==='abstractTitle' || role==='title' || role==='author'){
  setChildValue(pp,'jc', role==='heading'?'left':'center');
  if(role==='heading'){setChildValue(pp,'keepNext','1');}
  if(role==='title'){
   const s=ensure(pp,'spacing');setAttr(s,'before','80');setAttr(s,'after','120');
  }
 }else if(role==='body' || role==='abstract'){
  if(flags.justify)setChildValue(pp,'jc','both');
 }
 if(role==='body' || role==='abstract'){
  const spacing=ensure(pp,'spacing');
  setAttr(spacing,'line',String(Math.round((profile.line||1)*240)));
  setAttr(spacing,'lineRule','auto');
  if(flags.indent && role==='body'){
   const indent=ensure(pp,'ind');
   if(!indent.hasAttributeNS(W,'firstLine')&&!indent.hasAttributeNS(W,'hanging'))setAttr(indent,'firstLine',String(profile.indentTwips||560));
  }
 }
 const size=role==='title'?profile.titleSize:(role==='author'?profile.authorSize:(role==='abstract'||role==='abstractTitle'||role==='keywords'?profile.abstractSize:role==='heading'?profile.headingSize:profile.bodySize));
 for(const run of own(p,'r')){
  if(!get(run,'t').length)continue;
  let rpr=first(run,'rPr');if(!rpr){rpr=el(doc,'rPr');run.insertBefore(rpr,run.firstChild)}
  const fonts=ensure(rpr,'rFonts');
  for(const a of ['ascii','hAnsi','cs','eastAsia'])setAttr(fonts,a,profile.font);
  setChildValue(rpr,'sz',Math.round(size*2));setChildValue(rpr,'szCs',Math.round(size*2));
  if(role==='title'||role==='abstractTitle'||role==='heading')ensure(rpr,'b');
  updated++;
 }
}
return {xml:serial(doc),paragraphs:paragraphs.length,runs:updated,doc};
}
function sectionParts(doc){return get(doc,'sectPr')}
function transferPageLayout(sourceDoc,templateDoc){
 const tplSection=sectionParts(templateDoc)[0],sections=sectionParts(sourceDoc);
 if(!tplSection)throw Error('Ukuran halaman template jurnal tidak ditemukan');
 if(sections.length===0){const body=get(sourceDoc,'body')[0];if(!body)throw Error('DOCX tidak memiliki body');const s=el(sourceDoc,'sectPr');body.appendChild(s);sections.push(s)}
 for(const sec of sections){
  for(const n of ['pgSz','pgMar','cols','docGrid']){
   const existing=first(sec,n),incoming=first(tplSection,n);
   if(incoming){const next=sourceDoc.importNode(incoming,true);if(existing)sec.replaceChild(next,existing);else sec.appendChild(next)}
  }
 }
 return {sections:sections.length,templateSection:tplSection,sectionsDest:sections};
}
function pkgRelDoc(xml){return xml?parse(xml):parse('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="'+REL+'"/>')}
function relationshipInfo(rels,id){return get(rels,'Relationship',REL).find(n=>n.getAttribute('Id')===id)||null}
function appendRel(doc,type,target){const ids=new Set(get(doc,'Relationship',REL).map(n=>n.getAttribute('Id')));let i=1;while(ids.has('rIdZain'+i))i++;
 const v='rIdZain'+i,r=doc.createElementNS(REL,'Relationship');r.setAttribute('Id',v);r.setAttribute('Type','http://schemas.openxmlformats.org/officeDocument/2006/relationships/'+type);r.setAttribute('Target',target);doc.documentElement.appendChild(r);return v}
function contentTypeOverride(doc,part,ctype){const elems=get(doc,'Override',CT);if(elems.some(n=>n.getAttribute('PartName')==='/'+part))return;const x=doc.createElementNS(CT,'Override');x.setAttribute('PartName','/'+part);x.setAttribute('ContentType',ctype);doc.documentElement.appendChild(x)}
function contentTypeDefault(doc,ext,mime){const elems=get(doc,'Default',CT);if(elems.some(n=>n.getAttribute('Extension')===ext))return;const x=doc.createElementNS(CT,'Default');x.setAttribute('Extension',ext);x.setAttribute('ContentType',mime);doc.documentElement.appendChild(x)}
function fileExt(f){return (f.match(/\.([a-zA-Z0-9]+)$/)||[])[1]?.toLowerCase()||'bin'}
function normalized(path){const bits=[];for(const p of path.split('/')){if(p==='..')bits.pop();else if(p&&p!=='.')bits.push(p)}return bits.join('/')}
function dirname(p){return p.slice(0,p.lastIndexOf('/')+1)}
function relativePath(fromDir,toFile){const from=fromDir.split('/').filter(Boolean),to=toFile.split('/').filter(Boolean);while(from.length&&to.length&&from[0]===to[0]){from.shift();to.shift()}return '../'.repeat(from.length)+to.join('/')}
function mime(ext){return {png:'image/png',jpg:'image/jpeg',jpeg:'image/jpeg',gif:'image/gif',bmp:'image/bmp',svg:'image/svg+xml',emf:'image/x-emf',wmf:'image/x-wmf',tif:'image/tiff',tiff:'image/tiff',webp:'image/webp'}[ext]||'application/octet-stream'}
async function cloneHeaderFooter(tplZip,outZip,templateRels,sourceRels,sourceCT,templateCT,sectionsDest,templSections,opts){
 let imported=0,images=0;let serialNum=0;const byTarget=new Map(), seen=new Map();
 // Import only header/footer definitions referenced by template sections. Selected section is authoritative.
 const referenceNodes=[];
 for(const s of (opts.headerFooter?templSections:[])){for(const x of Array.from(s.childNodes)){if(x.nodeType===1&&x.namespaceURI===W&&['headerReference','footerReference'].includes(x.localName)){
  const key=x.localName+'|'+(val(x,'type')||'default');if(!referenceNodes.some(o=>o.key===key))referenceNodes.push({key,elem:x});
 }}if(referenceNodes.length)break}
 const templateHasTitle=templSections.some(s=>first(s,'titlePg'));
 async function copyPart(oldFile,newFile){
  if(seen.has(oldFile))return seen.get(oldFile);seen.set(oldFile,newFile);
  const source=tplZip.file(oldFile);if(!source)throw Error('Bagian '+oldFile+' tidak ditemukan dalam template');
  outZip.file(newFile,await source.async('uint8array'));
  const oldRels=dirname(oldFile)+'_rels/'+oldFile.split('/').pop()+'.rels';
  if(tplZip.file(oldRels)){
   const d=parse(await tplZip.file(oldRels).async('string'));
   for(const rel of get(d,'Relationship',REL)){
    if(rel.getAttribute('TargetMode')==='External')continue;
    const oldTarget=rel.getAttribute('Target')||'';if(!oldTarget||oldTarget.startsWith('#'))continue;
    const absolute=normalized(dirname(oldFile)+oldTarget);
    if(!tplZip.file(absolute))continue;
    const ext=fileExt(absolute);
    const newDest=seen.get(absolute) || ('word/zain-assets/'+(++serialNum)+'-'+(absolute.split('/').pop().replace(/[^A-Za-z0-9_.-]/g,'_')));
    await copyPart(absolute,newDest);
    // relationship Target is relative to the current part directory
    const newTarget=relativePath(dirname(newFile),newDest);
    rel.setAttribute('Target',newTarget);
    const originalOv=get(templateCT,'Override',CT).find(n=>n.getAttribute('PartName')==='/'+absolute);
    if(originalOv)contentTypeOverride(sourceCT,newDest,originalOv.getAttribute('ContentType'));
    else contentTypeDefault(sourceCT,ext,mime(ext));
    images++;
   }
   outZip.file(dirname(newFile)+'_rels/'+newFile.split('/').pop()+'.rels',serial(d));
  }
  return newFile;
 }
 for(const entry of referenceNodes){
  const type=entry.elem.localName==='headerReference'?'header':'footer';
  const relId=val(entry.elem,'id',R);const relation=relationshipInfo(templateRels,relId);
  if(!relation)continue;
  const oldPath=normalized('word/'+(relation.getAttribute('Target')||''));
  if(!tplZip.file(oldPath))continue;
  const newName=type+'Zain'+(++imported)+'.xml',newPath='word/'+newName;
  await copyPart(oldPath,newPath);
  const newId=appendRel(sourceRels,type,newName);
  byTarget.set(entry.key,{type:entry.elem.localName,wordType:val(entry.elem,'type')||'default',id:newId});
  const srcOv=get(templateCT,'Override',CT).find(n=>n.getAttribute('PartName')==='/'+oldPath);
  contentTypeOverride(sourceCT,newPath,srcOv?.getAttribute('ContentType')||'application/vnd.openxmlformats-officedocument.wordprocessingml.'+type+'+xml');
  if(type==='footer'){
   const footer=parse(await outZip.file(newPath).async('string'));
   scrubEditorialDates(footer);
   const containsPage=(serial(footer).match(/\bPAGE\b/gi)||[]).length>0;
   if(opts.pageNumbers && !containsPage)appendPageNumber(footer);
   outZip.file(newPath,serial(footer));
  }
 }
 if(opts.pageNumbers && ![...byTarget.values()].some(x=>x.type==='footerReference')){
  const doc=parse('<?xml version="1.0" encoding="UTF-8"?><w:ftr xmlns:w="'+W+'"/>');appendPageNumber(doc);
  const newName='footerZainNumbers.xml';outZip.file('word/'+newName,serial(doc));
  const id=appendRel(sourceRels,'footer',newName);
  contentTypeOverride(sourceCT,'word/'+newName,'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml');
  byTarget.set('footerReference|default',{type:'footerReference',wordType:'default',id});
 }
 if(byTarget.size){
  for(const sec of sectionsDest){
   for(const n of Array.from(sec.childNodes))if(n.nodeType===1&&n.namespaceURI===W&&['headerReference','footerReference'].includes(n.localName) && (opts.headerFooter||n.localName==='footerReference'))sec.removeChild(n);
   const anchor=sec.firstChild;
   for(const ref of byTarget.values()){
    const x=el(sec.ownerDocument,ref.type);setAttr(x,'type',ref.wordType);x.setAttributeNS(R,'r:id',ref.id);
    sec.insertBefore(x,anchor);
   }
   if(templateHasTitle){ensure(sec,'titlePg')}
   else{for(const n of own(sec,'titlePg'))sec.removeChild(n)}
  }
 }
 return {count:imported,extraAssets:images,even:referenceNodes.some(x=>x.key.endsWith('|even'))};
}
function scrubEditorialDates(footer){
 for(const p of get(footer,'p')){const line=paraText(p);
 if(/Received\s*:/i.test(line)&&/Accepted\s*:/i.test(line) && /(?:\b(?:19|20)\d{2}\b|\d{1,2}\s+[A-Za-z]+\s+\d{4})/.test(line)){
 const parts=get(p,'t'); if(parts.length){parts[0].textContent='Received: —, Accepted: —, Published: —';
 for(const x of parts.slice(1))x.textContent='';}
 }}
}
function appendPageNumber(footer){
 const d=footer, root=d.documentElement,p=el(d,'p'),ppr=el(d,'pPr'),jc=el(d,'jc');setAttr(jc,'val','center');ppr.appendChild(jc);p.appendChild(ppr);
 const r1=el(d,'r'),begin=el(d,'fldChar');setAttr(begin,'fldCharType','begin');r1.appendChild(begin);p.appendChild(r1);
 const r2=el(d,'r'),instr=el(d,'instrText');instr.setAttribute('xml:space','preserve');instr.textContent=' PAGE ';r2.appendChild(instr);p.appendChild(r2);
 const r3=el(d,'r'),separate=el(d,'fldChar');setAttr(separate,'fldCharType','separate');r3.appendChild(separate);p.appendChild(r3);
 const r4=el(d,'r'),t=el(d,'t');t.textContent='1';r4.appendChild(t);p.appendChild(r4);
 const r5=el(d,'r'),end=el(d,'fldChar');setAttr(end,'fldCharType','end');r5.appendChild(end);p.appendChild(r5);
 root.appendChild(p);
}
function setEvenOdd(settingsXml,flag){if(!settingsXml)return settingsXml;const d=parse(settingsXml),root=d.documentElement,existing=get(d,'evenAndOddHeaders')[0];if(flag&&!existing)root.appendChild(el(d,'evenAndOddHeaders'));return serial(d)}
async function build(article,template,profile,options={}){
 if(typeof JSZip==='undefined')throw Error('Pustaka JSZip belum dimuat');
 const opts={justify:true,indent:true,headerFooter:false,pageNumbers:false,...options};
 const [src,tpl]=await Promise.all([JSZip.loadAsync(article),JSZip.loadAsync(template)]);
 const srcDocFile=src.file('word/document.xml'),tplDocFile=tpl.file('word/document.xml');
 if(!srcDocFile||!tplDocFile)throw Error('Dokumen bukan DOCX yang valid. File .doc tidak didukung untuk upload artikel.');
 const sourceXml=await srcDocFile.async('string'),tplXml=await tplDocFile.async('string');
 const formatted=normalizeStyle(sourceXml,profile,opts),sourceDoc=formatted.doc,templateDoc=parse(tplXml);
 const layout=transferPageLayout(sourceDoc,templateDoc);
 let extras={count:0,extraAssets:0};
 if(opts.headerFooter||opts.pageNumbers){
  const sourceRelsPath='word/_rels/document.xml.rels';
  const sourceRels=pkgRelDoc(src.file(sourceRelsPath)?await src.file(sourceRelsPath).async('string'):null);
  const templateRels=pkgRelDoc(tpl.file(sourceRelsPath)?await tpl.file(sourceRelsPath).async('string'):null);
  const ctPath='[Content_Types].xml';const sourceCT=parse(await src.file(ctPath).async('string'));
  const templateCT=parse(await tpl.file(ctPath).async('string'));
  extras=await cloneHeaderFooter(tpl,src,templateRels,sourceRels,sourceCT,templateCT,layout.sectionsDest,
   sectionParts(templateDoc),opts);
  src.file(sourceRelsPath,serial(sourceRels));src.file(ctPath,serial(sourceCT));
  const settingsPath='word/settings.xml';if(extras.even&&src.file(settingsPath))src.file(settingsPath,setEvenOdd(await src.file(settingsPath).async('string'),true));
 }
 src.file('word/document.xml',serial(sourceDoc));
 const blob=await src.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:6}});
 return {blob,details:{paragraphs:formatted.paragraphs,formattedRuns:formatted.runs,sections:layout.sections,headersFooters:extras.count,assets:extras.extraAssets,note:'Hasil format dasar; kesesuaian identik dengan template resmi belum dijamin.'}};
}
global.ZainDocxEngine={build};
})(typeof window!=='undefined'?window:globalThis);
