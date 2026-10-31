/* ZAIN.NET local-only DOCX editor — OOXML package-aware formatting.
 * All source files remain in memory within the browser; no document upload/API.
 * LGPL? no, ZAIN code MIT. JSZip bundled separately under MIT.
 */
(function () {
  'use strict';
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const P = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
  const XML = 'http://www.w3.org/XML/1998/namespace';
  const IMG = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const E = (ns, tag, doc) => doc.createElementNS(ns, tag);
  const w = (name, doc) => E(W, 'w:' + name, doc);
  const attr = (node, ns, name) => node && node.getAttributeNS(ns, name);
  const put = (node, ns, prefix, name, value) => node.setAttributeNS(ns, prefix + ':' + name, String(value));
  const get = (node, ns, tag) => node ? Array.from(node.childNodes).find(c => c.nodeType === 1 && c.localName === tag && c.namespaceURI === ns) || null : null;
  const gets = (node, ns, tag) => node ? Array.from(node.childNodes).filter(c => c.nodeType === 1 && c.localName === tag && c.namespaceURI === ns) : [];
  const all = (node, ns, tag) => node ? Array.from(node.getElementsByTagNameNS(ns, tag)) : [];
  const normalize = s => (s || '').replace(/\s+/g, ' ').trim();
  const docParaText = p => all(p, W, 't').map(t => t.textContent).join('');
  const bytes = s => new TextEncoder().encode(s);
  function xmlParse(src) {
    const d = new DOMParser().parseFromString(src, 'application/xml');
    if (d.getElementsByTagName('parsererror').length) throw new Error('Struktur XML dalam DOCX tidak valid: '+d.getElementsByTagName('parsererror')[0].textContent.slice(0,120));
    return d;
  }
  function xmlString(doc) { const serialized=new XMLSerializer().serializeToString(doc);return serialized.startsWith('<?xml')?serialized:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'+serialized; }
  async function readPart(zip, path, required = true) {
    const item = zip.file(path);
    if (!item) { if (!required) return null; throw new Error('Berkas Word tidak lengkap: ' + path); }
    return xmlParse(await item.async('string'));
  }
  function setTag(parent, localName, value, ns = W) {
    let el = get(parent, ns, localName);
    if (!el) { el = w(localName, parent.ownerDocument || parent); parent.appendChild(el); }
    if (value !== undefined) put(el, W, 'w', 'val', value);
    return el;
  }
  function setOnOff(parent, localName, enabled) { const el = setTag(parent, localName); put(el, W, 'w', 'val', enabled ? '1' : '0'); }
  function setProp(parent, name, attrs) {
    const node = setTag(parent, name);
    for (const [k, v] of Object.entries(attrs)) put(node, W, 'w', k, v);
    return node;
  }
  function setPPr(p, config = {}) {
    let pPr = get(p, W, 'pPr');
    if (!pPr) { pPr = w('pPr', p.ownerDocument); p.insertBefore(pPr, p.firstChild); }
    const jc = config.align || 'both';
    setProp(pPr, 'jc', {val: jc});
    const spacing = setProp(pPr, 'spacing', {
      before: config.before ?? '0', after: config.after ?? '0', line: String(Math.round((config.line ?? 1.5) * 240)), lineRule: 'auto'
    });
    // w:ind firstLine in twentieths of pt, heading & references override.
    setProp(pPr, 'ind', {left: config.left ?? '0', right: '0', ...(config.hanging ? {hanging:String(config.hanging)} : {firstLine:String(config.first ?? 0)})});
    setOnOff(pPr, 'keepNext', Boolean(config.keepNext));
    setOnOff(pPr, 'keepLines', false);
    setOnOff(pPr, 'pageBreakBefore', false);
    setOnOff(pPr, 'widowControl', true);
    return pPr;
  }
  function setRunStyle(r, options = {}) {
    let rPr = get(r, W, 'rPr');
    if (!rPr) { rPr = w('rPr', r.ownerDocument); r.insertBefore(rPr, r.firstChild); }
    const f = options.font || 'Book Antiqua';
    setProp(rPr, 'rFonts', {ascii:f, hAnsi:f, cs:f, eastAsia:f});
    const v = options.pt ?? 12;
    setProp(rPr, 'sz', {val:String(Math.round(v*2))});
    setProp(rPr, 'szCs', {val:String(Math.round(v*2))});
    setProp(rPr, 'color', {val:'000000'});
    if (options.bold !== undefined) setOnOff(rPr, 'b', options.bold);
    if (options.italic !== undefined) setOnOff(rPr, 'i', options.italic);
    if (options.underline !== undefined) setProp(rPr, 'u', {val:options.underline ? 'single':'none'});
  }
  function formatP(p, props, runs) {
    setPPr(p,props);
    // Only change rPr; preserve all text, drawings, captions, footnote references and hyperlinks.
    all(p, W, 'r').forEach(r => setRunStyle(r,runs));
  }
  function setNormalStyle(stylesDoc, font, size, line = 1.5) {
    if (!stylesDoc) return;
    const root = stylesDoc.documentElement;
    const normal = all(root, W, 'style').find(s => attr(s,W,'styleId') === 'Normal');
    if (normal) {
      let rPr=get(normal,W,'rPr'); if(!rPr){rPr=w('rPr',stylesDoc);normal.appendChild(rPr);}
      setProp(rPr,'rFonts',{ascii:font,hAnsi:font,eastAsia:font,cs:font});
      setProp(rPr,'sz',{val:String(size*2)});
      setProp(rPr,'szCs',{val:String(size*2)});
      let pPr=get(normal,W,'pPr');if(!pPr){pPr=w('pPr',stylesDoc);normal.appendChild(pPr);}
      setProp(pPr,'spacing',{before:'0',after:'0',line:String(line*240),lineRule:'auto'});
    }
  }
  function sectionHead(text) {
    const x = normalize(text).replace(/^(?:[A-Z]|\d+|[IVX]+)[.)]\s*/i,'').replace(/\s*\([^)]*paragraphs[^)]*\)\s*$/i,'').replace(/:$/,'').trim().toUpperCase();
    const names = {
      'ABSTRACT':'abstract','ABSTRAK':'abstrak', 'PENDAHULUAN':'intro', 'INTRODUCTION':'intro',
      'METODE PENELITIAN':'method','METODE':'method','METHODS':'method','METHOD':'method',
      'RESEARCH METHODOLOGY':'method','METHODOLOGY':'method',
      'HASIL DAN PEMBAHASAN':'results','HASIL PENELITIAN DAN PEMBAHASAN':'results',
      'RESULTS AND DISCUSSION':'results','RESULT AND DISCUSSION':'results','HASIL PENELITIAN':'results','PEMBAHASAN':'results',
      'KESIMPULAN DAN SARAN':'conclusion','KESIMPULAN':'conclusion','SIMPULAN':'conclusion','CONCLUSION':'conclusion',
      'DAFTAR PUSTAKA':'references','BIBLIOGRAPHY':'references','REFERENCES':'references','REFERENCE':'references',
      'DECLARATION OF GENERATIVE AI AND AI-ASSISTED TECHNOLOGIES':'ai', 'DEKLARASI PENGGUNAAN AI':'ai',
      'ACKNOWLEDGMENT':'ack','ACKNOWLEDGEMENTS':'ack', 'ACKNOWLEDGMENTS':'ack','UCAPAN TERIMA KASIH':'ack'
    };
    return names[x] || null;
  }
  function paragraphIsInside(p, parent) { for(let cur=p.parentNode;cur;cur=cur.parentNode) if(cur === parent) return true; return false; }
  function createParagraph(doc, s, properties = {}, font = {}) {
    const p = w('p',doc), r=w('r',doc),t=w('t',doc);
    t.textContent=s;put(t,XML,'xml','space','preserve');r.appendChild(t);p.appendChild(r);
    formatP(p,properties,font);return p;
  }
  function appendField(doc, p) {
    const f = w('fldSimple',doc);put(f,W,'w','instr','PAGE');
    const r=w('r',doc),t=w('t',doc);t.textContent='1';r.appendChild(t);f.appendChild(r);p.appendChild(f);
  }
  function setTextPreservingRuns(root, needle, replacement) {
    // Cross-run replacement, retains XML drawing elements instead of rewriting paragraph innerHTML.
    if(!needle || needle===replacement) return false;
    let changed=false;
    for(const p of all(root,W,'p')) {
      const texts=all(p,W,'t');if(!texts.length)continue;
      const line=texts.map(n=>n.textContent).join('');let at=line.indexOf(needle);if(at<0)continue;
      // Rewrite one match by character boundaries across individual w:t nodes.
      let end=at+needle.length, offset=0, started=false;
      for(const node of texts) {
        const str=node.textContent, stop=offset+str.length;
        if(stop>at && offset<end) {
          const left=str.slice(0,Math.max(0,at-offset));
          const right=str.slice(Math.max(0,end-offset));
          node.textContent=left+(started?'':replacement)+right;
          started=true;
          put(node,XML,'xml','space','preserve');
        }
        offset=stop;
      }
      changed=true;
    }
    return changed;
  }
  function overwriteParagraphText(p,newText) {
    const nodes=all(p,W,'t');if(!nodes.length){const r=w('r',p.ownerDocument),t=w('t',p.ownerDocument);t.textContent=newText;put(t,XML,'xml','space','preserve');r.appendChild(t);p.appendChild(r);return;}
    nodes.forEach((node,i)=>{node.textContent=i===0?newText:'';put(node,XML,'xml','space','preserve');});
  }
  function replaceSectionText(p) {
    const old=docParaText(p);
    const cleaned=old.replace(/^(?:[A-Z]|\d+|[IVX]+)[.)]\s*/i,'');
    if(old!==cleaned) setTextPreservingRuns(p,old,cleaned);
  }
  function setPageSizeFromTemplate(doc, tpl) {
    const sourceSection=all(tpl,W,'sectPr').at(-1);
    if(!sourceSection)throw new Error('Template tidak memiliki pengaturan halaman.');
    const targets=all(doc,W,'sectPr');
    if(!targets.length)throw new Error('Dokumen sumber tidak memiliki pengaturan section halaman.');
    for(const sect of targets){
      for(const key of ['pgSz','pgMar','pgBorders']){
        const src=get(sourceSection,W,key);if(!src)continue;
        const old=get(sect,W,key);if(old)sect.removeChild(old);
        sect.appendChild(doc.importNode(src,true));
      }
      const cols=get(sect,W,'cols');
      if(cols) { put(cols,W,'w','num','1');for(const child of Array.from(cols.children)) cols.removeChild(child); }
      if(sect===targets[0]) setTag(sect,'titlePg');
    }
    return targets;
  }
  function addEvenOddSettings(doc) { const root=doc.documentElement;setTag(root,'evenAndOddHeaders').setAttributeNS(W,'w:val','1'); }
  function addHeaderRef(sect,type,rid) {
    const old=gets(sect,W,'headerReference').find(n=>attr(n,W,'type')===type);if(old)sect.removeChild(old);
    const el=w('headerReference',sect.ownerDocument);put(el,W,'w','type',type);put(el,R,'r','id',rid);
    sect.insertBefore(el,sect.firstChild);
  }
  function addFooterRef(sect,type,rid) {
    const old=gets(sect,W,'footerReference').find(n=>attr(n,W,'type')===type);if(old)sect.removeChild(old);
    const el=w('footerReference',sect.ownerDocument);put(el,W,'w','type',type);put(el,R,'r','id',rid);
    const headerRefs=gets(sect,W,'headerReference');
    const after=headerRefs.at(-1);if(after)sect.insertBefore(el,after.nextSibling);else sect.insertBefore(el,sect.firstChild);
  }
  function ensureCT(ct, name, type) {
    const root=ct.documentElement;
    const exists=Array.from(root.children).find(n=>n.localName==='Override'&&n.getAttribute('PartName')===('/'+name));
    if(!exists) {const n=ct.createElementNS(CT,'Override');n.setAttribute('PartName','/'+name);n.setAttribute('ContentType',type);root.appendChild(n);}
  }
  function ensureDefaultCT(ct, ext, type) {
    const root=ct.documentElement;
    if(Array.from(root.children).some(n=>n.localName==='Default'&&n.getAttribute('Extension')===ext))return;
    const n=ct.createElementNS(CT,'Default');n.setAttribute('Extension',ext);n.setAttribute('ContentType',type);root.appendChild(n);
  }
  function folder(path) { return path.slice(0,path.lastIndexOf('/')+1); }
  function pathResolve(from, target) {
    const bits = (folder(from)+target).split('/'), out=[];
    for(const bit of bits){if(!bit||bit==='.')continue;if(bit==='..')out.pop();else out.push(bit);}
    return out.join('/');
  }
  function relative(from,to) {
    const a=folder(from).split('/').filter(Boolean),b=to.split('/').filter(Boolean);
    while(a.length&&b.length&&a[0]===b[0]){a.shift();b.shift();}
    return '../'.repeat(a.length)+b.join('/');
  }
  function relPath(path){return folder(path)+'_rels/'+path.split('/').at(-1)+'.rels';}
  function allocateRid(rels){let i=20;const seen=new Set(Array.from(rels.documentElement.children).map(n=>n.getAttribute('Id')));while(seen.has('rId'+i))i++;return 'rId'+i;}
  function addRelationship(rels,id,type,target,external=false){
    const r=rels.createElementNS(P,'Relationship');r.setAttribute('Id',id);r.setAttribute('Type',type);r.setAttribute('Target',target);
    if(external)r.setAttribute('TargetMode','External');rels.documentElement.appendChild(r);
  }
  function emptyRels(){ return xmlParse('<?xml version="1.0"?><Relationships xmlns="'+P+'"/>'); }
  function getRel(rels,rid){return rels && Array.from(rels.documentElement.children).find(n=>n.getAttribute('Id')===rid);}
  function imageContentType(ext){return ({png:'image/png',jpg:'image/jpeg',jpeg:'image/jpeg',gif:'image/gif',bmp:'image/bmp',svg:'image/svg+xml',tif:'image/tiff',tiff:'image/tiff',emf:'image/x-emf',wmf:'image/x-wmf'})[ext.toLowerCase()]||'application/octet-stream';}
  async function transplantStories(outZip, tplZip, sourceDoc, tplDoc, sourceRels, tplRels, ct, meta, warnings, templateId) {
    const srcSection=all(tplDoc,W,'sectPr').at(-1);
    const sourceSections=all(sourceDoc,W,'sectPr');
    let seq=0;
    const cached=new Map();
    async function copyPart(originalPath) {
      if(cached.has(originalPath))return cached.get(originalPath);
      const original=tplZip.file(originalPath);
      if(!original){warnings.push('Aset template tidak tersedia: '+originalPath);return null;}
      seq++;
      const ext=(originalPath.split('.').at(-1)||'xml').toLowerCase();
      const isXml=ext==='xml';
      const base = originalPath.split('/').at(-1).replace(/[^a-z0-9_.-]/gi,'_');
      const dest = isXml ? 'word/zainnet_story_'+seq+'.xml' : 'word/media/zainnet_asset_'+seq+'_'+base;
      cached.set(originalPath,dest);
      if(!isXml){ outZip.file(dest,await original.async('uint8array'));ensureDefaultCT(ct,ext,imageContentType(ext));return dest; }
      const d=await readPart(tplZip,originalPath);
      const relationships=await readPart(tplZip,relPath(originalPath),false);
      const updatedRels=emptyRels();
      if(relationships){
        for(const r of Array.from(relationships.documentElement.children)){
          const rid=r.getAttribute('Id'),target=r.getAttribute('Target'),type=r.getAttribute('Type'),external=r.getAttribute('TargetMode')==='External';
          if(!rid||!target||!type)continue;
          let newTarget=target;
          if(!external) {
            const srcTarget=pathResolve(originalPath,target);
            const destTarget=await copyPart(srcTarget);
            if(!destTarget)continue;
            newTarget=relative(dest,destTarget);
          } else if(/\b10\.19105\//i.test(target)) {
            newTarget= /^https?:\/\//i.test(meta.doi||'') ? meta.doi : (meta.doi && meta.doi!=='—' ? 'https://doi.org/'+meta.doi : 'https://ejournal.uinmadura.ac.id/');
          } else if(/^mailto:/i.test(target) && /address@mail\.ac\.id/i.test(target)) {
            newTarget=meta.correspondence_email?'mailto:'+meta.correspondence_email:'https://ejournal.uinmadura.ac.id/';
          }
          addRelationship(updatedRels,rid,type,newTarget,external);
        }
      }
      // Preserve template artwork, headers and page counts while removing sample author fields.
      const textReplacements={
        'Rahmat Dandi, Pendidikan Karakter dalam Islam':(meta.author||'Penulis')+', '+(meta.running_title||meta.title||'Artikel'),
        'The Title Article Journal …………………….':meta.running_title||meta.title||'Artikel',
        'full name of first author, et al.':meta.running_author||meta.author||'Penulis',
        'DOI: https://doi.org/10.19105/rjpai.v1i1.3002':'DOI: '+(meta.doi||'—'),
        'https://doi.org/10.19105/al-lhkam.v17i2.7229':meta.doi||'—',
        'address@mail.ac.id':meta.correspondence_email||meta.email||'—',
        'Volume 1, Nomor 1, Maret 2020':'Volume '+(meta.volume||'—')+', Nomor '+(meta.issue||'—')+', '+(meta.month||'—')+' '+(meta.year||'—'),
        'Vol. 1, No. 1, Maret 2020':'Vol. '+(meta.volume||'—')+', No. '+(meta.issue||'—')+', '+(meta.month||'—')+' '+(meta.year||'—'),
        'Historical and Religious Site as Tourism Attraction: The Case  ....':meta.running_title||meta.title||'Artikel',
        'Indah  Susilowati ,  Hapsari  Ayu  Kusumawardhani ,  ....':meta.author||'Penulis',
        'Indah Susilowati, Hapsari Ayu Kusumawardhani, ....':meta.author||'Penulis',
        'Historical and Religious Site as Tourism Attraction: The Case ....':meta.running_title||meta.title||'Artikel'
      };
      for(const [a,b] of Object.entries(textReplacements))setTextPreservingRuns(d,a,b);
      // Remove authors/volume/DOI sample metadata instead of leaving fabricated journal records.
      if(originalPath.includes('header')) for(const paragraph of all(d,W,'p')) {
        const txt=normalize(docParaText(paragraph));
        if(!txt)continue;
        if(templateId==='karsa') {
          if(/Indah\s+Susilowati|Hapsari\s+Ayu/i.test(txt)) overwriteParagraphText(paragraph,meta.running_author||meta.author||'Penulis');
          else if(/Historical and Religious Site as Tourism/i.test(txt))overwriteParagraphText(paragraph,meta.running_title||meta.title||'Judul artikel');
        }
        if(templateId==='rabbani') {
          if(/Rahmat Dandi|Pendidikan Karakter dalam Islam/i.test(txt)) overwriteParagraphText(paragraph,(meta.running_author||meta.author||'Penulis')+', '+(meta.running_title||meta.title||'Judul artikel'));
          if(/^(?:Volume|Vol\.)\s*1[,\s]/i.test(txt))overwriteParagraphText(paragraph,'Volume '+(meta.volume||'—')+', Nomor '+(meta.issue||'—')+', '+(meta.month||'—')+' '+(meta.year||'—'));
        }
        if(templateId==='alihkam') {
          if(/full name of first author/i.test(txt))overwriteParagraphText(paragraph,meta.running_author||meta.author||'Penulis');
          else if(/The Title Article Journal/i.test(txt))overwriteParagraphText(paragraph,meta.running_title||meta.title||'Judul artikel');
          else if(/Al-Ihkam: Jurnal Hukum dan Pranata Sosial.*21\s*\(1\)/i.test(txt))overwriteParagraphText(paragraph,'Al-Ihkam: Jurnal Hukum dan Pranata Sosial'+(meta.year?' · '+meta.year:''));
          else if(/^DOI:/i.test(txt))overwriteParagraphText(paragraph,'DOI: '+(meta.doi||'—'));
        }
      }
      outZip.file(dest,xmlString(d));
      if(relationships)outZip.file(relPath(dest),xmlString(updatedRels));
      ensureCT(ct,dest,originalPath.includes('header')?'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml':
        originalPath.includes('footer')?'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml':'application/xml');
      return dest;
    }
    // Only copy visual headers from journal examples. Sample footers contain
    // actual authors, editorial dates and pagination: never publish those as real metadata.
    for(const ref of gets(srcSection,W,'headerReference')) {
      const type=attr(ref,W,'type')||'default',rid=attr(ref,R,'id'),rel=getRel(tplRels,rid);
      if(!rel)continue;
      const originalPath=pathResolve('word/document.xml',rel.getAttribute('Target'));
      const dest=await copyPart(originalPath);if(!dest)continue;
      const newid=allocateRid(sourceRels);
      addRelationship(sourceRels,newid,'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',relative('word/document.xml',dest));
      sourceSections.forEach(s=>addHeaderRef(s,type,newid));
    }
    // Fresh centered PAGE field, in place of sample numbers/dates from journal examples.
    // Three types ensure first/even/ordinary pages all show one real page number.
    for(const type of ['default','first','even']) {
      const dest='word/zainnet_page_footer_'+type+'.xml';
      const d=xmlParse('<?xml version="1.0"?><w:ftr xmlns:w="'+W+'"/>');
      if(type==='first') {
        const text = templateId==='rabbani'
          ? [meta.year && '© '+meta.year+' '+(meta.author||''),
             (meta.received || meta.revised || meta.accepted) && 'Received: '+(meta.received||'—')+' | Revised: '+(meta.revised||'—')+' | Accepted: '+(meta.accepted||'—'),
             meta.pages && 'Pages: '+meta.pages].filter(Boolean)
          :templateId==='alihkam'
           ? [meta.correspondence_email && 'Author correspondence email: '+meta.correspondence_email,
              meta.year && '© '+meta.year+' Al-Ihkam: Jurnal Hukum dan Pranata Sosial'].filter(Boolean)
           : [meta.received && 'Received: '+meta.received,meta.accepted && 'Accepted: '+meta.accepted,meta.published && 'Published: '+meta.published].filter(Boolean);
        for(const line of text)d.documentElement.appendChild(createParagraph(d,line,{align:'center',line:1},{pt:8}));
      }
      const p=createParagraph(d,'',{align:'center',line:1},{pt:10});appendField(d,p);d.documentElement.appendChild(p);
      outZip.file(dest,xmlString(d));ensureCT(ct,dest,'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml');
      const newid=allocateRid(sourceRels);
      addRelationship(sourceRels,newid,'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',relative('word/document.xml',dest));
      sourceSections.forEach(s=>addFooterRef(s,type,newid));
    }
  }
  function updateKarsa(paragraphs, warnings) {
    let state=null;
    paragraphs.forEach((p,i)=>{
      const text=normalize(docParaText(p)),head=sectionHead(text);
      if(!text)return;
      if(i===0)formatP(p,{align:'center',line:1.15,before:160,after:160,keepNext:true},{font:'Times New Roman',pt:13,bold:true});
      else if(head){state=head;formatP(p,{align:'left',line:1.2,before:160,after:80,keepNext:true},{font:'Times New Roman',pt:11,bold:true});}
      else if(!state)formatP(p,{align:'center',line:1.2},{font:'Times New Roman',pt:10});
      else if(state==='abstract'||state==='abstrak')formatP(p,{align:'both',line:1,first:0},{font:'Times New Roman',pt:10});
      else if(state==='references')formatP(p,{align:'left',line:1,first:0,hanging:340,left:340,after:80},{font:'Times New Roman',pt:11});
      else formatP(p,{align:'both',line:1.2,first:380},{font:'Times New Roman',pt:11});
    });
    warnings.push('KARSA: ukuran dan heading disesuaikan; tata letak dua kolom, header jurnal dan sitasi masih perlu pemeriksaan manual.');
  }
  function createRabbaniAbstractTable(doc,ps,meta,warnings) {
    const headings=[];let state=null;
    for(const p of ps){
      const t=normalize(docParaText(p)),h=sectionHead(t);
      if(h){state=h;if(h==='abstract'||h==='abstrak')headings.push(p);continue;}
      if(/^(keywords?|kata\s*kunci)\s*:/i.test(t)){state='keywords';continue;}
      if(state==='abstract'||state==='abstrak')headings.push(p);
    }
    if(!headings.length)return;
    // Keep footnotes/drawings with their paragraphs; avoid reparenting if tables or complex section breaks interrupt flow.
    if(headings.some(p=>all(p,W,'sectPr').length)){warnings.push('Kotak abstrak tidak dibuat karena ada section break di abstrak.');return;}
    const body = get(doc.documentElement,W,'body');
    const first = headings[0];
    if(!first || !paragraphIsInside(first,body)||first.parentNode!==body)return;
    const tbl = w('tbl',doc),tblPr=w('tblPr',doc),tblW=w('tblW',doc);put(tblW,W,'w','type','dxa');put(tblW,W,'w','w','8500');tblPr.appendChild(tblW);tbl.appendChild(tblPr);
    const grid=w('tblGrid',doc);for(const width of [5500,150,2850]){const c=w('gridCol',doc);put(c,W,'w','w',width);grid.appendChild(c);}tbl.appendChild(grid);
    const tr=w('tr',doc);
    function cell(width,shade,borderColor){const tc=w('tc',doc),tcPr=w('tcPr',doc),cw=w('tcW',doc);put(cw,W,'w','type','dxa');put(cw,W,'w','w',width);tcPr.appendChild(cw);
      if(shade){const sh=w('shd',doc);put(sh,W,'w','fill',shade);tcPr.appendChild(sh);}if(borderColor){const borders=w('tcBorders',doc);for(const side of ['top','bottom','left','right']){const e=w(side,doc);put(e,W,'w','val','single');put(e,W,'w','color',borderColor);put(e,W,'w','sz','6');borders.appendChild(e);}tcPr.appendChild(borders);}tc.appendChild(tcPr);return tc;}
    const left=cell('5500','D4F6FF','45A9C4'),middle=cell('150'),right=cell('2850','FFFFFF','00B0F0');
    const anchor=first;body.insertBefore(tbl,anchor);
    for(const p of headings)if(p.parentNode===body)left.appendChild(p);
    if(!gets(left,W,'p').length)left.appendChild(createParagraph(doc,'Abstract',{align:'left',line:1},{pt:11,bold:true}));
    middle.appendChild(createParagraph(doc,'',{align:'left',line:1},{pt:1}));
    right.appendChild(createParagraph(doc,'ARTICLE HISTORY',{align:'left',line:1,after:100},{pt:10,bold:true}));
    for(const k of ['received','revised','accepted'])right.appendChild(createParagraph(doc,k[0].toUpperCase()+k.slice(1)+': '+(meta[k]||'—'),{align:'left',line:1},{pt:10}));
    right.appendChild(createParagraph(doc,'KEYWORDS',{align:'left',line:1,before:100},{pt:10,bold:true}));
    // Move keyword paragraphs into sidebar only if they occur between abstracts and introduction.
    const keywordPs=ps.filter(p=>/^(?:keywords?|kata\s*kunci)\s*:/i.test(normalize(docParaText(p))));
    for(const p of keywordPs){if(p.parentNode===body){right.appendChild(p);formatP(p,{align:'left',line:1},{pt:10});}}
    for(const s of [left,middle,right])tr.appendChild(s);tbl.appendChild(tr);
    warnings.push('Kotak abstrak/sidebar dibuat otomatis; periksa tinggi tabel dan pembagian halaman.');
  }
  function updateRabbani(doc,paragraphs,opts,meta,warnings) {
    const titleCount=opts.titleParagraphs,stateSet=new Set();let state=null;let titleWords=[];
    const abstracts={abstract:[],abstrak:[]};
    paragraphs.forEach((p,i)=>{
      const txt=normalize(docParaText(p)), head=sectionHead(txt);
      if(!txt)return;
      if(i<titleCount){titleWords.push(txt);formatP(p,{align:'center',line:opts.profile==='guidelines'?1:1.5,after:0,before:i?0:140,keepNext:true},{pt:opts.profile==='guidelines'?12:13,bold:true});return;}
      if(head){state=head;stateSet.add(head);replaceSectionText(p);formatP(p,{align:'left',line:head==='references'||head==='abstract'||head==='abstrak'?1:1.5,before:130,keepNext:true},{pt:head==='abstract'||head==='abstrak'?11:12,bold:true});return;}
      if(/^(keywords?|kata\s*kunci)\s*:/i.test(txt)){formatP(p,{align:'left',line:1},{pt:11});return;}
      if(!state){let isAff=/@|universitas|institut|university|\bUIN\b|\bIAIN\b/i.test(txt);formatP(p,{align:'center',line:1.5},{pt:isAff?10:11,bold:true,italic:false});return;}
      if(state==='abstract'||state==='abstrak'){abstracts[state].push(txt);formatP(p,{align:'both',line:1},{pt:11,italic:false});return;}
      if(state==='references'){
        const category=/^(buku|book|artikel\s+dalam\s+jurnal|artikel\s+jurnal|journal\s+articles?)$/i.test(txt);
        formatP(p,{align:'left',line:1,left:category?0:567,hanging:category?0:567,after:category?0:100},{pt:12,bold:category,underline:category});return;
      }
      formatP(p,{align:'both',line:1.5,first:850},{pt:12});
    });
    if(titleWords.join(' ').split(/\s+/).length>20)warnings.push('Judul RABBANI melebihi 20 kata; isi tidak dipotong.');
    for(const kind of ['abstract','abstrak']){const n=abstracts[kind].join(' ').split(/\s+/).filter(Boolean).length;if(!n)warnings.push(kind.toUpperCase()+' tidak ditemukan.');else if(n<150||n>200)warnings.push(kind.toUpperCase()+': '+n+' kata (panduan menyarankan sekitar 150–200).');}
    for(const k of ['intro','method','results','conclusion','references'])if(!stateSet.has(k))warnings.push('RABBANI: bagian tidak ditemukan: '+k);
    if(opts.profile==='template2023')createRabbaniAbstractTable(doc,paragraphs,meta,warnings);
    warnings.push('RABBANI: nomor WA dan metadata publikasi yang kosong perlu diisi manual. Chicago 17 Full Footnote tidak dikonversi otomatis.');
  }
  function updateAlihkam(doc,paragraphs,opts,meta,warnings) {
    let state=null,seen=new Set(),abstract=[];
    paragraphs.forEach((p,i)=>{
      const txt=normalize(docParaText(p)),h=sectionHead(txt);if(!txt)return;
      if(i<opts.titleParagraphs){formatP(p,{align:'center',line:1,before:i?0:150,after:150,keepNext:true},{pt:13,bold:true});return;}
      if(h){state=h;seen.add(h);formatP(p,{align:'left',line:1,before:150,after:0,keepNext:true},{pt:['abstract','abstrak'].includes(h)?10:11,bold:true});return;}
      if(/^(keywords?|kata\s*kunci)\s*:/i.test(txt)){
        formatP(p,{align:'center',line:1,after:150},{pt:10});
        const parts=txt.split(':');if(parts.length>1){const values=parts.slice(1).join(':').split(/[;,]/).map(x=>x.trim()).filter(Boolean);
          if(values.length<3||values.length>5)warnings.push('AL-IHKAM mensyaratkan 3–5 kata kunci.');}
        state='keywords';return;
      }
      if(state===null){let aff=/@|universitas|institut|university|\bUIN\b|\bIAIN\b/i.test(txt);formatP(p,{align:'center',line:1},{pt:aff?10:11,bold:!aff,italic:aff});return;}
      if(state==='abstract'||state==='abstrak'){abstract.push(txt);formatP(p,{align:'both',line:1},{pt:10});return;}
      if(state==='references'){formatP(p,{align:'left',line:1,after:80},{pt:11});return;}
      if(/^(table|tabel|figure|gambar)\s+\d+[.:]/i.test(txt)){formatP(p,{align:'left',line:1},{pt:11});return;}
      formatP(p,{align:'both',line:1,first:0},{pt:11});
    });
    if(abstract.join(' ').split(/\s+/).filter(Boolean).length>250)warnings.push('Abstrak lebih dari 250 kata, tidak dipotong.');
    for(const k of ['intro','method','results','conclusion','references'])if(!seen.has(k))warnings.push('AL-IHKAM: bagian tidak ditemukan: '+k);
    warnings.push('AL-IHKAM: aturan sitasi Chicago 17, histori artikel, DOI, dan tata letak header perlu verifikasi di Word.');
  }
  function validateFile(file,opts) {
    if(!file)throw new Error('Silakan pilih artikel DOCX terlebih dahulu.');
    if(!/\.docx$/i.test(file.name||''))throw new Error('Hanya file .docx yang didukung.');
    if(file.size>15_000_000)throw new Error('Ukuran artikel maksimum 15 MB.');
    if(file.size<100)throw new Error('File terlalu kecil atau kosong.');
    if(!['karsa','rabbani','alihkam'].includes(opts.templateId))throw new Error('Template tidak dikenal.');
    if(![1,2,3].includes(opts.titleParagraphs))throw new Error('Jumlah paragraf judul harus 1–3.');
  }
  async function formatFile(file,opts,progress=()=>{}){
    validateFile(file,opts);
    if(typeof JSZip==='undefined')throw new Error('Pustaka JSZip belum termuat; periksa folder vendor.');
    const warnings=[], meta={...(opts.metadata||{})};
    progress('Membaca struktur dokumen Word…');
    const orig=await JSZip.loadAsync(await file.arrayBuffer(),{checkCRC32:false});
    const d=await readPart(orig,'word/document.xml');
    const relationships=await readPart(orig,'word/_rels/document.xml.rels');
    const types=await readPart(orig,'[Content_Types].xml');
    const styles=await readPart(orig,'word/styles.xml',false);
    const settings=await readPart(orig,'word/settings.xml',false);
    const paras=all(d,W,'p').filter(p=>p.parentNode && p.parentNode.localName==='body' && normalize(docParaText(p)));
    if(paras.length<5)throw new Error('Artikel terlalu singkat. Dibutuhkan judul, nama penulis, dan isi minimal 5 paragraf.');
    if(opts.titleParagraphs>=paras.length || sectionHead(docParaText(paras[0])))throw new Error('Judul artikel belum terdeteksi pada paragraf awal.');
    if(opts.templateId!=='karsa' && !paras.some(p=>sectionHead(docParaText(p))==='intro'))throw new Error('Bagian PENDAHULUAN / INTRODUCTION harus berada di paragraf tersendiri.');
    meta.title=paras.slice(0,opts.titleParagraphs).map(docParaText).join(' ').trim();
    meta.author=normalize(docParaText(paras[opts.titleParagraphs]));
    if(!meta.running_title)meta.running_title=meta.title.length>65?meta.title.slice(0,62).replace(/\s+\S*$/,'')+'…':meta.title;
    const emailText=paras.slice(0,7).map(docParaText).join(' ').match(/[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}/);
    if(!meta.email && emailText)meta.email=emailText[0];
    progress('Menerapkan aturan paragraf jurnal…');
    if(opts.templateId==='rabbani')updateRabbani(d,paras,opts,meta,warnings);
    else if(opts.templateId==='alihkam')updateAlihkam(d,paras,opts,meta,warnings);
    else updateKarsa(paras,warnings);
    setNormalStyle(styles,opts.templateId==='karsa'?'Times New Roman':'Book Antiqua',opts.templateId==='alihkam'?11:12,opts.templateId==='alihkam'?1:1.5);
    if(styles)orig.file('word/styles.xml',xmlString(styles));
    progress('Menyalin tata halaman dan header jurnal…');
    const templateUrl=new URL('templates/'+encodeURIComponent(opts.templateId)+'/template.docx',document.baseURI);
    const tplResponse=await fetch(templateUrl,{cache:'force-cache'});
    if(!tplResponse.ok)throw new Error('Template belum tersedia di GitHub Pages: '+opts.templateId+' ('+tplResponse.status+').');
    const template=await JSZip.loadAsync(await tplResponse.arrayBuffer());
    const tplDoc=await readPart(template,'word/document.xml');
    const tplRels=await readPart(template,'word/_rels/document.xml.rels');
    setPageSizeFromTemplate(d,tplDoc);
    if(settings){addEvenOddSettings(settings);orig.file('word/settings.xml',xmlString(settings));}
    await transplantStories(orig,template,d,tplDoc,relationships,tplRels,types,meta,warnings,opts.templateId);
    orig.file('word/document.xml',xmlString(d));
    orig.file('word/_rels/document.xml.rels',xmlString(relationships));
    orig.file('[Content_Types].xml',xmlString(types));
    progress('Membuat file Word hasil pemformatan…');
    const output=await orig.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:6}}, data => {if(data.percent>50)progress('Membuat file Word… '+Math.round(data.percent)+'%');});
    return {blob:output,warnings,filename:'ZAINNET_'+opts.templateId.toUpperCase()+'_HASIL_'+(file.name||'artikel.docx').replace(/\.docx$/i,'').replace(/[^a-zA-Z0-9_\-.]/g,'_').slice(0,65)+'.docx'};
  }
  window.ZainFormatter={formatFile,sectionHead,version:'2.0.0-pages-only'};
})();
