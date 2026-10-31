from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, FileResponse
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Pt, Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pathlib import Path
from io import BytesIO
from copy import deepcopy
import json, re, os
from urllib.parse import quote
from .rabbani import format_rabbani
from .alihkam import format_alihkam

ROOT=Path(__file__).resolve().parent.parent
app=FastAPI(title='ZAIN.NET Article Formatter — Beta')
origins=[s.strip() for s in os.getenv('ALLOWED_ORIGINS','*').split(',')]
app.add_middleware(CORSMiddleware,allow_origins=origins,allow_methods=['GET','POST'],allow_headers=['*'],expose_headers=['X-Format-Warnings','Content-Disposition'])

# Satu hosting: FastAPI juga melayani halaman web. Tidak perlu URL API di UI.
@app.get('/', include_in_schema=False)
def homepage():
    return FileResponse(ROOT/'index.html', media_type='text/html')

@app.get('/app.js', include_in_schema=False)
def app_js():
    return FileResponse(ROOT/'app.js', media_type='application/javascript')

@app.get('/config.js', include_in_schema=False)
def config_js():
    return FileResponse(ROOT/'config.js', media_type='application/javascript')

@app.get('/health')
def health(): return {'ok':True,'mode':'beta','warning':'Format kompleks perlu validasi visual'}

@app.get('/templates')
def templates():
    return [json.loads(p.read_text(encoding='utf-8')) for p in (ROOT/'templates').glob('*/rules.json') if (p.parent/'template.docx').exists()]

def clean(s):return re.sub(r'\s+',' ',s).strip()
def text_nodes(doc):return [p for p in doc.paragraphs if clean(p.text)]
def find_head(s):
    t=clean(s).upper()
    if t in ('ABSTRAK','ABSTRACT','DAFTAR PUSTAKA','REFERENCES','REFERENCE','BIBLIOGRAPHY'):return t
    m=re.match(r'^[A-E]\.\s*(PENDAHULUAN|METODE PENELITIAN|HASIL PENELITIAN|PEMBAHASAN|KESIMPULAN DAN SARAN|KESIMPULAN)',t)
    return m.group(1) if m else None

def parse_article(doc):
    paras=text_nodes(doc)
    if len(paras)<5:raise ValueError('Dokumen tidak berisi struktur artikel yang cukup')
    title=[];i=0
    while i<len(paras) and len(title)<5:
        s=clean(paras[i].text)
        if ('@' in s or re.search('^(ABSTRAK|ABSTRACT)$',s,re.I)):break
        if i>0 and ('universitas' in s.lower() or 'program studi' in s.lower()):break
        if i>0 and s.istitle() and len(s)<80:break
        title.append(s);i+=1
    if not title: title=[clean(paras[0].text)];i=1
    author=[]
    while i<len(paras) and not find_head(paras[i].text):
        author.append(clean(paras[i].text));i+=1
    sections=[];current=None
    for p in paras[i:]:
        h=find_head(p.text)
        if h:
            current={'heading':h,'paragraphs':[]};sections.append(current)
        elif current:current['paragraphs'].append(p)
    return {'title':' '.join(title),'author':author,'sections':sections}

def replace_para(p,text):
    for r in p.runs:r.text=''
    if p.runs:p.runs[0].text=text
    else:p.add_run(text)

def set_page_number(footer):
    for p in footer.paragraphs:
        if 'PAGE' in p._p.xml or re.fullmatch(r'\s*\d+\s*',p.text or ''):
            p._element.getparent().remove(p._element)
    p=footer.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    field=OxmlElement('w:fldSimple');field.set(qn('w:instr'),'PAGE');p._p.append(field)

def make_docx(source,template):
    # Preserve source DOCX as the base to retain genuine Word footnotes,
    # tables, drawings, and relationship IDs. Do NOT recreate them from plain text.
    doc=Document(BytesIO(source))
    tpl=Document(template)
    data=parse_article(doc)
    # Copy page settings from original template, not its article content.
    for sec in doc.sections:
        ref=tpl.sections[0]
        for a in ('page_width','page_height','top_margin','bottom_margin','left_margin','right_margin','header_distance','footer_distance'):
            setattr(sec,a,getattr(ref,a))
        sec.different_first_page_header_footer=True
        set_page_number(sec.footer)
        # Copy template's actual header/footer XML for first page when available.
        # Word header/footer structures are not blindly transplanted: external
        # image relationship IDs would be broken. This is a known beta limit.
    paras=text_nodes(doc)
    if paras:
        # Keep text unchanged; only adjust formatting.
        title_parts=data['title'].split(' ')
        # Apply title formatting to the paragraphs preceding author.
        for p in paras:
            if '@' in p.text or 'program studi' in p.text.lower():break
            if find_head(p.text):break
            if len(p.text)>110 and p not in paras[:3]:break
            p.alignment=WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:r.font.size=Pt(13);r.bold=True
    for sec in data['sections']:
        h=sec['heading']
        # Locate matching original heading without altering its content.
        for p in doc.paragraphs:
            if find_head(p.text)==h:
                p.alignment=WD_ALIGN_PARAGRAPH.LEFT
                for r in p.runs:r.bold=True
                break
        for p in sec['paragraphs']:
            if p.style.name.lower().startswith(('caption','footnote')):continue
            p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    # Insert journal identity as a true first-page header table with fixed cells.
    first=doc.sections[0].first_page_header
    for node in list(first._element):
        if node.tag not in (qn('w:hdr'),): first._element.remove(node)
    # Use the template's logo image as a byte source, without inventing it.
    logo=None
    for shape in tpl.inline_shapes:
        try:
            rel=tpl.part.related_parts[shape._inline.graphic.graphicData.pic.blipFill.blip.embed]
            logo=rel.blob;break
        except Exception:pass
    tab=first.add_table(rows=1,cols=2,width=Cm(11.9))
    tab.autofit=False;tab.columns[0].width=Cm(2.6);tab.columns[1].width=Cm(9.3)
    c0,c1=tab.rows[0].cells;c0.vertical_alignment=c1.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    if logo:
        try:c0.paragraphs[0].add_run().add_picture(BytesIO(logo),width=Cm(2.0))
        except Exception:pass
    lines=['Karsa: Journal of Social and Islamic Culture','ISSN: 2442-3289 (p); 2442-8285 (e)','Vol. XX No.X, December 20XX, pp. XX–XX','DOI: [belum ditetapkan]']
    for j,s in enumerate(lines):
        p=c1.paragraphs[0] if j==0 else c1.add_paragraph()
        p.add_run(s).font.size=Pt(8)
        p.paragraph_format.space_after=Pt(0)
    # These are journal sample placeholders, not real publication metadata.
    return doc

@app.post('/format')
async def format_article(file:UploadFile=File(...),template_id:str=Form('karsa'),profile:str=Form('template2023'),title_paragraphs:int=Form(1),metadata:str=Form('{}')):
    if not file.filename.lower().endswith('.docx'):raise HTTPException(400,'Hanya DOCX didukung pada beta')
    if template_id not in ('karsa','rabbani','alihkam'):raise HTTPException(404,'Template tidak ditemukan')
    if profile not in (('template2026','template2023') if template_id=='alihkam' else ('template2023','guidelines')):raise HTTPException(400,'Pilihan aturan tidak valid')
    template=ROOT/'templates'/template_id/'template.docx'
    if not template.exists():raise HTTPException(404,'Template tidak ditemukan')
    source=await file.read()
    if len(source)>15_000_000:raise HTTPException(413,'Maksimum 15 MB')
    try:
        warnings=[]
        if template_id in ('rabbani','alihkam'):
            meta=json.loads(metadata)
            if not isinstance(meta,dict) or any(not isinstance(v,str) or len(v)>200 for v in meta.values()):
                raise ValueError('Metadata harus berisi teks maksimal 200 karakter per kolom.')
            if template_id=='rabbani':doc,warnings=format_rabbani(source,template,profile,title_paragraphs,meta)
            else:doc,warnings=format_alihkam(source,template,title_paragraphs,meta)
        else:doc=make_docx(source,template)
        out=BytesIO();doc.save(out);out.seek(0)
    except Exception as e:raise HTTPException(422,f'Gagal memproses: {str(e)}')
    return StreamingResponse(out,media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',headers={'Content-Disposition':f'attachment; filename="ZAINNET_{template_id.upper()}_BETA.docx"','X-Processing-Status':'BETA-REQUIRES-REVIEW','X-Format-Warnings':quote(json.dumps(warnings,ensure_ascii=False))})
