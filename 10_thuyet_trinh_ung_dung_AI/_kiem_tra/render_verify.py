from pathlib import Path
from xml.sax.saxutils import escape
import subprocess, fitz, json, zipfile, xml.etree.ElementTree as E, hashlib
root=Path(__file__).resolve().parent.parent
docx=next(root.glob('*.docx'));qa=root/'_kiem_tra';folder=qa/'render';folder.mkdir(exist_ok=True)
pdf=folder/'render.pdf';params=qa/'convert.xml'
params.write_text('<TaskQueueDataConvert><m_sFileFrom>'+escape(str(docx))+'</m_sFileFrom><m_sFileTo>'+escape(str(pdf))+'</m_sFileTo><m_nFormatTo>513</m_nFormatTo><m_sFontDir>C:/Users/PC/AppData/Local/ONLYOFFICE/DesktopEditors/data/fonts</m_sFontDir><m_sAllFontsPath>C:/Users/PC/AppData/Local/ONLYOFFICE/DesktopEditors/data/fonts/AllFonts.js</m_sAllFontsPath><m_bIsNoBase64>true</m_bIsNoBase64></TaskQueueDataConvert>',encoding='utf-8')
r=subprocess.run(['C:/Program Files/ONLYOFFICE/DesktopEditors/converter/x2t.exe',str(params)],capture_output=True,text=True,timeout=60)
assert r.returncode==0,(r.stdout,r.stderr)
report={'file':str(docx),'sha256':hashlib.sha256(docx.read_bytes()).hexdigest(),'pages':[],'timing_seconds':420}
d=fitz.open(pdf)
for i,p in enumerate(d):
    words=p.get_text('words');assert words
    assert abs(p.rect.width-595.3)<1 and abs(p.rect.height-841.9)<1
    assert all(w[0]>=50 and w[2]<=p.rect.width-45 for w in words)
    p.get_pixmap(matrix=fitz.Matrix(1.5,1.5)).save(folder/f'page-{i+1}.png')
    text=p.get_text();(folder/f'page-{i+1}.txt').write_text(text,encoding='utf-8')
    report['pages'].append({'page':i+1,'characters':len(text),'start':text[:90]})
with zipfile.ZipFile(docx) as z:
    w=lambda x:'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'+x
    doc=E.fromstring(z.read('word/document.xml'))
    text=' '.join(t.text or '' for t in doc.iter(w('t')))
    for i in range(1,8):assert f'Slide {i}' in text
    for size in doc.iter(w('sz')):assert size.get(w('val'))=='26'
    for color in doc.iter(w('color')):assert color.get(w('val'))=='000000'
    assert len(doc.findall('.//'+w('pageBreakBefore')))==6
(qa/'verification.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(report,ensure_ascii=False,indent=2))
