import subprocess,re,json,sys,docx
pdf=sys.argv[1]; docxf=sys.argv[2]
d=docx.Document(docxf)
heads=[p.text.strip() for p in d.paragraphs if p.style.name in ('Heading 1','Heading 2')]
n=int(re.search(r'Pages:\s+(\d+)',subprocess.run(['pdfinfo',pdf],capture_output=True,text=True).stdout).group(1))
norm=lambda s: re.sub(r'\s+','',s)
pages=[]
for i in range(1,n+1):
    t=subprocess.run(['pdftotext','-f',str(i),'-l',str(i),pdf,'-'],capture_output=True,text=True).stdout
    m=re.findall(r'-\s*([ivx]+|\d+)\s*-',t); pages.append((t,m[-1] if m else ''))
out={}
for h in heads:
    for t,pn in pages:
        if '차례' in norm(t)[:20] or norm(t).startswith('차례'): continue
        lines=[norm(l) for l in t.split('\n')]
        if norm(h) in lines or any(l.startswith(norm(h)) and len(l)<=len(norm(h))+2 for l in lines):
            out[norm(h)]=pn; break
print(len(out),'/',len(heads),[h for h in heads if norm(h) not in out])
json.dump(out,open(sys.argv[3],'w'),ensure_ascii=False)
