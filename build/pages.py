import subprocess,json,re
n=int(re.search(r'Pages:\s+(\d+)',subprocess.run(['pdfinfo','../FP_Report.pdf'],capture_output=True,text=True).stdout).group(1))
txt=[subprocess.run(['pdftotext','-f',str(i),'-l',str(i),'-layout','../FP_Report.pdf','-'],capture_output=True,text=True).stdout for i in range(1,n+1)]
def roman(k):
    v=[(10,'x'),(9,'ix'),(5,'v'),(4,'iv'),(1,'i')];o=''
    for a,s in v:
        while k>=a:o+=s;k-=a
    return o
ch1=next(i for i,t in enumerate(txt) if re.search(r'^\s*CHAPTER 1\s*$',t,re.M))
def find(pat,start=0,caption=False):
    for i in range(start,n):
        if re.search(pat,txt[i],re.M): return i
    raise Exception(pat)
def lab(i): return roman(i+1) if i<ch1 else str(i-ch1+1)
keys={'cert':r'^\s*CERTIFICATE\s*$','certorg':r'CERTIFICATE OF ORGANISATION','decl':r'^\s*SELF-DECLARATION','ack':r'^\s*ACKNOWLEDGEMENTS','lof':r'^\s*LIST OF FIGURES AND TABLES\s*$',
 'ch1':r'^\s*CHAPTER 1\s*$','ch2':r'^\s*CHAPTER 2\s*$','ch3':r'^\s*CHAPTER 3\s*$','ch4':r'^\s*CHAPTER 4\s*$','ch5':r'^\s*CHAPTER 5\s*$','refs':r'^\s*REFERENCES\s*$','app':r'^\s*APPENDICES\s*$','prog':r'^\s*PROGRESS REPORT\s*$'}
out={k:lab(find(p)) for k,p in keys.items()}
for k in range(1,8): out[f'fig{k}']=lab(find(rf'^\s*Figure {k}: ',ch1))
for k in range(1,4): out[f'g{k}']=lab(find(rf'^\s*Graph {k}: ',ch1))
for t in ['1.1','3.1','4.1','4.2','4.3','4.4','5.1']: out['t'+t.replace('.','')]=lab(find(rf'^\s*Table {re.escape(t)}: ',ch1))
json.dump(out,open('pages.json','w'),indent=0);print(out)
