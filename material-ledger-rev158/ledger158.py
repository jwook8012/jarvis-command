# -*- coding: utf-8 -*-
"""자재검수대장 Rev1.5.8 — offline desktop ledger, conservative OCR review."""
from __future__ import annotations
import argparse, copy, csv, hashlib, json, os, re, shutil, sys, tempfile, threading, time, traceback
from datetime import datetime
from pathlib import Path
from PIL import Image, ImageOps, ImageEnhance
import fitz
import numpy as np
import pytesseract
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

VERSION='Rev1.5.8'
HEADERS=['문서번호','검수일자','자재명','규격','단위','반입수량','합격수량','불합격수량','반입자','관련문서번호','시공자','감리자','공구']
COLS=[(.109,.313),(.314,.410),(.418,.485),(.494,.577),(.585,.668),(.674,.758)]
RE_DOC=re.compile(r'자검\s*\(\s*(건축|토목|기계|전기)\s*\)\s*[-_]?\s*(\d{2})\s*[-_]?\s*(\d{3,5})')
RE_DATE=re.compile(r'(20\d{2})\s*[-./년]\s*(1[0-2]|0?[1-9])\s*[-./월]\s*(3[01]|[12]\d|0?[1-9])(?!\d)')
DEFAULT_NAMES='김진상, 류재호, 윤희영, 김병윤, 최근수, 유재율, 이희재, 신의규, 김대식, 신경섭, 정동진, 서재원, 이석기, 최용현, 최경수, 박노준, 김창현, 장종국, 장원익, 강경준'
SPEC_KNOWN=['HD10','HD13','HD16','HD19','HD22','SHD10','SHD13','SHD16','SHD19','SHD22','UHD16','UHD19','UHD22','UHD22S','UHD25','UHD29','UHD32']

def app_dir():
    p=Path(os.getenv('LOCALAPPDATA') or (Path.home()/'.config'))/'MaterialLedgerRev158'
    p.mkdir(parents=True,exist_ok=True); return p

def cfg_load():
    p=app_dir()/'settings.json'
    if p.exists():
        try: return json.loads(p.read_text(encoding='utf-8'))
        except Exception: pass
    p_old=Path(os.getenv('LOCALAPPDATA') or (Path.home()/'.config'))/'MaterialLedgerRev15'/'settings.json'
    if p_old.exists():
        try:
            o=json.loads(p_old.read_text(encoding='utf-8'))
            return {'names':o.get('Names') or o.get('names') or DEFAULT_NAMES,'folder':o.get('Folder') or '', 'output':o.get('Output') or ''}
        except Exception:pass
    return {'names':DEFAULT_NAMES,'folder':'','output':''}

def cfg_save(data):
    p=app_dir()/'settings.json';tmp=p.with_suffix('.tmp');tmp.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8');os.replace(tmp,p)

def find_tesseract():
    root=Path(getattr(sys,'_MEIPASS',Path(__file__).resolve().parent))
    base=Path(sys.executable).resolve().parent if getattr(sys,'frozen',False) else root
    candidates=[base/'tesseract'/'tesseract.exe',root/'tesseract'/'tesseract.exe',base/'tesseract.exe']
    for p in candidates:
        if p.exists():
            pytesseract.pytesseract.tesseract_cmd=str(p)
            os.environ['TESSDATA_PREFIX']=str(p.parent/'tessdata')
            return str(p)
    if sys.platform!='win32':
        return shutil.which('tesseract')
    raise RuntimeError('내장된 tesseract.exe를 찾지 못했습니다. OCR 런타임을 포함한 정식 배포본이 필요합니다.')

def page_first(path):
    doc=fitz.open(str(path))
    try:
        page=doc.load_page(0)
        # 1370px on A4 page: each form row gets ~40px, adequate for text OCR after 3x upsampling
        zoom=2.3
        pix=page.get_pixmap(matrix=fitz.Matrix(zoom,zoom),alpha=False,colorspace=fitz.csRGB)
        return Image.frombytes('RGB',(pix.width,pix.height),pix.samples)
    finally: doc.close()

def norm_text(s):return re.sub(r'\s+',' ',s).strip()

def ocr(img,lang='kor+eng',psm=7,whitelist=None):
    if img.width<5 or img.height<5:return ''
    if img.width<90 or img.height<30: img=img.resize((img.width*3,img.height*3),Image.Resampling.LANCZOS)
    config=f'--oem 1 --psm {psm}'
    if whitelist:config+=' -c tessedit_char_whitelist='+whitelist
    try:
        return norm_text(pytesseract.image_to_string(img,lang=lang,config=config,timeout=18))
    except Exception as e:
        raise RuntimeError('한글 OCR 실패: '+str(e))

def crop_pct(img,x0,y0,x1,y1,scale=3,mode='color'):
    w,h=img.size
    b=(max(0,int(x0*w)),max(0,int(y0*h)),min(w,int(x1*w)),min(h,int(y1*h)))
    c=img.crop(b)
    if scale!=1: c=c.resize((max(1,c.width*scale),max(1,c.height*scale)),Image.Resampling.LANCZOS)
    return c

def row_bands(img):
    a=np.array(img.convert('L'))
    h,w=a.shape
    xa,xb=int(w*.12),int(w*.70)
    scores=np.mean(a[int(h*.257):int(h*.512),xa:xb:3]<155,axis=1)
    hits=np.where(scores>=.37)[0]+int(h*.257)
    groups=[]
    for y in hits:
        if not groups or y-groups[-1][-1]>max(2,int(.0019*h)):groups.append([])
        groups[-1].append(y)
    lines=[np.mean(q)/h for q in groups]
    firsts=[q for q in lines if .274<=q<=.301]
    if not firsts:return []
    start=min(range(len(lines)),key=lambda k:abs(lines[k]-firsts[0]))
    bands=[]
    for i in range(start,len(lines)-1):
        dy=lines[i+1]-lines[i]
        if not (.014<dy<.034):break
        top,bottom=lines[i]+.0023,lines[i+1]-.0022
        if bottom>top:bands.append((top,bottom))
    return bands[:12]

def real_ink(img):
    a=np.asarray(img.convert('L'))
    # colors red stamp & blue seal do not count as typed material cells
    return np.mean(a[::3,::3]<155)

def normalize_spec(s):
    s=s.upper().replace('|','').replace(' ','').lstrip('J{[(〈').rstrip('.,')
    s=re.sub(r'[^0-9A-Z.\-]','',s)
    if not s:return ''
    if s in SPEC_KNOWN:return s
    if s.startswith('UHD') and s.endswith('225') and 'UHD22S' in SPEC_KNOWN:return 'UHD22S'
    # Optical confusions mainly affect the alphabetic prefix of steel grades.
    s=s.replace('0HD','UHD').replace('1HD','UHD')
    if s in SPEC_KNOWN:return s
    # Conservative one-edit correction of standard steel grades only.
    def lev(a,b):
        old=list(range(len(b)+1))
        for i,c in enumerate(a,1):
            cur=[i]
            for j,d in enumerate(b,1):cur.append(min(cur[-1]+1,old[j]+1,old[j-1]+(c!=d)))
            old=cur
        return old[-1]
    if len(s)<7 and '-' not in s:
        ranked=sorted((lev(s,k),k) for k in SPEC_KNOWN)
        if ranked[0][0]<=2 and (len(ranked)<2 or ranked[1][0]>ranked[0][0]):return ranked[0][1]
    return s

def num_ocr(c):
    texts=[]
    for psm in (7,8):
        t=ocr(c,lang='eng',psm=psm,whitelist='0123456789.,')
        t=re.sub(r'[^0-9.,]','',t).replace(',','')
        if re.fullmatch(r'\d+(?:\.\d+)?',t):texts.append(t)
    return texts[0] if texts else ''

def parse_num(s):
    try:return float(s)
    except:return None

def normalize_unit(s):
    s=s.upper().replace('|','').replace(' ','')
    if 'TON' in s or s in ('T0N','T0M'):return 'ton'
    if 'EA' in s or s in ('E4','F4'):return 'EA'
    if 'KG' in s:return 'kg'
    return s

def extract_line_items(img):
    bands=row_bands(img)
    if not bands:return [],'표 가로선 탐지 실패'
    records=[]
    for i,(y0,y1) in enumerate(bands):
        material_roi=crop_pct(img,COLS[0][0]+.005,y0,COLS[0][1]-.006,y1,scale=1)
        if real_ink(material_roi)<.035:continue
        # printed header/footer & stamps may be interpreted as material; validate spec and numeric below
        values=[]
        for j,(x0,x1) in enumerate(COLS):
            c=crop_pct(img,x0+.005,y0,x1-.005,y1,scale=3)
            if j in (3,4,5):s=num_ocr(c)
            elif j in (1,2):s=ocr(c,lang='eng',psm=7 if j==1 else 8)
            else:s=ocr(c,lang='kor+eng',psm=7)
            values.append(s)
        material,spec,unit,qty,passed,failed=values
        material=material.strip(' |')
        spec=normalize_spec(spec)
        unit=normalize_unit(unit)
        if i>0 and (not spec or not re.search(r'\d',spec)):
            break  # First empty line ends the occupied material portion; totals are not material.
        if not material and not spec:continue
        issues=[]
        if not material:issues.append('품명 미인식')
        if not spec:issues.append('규격 미인식')
        if unit not in ('EA','ton','kg','m3','M3','개','본','매'):issues.append('단위 미인식')
        for k,v in [('반입',qty),('합격',passed),('불합격',failed)]:
            if parse_num(v) is None:issues.append(f'{k}수량 미인식')
        a,b,c=map(parse_num,(qty,passed,failed))
        if all(v is not None for v in (a,b,c)) and abs(a-b-c)>0.0001:
            # Retry the received quantity with alternate OCR segmentation and crop
            # before accepting an inconsistent row (never fabricate a new number).
            expected=b+c
            x0,x1=COLS[3]
            fixed=False
            for pad,mode in ((.002,10),(.002,8),(.016,8),(.002,7)):
                test=ocr(crop_pct(img,x0+pad,y0,x1-pad,y1),lang='eng',psm=mode,whitelist='0123456789.,')
                match=re.search(r'\d+(?:\.\d+)?',test.replace(',',''))
                if match and abs(float(match[0])-expected)<0.0001:
                    qty=match[0];fixed=True;break
            if not fixed:issues.append('수량 합계 불일치')
        records.append({'material':material,'spec':spec,'unit':unit,'qty':qty,'passed':passed,'failed':failed,'issues':issues})
    return records,'' if records else '자재표 인식 실패'

def parse_date_text(s):
    for m in RE_DATE.finditer(s):
        try:
            dt=datetime(int(m[1]),int(m[2]),int(m[3]))
            if 2020<=dt.year<=2100:return dt.strftime('%Y-%m-%d')
        except ValueError:pass
    return ''

def closest_name(handwritten,names):
    toks=re.findall(r'[가-힣]{2,4}',handwritten.replace(' ','').replace('\n',' '))
    roster=[x.strip() for x in re.split(r'[,，;\n]',names) if x.strip()]
    out=[]
    for tok in toks:
        if tok in roster and tok not in out:out.append(tok)
    return '\n'.join(out[:2])

def parse_request(path, names=''):
    img=page_first(path)
    # OCR only likely date and identity areas, not entire page where table lines scatter the text
    date_raw=ocr(crop_pct(img,.13,.563,.47,.635),psm=6)
    date=parse_date_text(date_raw)
    if not date:
        date_raw=ocr(crop_pct(img,.09,.51,.52,.65),psm=6)
        date=parse_date_text(date_raw)
    material,mat_error=extract_line_items(img)
    m=RE_DOC.search(Path(path).name)
    if not m:raise ValueError('PDF 파일명에서 문서번호를 읽지 못했습니다')
    trade,yy,num=m.groups();doc=f'자검({trade})-{yy}-{num}'
    request=f'MRI-GF79-A-20{yy}-{num}'
    reviewer_roi=crop_pct(img,.60,.835,.91,.884)
    # OCR is conservative: only exact readable names; signatures alone remain blank.
    reviewer_ocr=ocr(reviewer_roi,psm=6)
    inspector=closest_name(reviewer_ocr,names)
    contractor=ocr(crop_pct(img,.69,.568,.87,.610),psm=7)
    contractor=''.join(re.findall(r'[가-힣]{2,4}',contractor))[:4]
    area_ocr=ocr(crop_pct(img,.51,.15,.90,.23),psm=6)
    area=re.search(r'([1-9])\s*공구',area_ocr)
    area=area[1]+'공구' if area else ''
    rows=[]
    if not material:material=[dict(material='',spec='',unit='',qty='',passed='',failed='',issues=[mat_error or '자재표 인식 실패'])]
    for it in material:
        issues=list(it['issues'])
        if not date:issues.append('검수일자 확인 필요')
        rows.append({'file':str(path),'doc':doc,'date':date,'request':request,'contractor':contractor,
                     'inspector':inspector,'area':area,**it,'issues':issues})
    return rows

def full_key(row):return (row['request'],row['spec'].upper(),row['qty'],row['material'])

def valid_record(row):return bool(row['date'] and row['material'] and row['spec'] and row['unit'] and row['qty'] and row['passed'] and row['failed'] and not row['issues'])

def save_output(records,dest,base,existing=None,new=True):
    if not records: raise ValueError('저장할 분석 결과가 없습니다.')
    target=Path(dest).resolve();src=Path(existing or base).resolve()
    if target==src: raise ValueError('원본과 다른 결과 파일명을 지정해야 합니다.')
    if not src.is_file():raise FileNotFoundError(str(src))
    wb=load_workbook(src,keep_vba=True)
    if new:
        # Never write into an old month's historical data; blank the data cells only.
        for sheet in wb:
            if not re.fullmatch(r'\d{1,2}월',sheet.title):continue
            for row in sheet.iter_rows(min_row=5,max_row=max(sheet.max_row,5),min_col=1,max_col=13):
                for c in row:
                    if c.__class__.__name__=='MergedCell':continue
                    c.value=None
    existing_keys=set()
    if not new:
        for sheet in wb:
            if not re.fullmatch(r'\d{1,2}월',sheet.title):continue
            last_ref=''
            for r in range(5,sheet.max_row+1):
                ref=sheet.cell(r,10).value
                if isinstance(ref,str) and ref.startswith('MRI'):last_ref=ref
                spec=sheet.cell(r,4).value; qty=sheet.cell(r,6).value; name=sheet.cell(r,3).value
                if last_ref and spec and qty is not None:existing_keys.add((last_ref,str(spec).upper(),str(qty),str(name or '')))
    month_data={}
    for rec in records:
        if not valid_record(rec):continue
        m=int(rec['date'][5:7]);month_data.setdefault(m,[]).append(rec)
    counts={'written':0,'duplicate':0,'review':sum(not valid_record(r) for r in records)}
    for m,data in month_data.items():
        title=f'{m}월'
        if title not in wb.sheetnames:
            proto=wb['10월'] if '10월' in wb else wb['9월']
            s=wb.copy_worksheet(proto);s.title=title
            for row in s.iter_rows(min_row=5,max_row=s.max_row,min_col=1,max_col=13):
                for c in row:
                    if c.__class__.__name__!='MergedCell':c.value=None
        s=wb[title]
        # Preserve historical entries, append new records. Excel rows use original column mapping.
        existing=[]
        if not new:
            lastdoc=''
            for r in range(5,s.max_row+1):
                dt=s.cell(r,2).value
                if dt:
                    if s.cell(r,1).value:lastdoc=str(s.cell(r,1).value)
                    existing.append([s.cell(r,c).value for c in range(1,14)])
        for rec in data:
            key=full_key(rec)
            if key in existing_keys:
                counts['duplicate']+=1; continue
            existing_keys.add(key)
            try:d=datetime.strptime(rec['date'],'%Y-%m-%d').date()
            except ValueError:continue
            existing.append([rec['doc'],d,rec['material'],rec['spec'],rec['unit'],float(rec['qty']),
                             float(rec['passed']),float(rec['failed']),None,rec['request'],rec['contractor'],rec['inspector'],rec['area']])
            counts['written']+=1
        existing.sort(key=lambda x:(str(x[1])[:10],str(x[0] or ''),str(x[3] or '')))
        for i,line in enumerate(existing,5):
            for j,val in enumerate(line,1):
                c=s.cell(i,j);c.value=val
                if j==2:c.number_format='yyyy-mm-dd'
                if j==12:c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
            s.row_dimensions[i].height=31 if '\n' in str(line[11] or '') else 22
        s.freeze_panes='C5'
        s.auto_filter.ref=f'A4:M{max(5,4+len(existing))}'
        for col,width in {'A':19,'B':15,'C':33,'D':20,'E':9,'F':13,'G':13,'H':13,'I':21,'J':30,'K':14,'L':19,'M':10}.items():
            s.column_dimensions[col].width=width
    tmp=target.with_name(target.stem+'-saving'+target.suffix)
    target.parent.mkdir(parents=True,exist_ok=True)
    wb.save(tmp)
    os.replace(tmp,target)
    review=[r for r in records if not valid_record(r)]
    review_path=target.with_name(target.stem+'_확인필요.xlsx')
    review_wb=Workbook();s=review_wb.active;s.title='확인필요'
    names=['원본 PDF','문서번호','날짜','품명','규격','단위','반입','합격','불합격','검토 사유']
    s.append(names)
    for r in review:
        s.append([r['file'],r['doc'],r['date'],r['material'],r['spec'],r['unit'],r['qty'],r['passed'],r['failed'],' / '.join(r['issues'])])
    for c in s[1]:c.fill=PatternFill('solid',fgColor='23364F');c.font=Font(color='FFFFFF',bold=True);c.alignment=Alignment(wrap_text=True)
    for col,width in {'A':50,'B':21,'C':14,'D':32,'E':19,'F':10,'G':12,'H':12,'I':12,'J':55}.items():s.column_dimensions[col].width=width
    s.freeze_panes='D2';s.auto_filter.ref=s.dimensions
    review_wb.save(review_path)
    return counts,review_path

def cli_self_test(paths):
    find_tesseract()
    for path in paths:
        try:
            t=time.time();rows=parse_request(path,DEFAULT_NAMES)
            print(Path(path).name,'seconds',round(time.time()-t,1))
            for r in rows:
                print(r['doc'],r['date'],r['material'],r['spec'],r['unit'],r['qty'],r['passed'],r['failed'],'inspector=',r['inspector'],'issues=',r['issues'])
        except Exception as e:print('ERROR',path,str(e));traceback.print_exc()

def main_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    cfg=cfg_load();result=[];stop=threading.Event();busy=threading.Event()
    root=tk.Tk();root.title('자재검수대장 자동작성 '+VERSION+' (설치 불필요)');root.geometry('1000x690')
    root.option_add('*Font','맑은 고딕 10')
    folder=tk.StringVar(value=cfg.get('folder',''));names=tk.StringVar(value=cfg.get('names',DEFAULT_NAMES));
    existing=tk.StringVar();mode=tk.StringVar(value='new');output=tk.StringVar(value=cfg.get('output',''));status=tk.StringVar(value='준비 완료');counter=tk.StringVar(value='0 / 0건')
    wrap=ttk.Frame(root,padding=12);wrap.pack(fill='both',expand=True)
    def line(label,var,command):
        frame=ttk.Frame(wrap);frame.pack(fill='x',pady=5)
        ttk.Label(frame,text=label,width=18).pack(side='left')
        ttk.Entry(frame,textvariable=var).pack(side='left',fill='x',expand=True)
        ttk.Button(frame,text='선택',command=command).pack(side='left',padx=6)
    def save_pref():cfg_save({'folder':folder.get(),'names':names.get(),'output':output.get()})
    line('PDF 폴더',folder,lambda:folder.set(filedialog.askdirectory() or folder.get()))
    frame=ttk.Frame(wrap);frame.pack(fill='x',pady=6)
    ttk.Label(frame,text='담당감리원 명단',width=18).pack(side='left')
    ttk.Entry(frame,textvariable=names).pack(side='left',fill='x',expand=True)
    ttk.Label(wrap,text='명단은 자동 기억합니다. 이름으로 읽히지 않는 서명은 공란 처리합니다.',foreground='#555').pack(anchor='w')
    frame=ttk.Frame(wrap);frame.pack(fill='x',pady=6)
    ttk.Label(frame,text='대장 저장 방식',width=18).pack(side='left')
    ttk.Radiobutton(frame,text='새 대장 작성',variable=mode,value='new').pack(side='left')
    ttk.Radiobutton(frame,text='기존 대장에 추가',variable=mode,value='existing').pack(side='left',padx=20)
    line('기존 대장 (선택)',existing,lambda:existing.set(filedialog.askopenfilename(filetypes=[('매크로 엑셀','*.xlsm')]) or existing.get()))
    line('결과 저장 위치',output,lambda:output.set(filedialog.asksaveasfilename(defaultextension='.xlsm',filetypes=[('매크로 엑셀','*.xlsm')]) or output.get()))
    panel=ttk.Frame(wrap);panel.pack(fill='x',pady=10)
    progress=ttk.Progressbar(panel,maximum=100);progress.pack(fill='x')
    info=ttk.Frame(wrap);info.pack(fill='x')
    ttk.Label(info,textvariable=status).pack(side='left');ttk.Label(info,textvariable=counter).pack(side='right')
    columns=('doc','date','material','spec','qty','issues')
    table=ttk.Treeview(wrap,columns=columns,show='headings',height=13)
    for key,label,width in zip(columns,['문서번호','검수일자','자재명','규격','수량','검토사항'],[160,110,260,160,75,330]):
        table.heading(key,text=label);table.column(key,width=width)
    table.pack(fill='both',expand=True,pady=8)
    scroll=ttk.Scrollbar(table,orient='vertical',command=table.yview);table.configure(yscroll=scroll.set);scroll.pack(side='right',fill='y')
    def ui(fun):root.after(0,fun)
    def do_scan():
        if busy.is_set():return
        f=folder.get().strip()
        if not Path(f).is_dir():messagebox.showerror('폴더 오류','PDF 폴더가 존재하지 않습니다.');return
        save_pref();stop.clear();busy.set();result.clear();table.delete(*table.get_children());progress['value']=0
        paths=list(Path(f).rglob('*.pdf'));total=len(paths)
        if not total:
            busy.clear();messagebox.showwarning('PDF 없음','선택한 폴더에 PDF가 없습니다.');return
        def job():
            try:
                find_tesseract();start=time.time()
                for i,p in enumerate(paths,1):
                    if stop.is_set():break
                    try: rows=parse_request(p,names.get())
                    except Exception as e:
                        m=RE_DOC.search(p.name)
                        doc=f'자검({m[1]})-{m[2]}-{m[3]}' if m else p.stem
                        rows=[dict(file=str(p),doc=doc,date='',request='',contractor='',inspector='',area='',material='',spec='',unit='',qty='',passed='',failed='',issues=[str(e)])]
                    result.extend(rows)
                    def show(rows=rows,i=i):
                        for r in rows:
                            table.insert('', 'end', values=[r['doc'],r['date'],r['material'],r['spec'],r['qty'],' / '.join(r['issues'])])
                        progress['value']=100*i/total
                        counter.set(f'{i}/{total}건 · {time.time()-start:.0f}초 경과')
                    ui(show)
                ui(lambda:status.set('중단됨 · 지금까지 결과 저장 가능' if stop.is_set() else '분석 완료 · 엑셀 저장 가능'))
            except Exception as e:ui(lambda:messagebox.showerror('분석 오류',str(e)))
            finally:busy.clear()
        status.set('분석 중');threading.Thread(target=job,daemon=True).start()
    def save():
        if busy.is_set():messagebox.showwarning('처리 중','분석이 끝난 후 저장해 주세요.');return
        if not result:messagebox.showwarning('결과 없음','PDF 분석 결과가 없습니다.');return
        p=output.get().strip()
        if not p:messagebox.showwarning('저장 위치','저장할 XLSM 경로를 지정해 주세요.');return
        base=Path(getattr(sys,'_MEIPASS',Path(__file__).resolve().parent))/'base_template.xlsm'
        if mode.get()=='existing' and not Path(existing.get()).is_file():messagebox.showwarning('기존 파일','기존 XLSM 파일을 선택해 주세요.');return
        try:
            save_pref()
            counts,review=save_output(result,p,base,existing.get() if mode.get()=='existing' else None,mode.get()=='new')
            messagebox.showinfo('저장 완료',f"월별 대장: {counts['written']}행 입력\n중복: {counts['duplicate']}행 제외\n확인필요: {counts['review']}행\n\n확인 목록: {review}")
        except Exception as e:messagebox.showerror('저장 실패',str(e))
    bar=ttk.Frame(wrap);bar.pack(fill='x',pady=9)
    ttk.Button(bar,text='PDF 분석 시작 / 이어하기',command=do_scan).pack(side='left',padx=5)
    ttk.Button(bar,text='중단 (완료분 보존)',command=stop.set).pack(side='left',padx=5)
    ttk.Button(bar,text='현재 결과 엑셀 저장',command=save).pack(side='right',padx=5)
    root.protocol('WM_DELETE_WINDOW',lambda:(stop.set(),save_pref(),root.destroy()))
    root.mainloop()

if __name__=='__main__':
    arg=argparse.ArgumentParser();arg.add_argument('--self-test',nargs='*');ns=arg.parse_args()
    if ns.self_test is not None:cli_self_test(ns.self_test)
    else:main_gui()