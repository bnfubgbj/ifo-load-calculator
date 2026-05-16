import streamlit as st
import pdfplumber
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict
from datetime import datetime

st.set_page_config(page_title="คำนวณโหลดสินค้า IFO", page_icon="📦", layout="wide")

# ── CONSTANTS ──
BOX_CANVAS = 12
SACK_200 = 120
BOX_212 = 24

MONTH_MAP = {
    'ม.ค.':'01','ก.พ.':'02','มี.ค.':'03','เม.ย.':'04',
    'พ.ค.':'05','มิ.ย.':'06','ก.ค.':'07','ส.ค.':'08',
    'ก.ย.':'09','ต.ค.':'10','พ.ย.':'11','ธ.ค.':'12'
}

def clean(s):
    s = re.sub(r'\(cid:\d+\)', '', s)
    return re.sub(r'\s+', ' ', s).strip()

def th_date(s):
    if not s: return ''
    try:
        y,m,d = s.split('-')
        mn=['','ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.']
        return f"{int(d)} {mn[int(m)]} {int(y)+543}"
    except: return s

def detect_type(barcode):
    p = barcode[:3]
    if barcode[:2] == '11': return 'canvas'
    if p in ('122','123'): return 'foam212'
    if p in ('120','121'): return 'foam200'
    return 'gift'

def get_subtype(barcode, desc):
    if '205S' in desc: return '205S'
    if '205R' in desc: return '205R'
    if barcode[:3] == '123' or '213' in desc: return '213'
    if barcode[:3] == '122' or '212' in desc: return '212'
    if '200' in desc: return '200'
    return desc.split()[0] if desc else 'อื่นๆ'

def parse_pdf(file_bytes):
    # จัดกลุ่มหน้าตาม IFO
    groups = {}
    order = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            txt = clean(page.extract_text() or '')
            m = re.search(r'IFO-\d+', txt)
            if not m: continue
            ifo = m.group()
            if ifo not in groups:
                groups[ifo] = []
                order.append(ifo)
            groups[ifo].append(txt)

    docs = []
    for ifo in order:
        full_text = '\n'.join(groups[ifo])
        lines = [clean(l) for l in full_text.split('\n') if l.strip()]

        doc = {'docId': ifo, 'date': '', 'customer': '', 'province': '', 'amphoe': '', 'items': []}

        # date
        pat = r'(\d{1,2})\s+(' + '|'.join(re.escape(k) for k in MONTH_MAP) + r')\s+(\d{4})'
        m = re.search(pat, full_text)
        if m:
            d,mo,y = m.group(1),m.group(2),int(m.group(3))
            if y > 2500: y -= 543
            doc['date'] = f"{y}-{MONTH_MAP[mo]}-{d.zfill(2)}"

        # customer
        for line in lines[:20]:
            m = re.search(r'ช\S*\s+ล\S*\s+ค\S*\s*:\s*(.+?)(?:\s+ว\S*\s+ท|$)', line)
            if m:
                doc['customer'] = m.group(1).strip()
                break

        # province
        for line in lines[:15]:
            mp = re.search(r'จ\.([ก-๙]+(?:\s+[ก-๙]+)?)', line)
            ma = re.search(r'อ\.([ก-๙]+(?:\s+[ก-๙]+)?)(?:\s+จ\.)?', line)
            if mp: doc['province'] = re.sub(r'\d+.*','',mp.group(1)).strip()
            if ma: doc['amphoe'] = re.sub(r'\s*จ$','',re.sub(r'\d+.*','',ma.group(1)).strip()).strip()
            if doc['province']: break

        # items — รองรับทั้ง extract_text (บรรทัดยาว) และ extract_words
        seen = set()
        # แยก items จาก full_text โดยหา barcode 9 หลักตามด้วยจำนวนคู่
        for m in re.finditer(r'(\d{9})\s+(.+?)\s+(\d+)\s+คู่', full_text):
            bc, desc_raw, qty = m.group(1), m.group(2).strip(), int(m.group(3))
            if qty <= 0: continue
            if 'Z0001' in bc: continue
            # กรองเฉพาะ barcode สินค้าจริง (ขึ้นต้นด้วย 11 หรือ 12)
            if not re.match(r'^(11|12)', bc): continue
            key = (bc, qty)
            if key in seen: continue
            seen.add(key)
            # ตัดตัวเลขราคาท้ายออก
            desc = re.sub(r'\s+\d+(\.\d+)?(\s+\d+(\.\d+)?)*$','',desc_raw).strip()
            # ตัดเลขลำดับหน้าถ้ามี
            desc = re.sub(r'^\d+\s+','',desc).strip()
            ptype = detect_type(bc)
            subtype = get_subtype(bc, desc)
            doc['items'].append({'desc':desc,'type':ptype,'subtype':subtype,'qty':qty,'gift':False})

        if doc['items']:
            docs.append(doc)

    return docs

def calc_canvas(n): return {'qty':n,'boxes':n//BOX_CANVAS,'rem':n%BOX_CANVAS}
def calc_foam200(n):
    doz=n//12; sacks=doz//10; rem_doz=doz%10
    return {'qty':n,'doz':doz,'sacks':sacks,'rem_doz':rem_doz}
def calc_foam212(n): return {'qty':n,'boxes':n//BOX_212,'rem_doz':(n%BOX_212)//12}

def agg(doc):
    doc['_ct'] = sum(x['qty'] for x in doc['items'] if x['type']=='canvas')
    doc['_ft2'] = sum(x['qty'] for x in doc['items'] if x['type']=='foam200')
    doc['_ft3_212'] = sum(x['qty'] for x in doc['items'] if x['type']=='foam212' and x['subtype']=='212')
    doc['_ft3_213'] = sum(x['qty'] for x in doc['items'] if x['type']=='foam212' and x['subtype']=='213')

def build_excel(docs):
    wb = Workbook()
    BLUE='1B4F8A'; WHITE='FFFFFF'
    CG='D5E8D4'; CGH='2D6A27'
    CF='DBEAFE'; CFH='1E3A8A'
    C212='D0E4FF'; C212H='1E40AF'
    C213='E8D5F5'; C213H='5B21B6'
    CGIFT='FFF2CC'; CSUM='F0F0F0'
    thin=Side(style='thin',color='BBBBBB')
    bdr=Border(left=thin,right=thin,top=thin,bottom=thin)

    def hf(sz=10): return Font(name='TH Sarabun New',bold=True,color=WHITE,size=sz)
    def nf(sz=10,bold=False): return Font(name='TH Sarabun New',size=sz,bold=bold)
    def fl(c): return PatternFill('solid',start_color=c,end_color=c)
    def al(h='center'): return Alignment(horizontal=h,vertical='center',wrap_text=True)

    ws = wb.active; ws.title='สรุป'

    # หา subtypes จริงๆ
    canvas_subs = sorted(set(x['subtype'] for d in docs for x in d['items'] if x['type']=='canvas'))
    has212 = any(x['subtype']=='212' for d in docs for x in d['items'] if x['type']=='foam212')
    has213 = any(x['subtype']=='213' for d in docs for x in d['items'] if x['type']=='foam212')

    # build columns
    cols = ['วันที่','เลขที่IFO','ชื่อลูกค้า','อำเภอ','จังหวัด']
    col_fills = [BLUE]*5
    group_ranges = [(1,5,'ข้อมูล',BLUE)]
    col = 6
    for sub in canvas_subs:
        cols += ['คู่','กล่อง','เศษ']
        col_fills += [CGH]*3
        group_ranges.append((col,col+2,f'ผ้าใบ {sub} (12คู่/กล่อง)',CGH))
        col += 3
    cols += ['คู่','โหล','กระสอบ','เศษโหล']
    col_fills += [CFH]*4
    group_ranges.append((col,col+3,'ฟองน้ำ 200 (120คู่/กระสอบ)',CFH))
    col += 4
    if has212:
        cols += ['คู่','กล่อง','เศษ']
        col_fills += [C212H]*3
        group_ranges.append((col,col+2,'ฟองน้ำ 212 (24คู่/กล่อง)',C212H))
        col += 3
    if has213:
        cols += ['คู่','กล่อง','เศษ']
        col_fills += [C213H]*3
        group_ranges.append((col,col+2,'ฟองน้ำ 213 (24คู่/กล่อง)',C213H))
        col += 3
    cols += ['รวม']; col_fills += [BLUE]
    NCOLS = len(cols)

    # title
    ws.merge_cells(f'A1:{get_column_letter(NCOLS)}1')
    ws['A1'] = 'สรุปโหลดสินค้า — บริษัท นันยางมาร์เก็ตติ้ง จำกัด'
    ws['A1'].font=hf(13); ws['A1'].fill=fl(BLUE); ws['A1'].alignment=al()
    ws.row_dimensions[1].height=26

    ws.merge_cells(f'A2:{get_column_letter(NCOLS)}2')
    ws['A2'] = f'พิมพ์: {datetime.now().strftime("%d/%m/%Y %H:%M")}'
    ws['A2'].font=nf(9); ws['A2'].alignment=al('right')

    # group row
    for sc,ec,label,color in group_ranges:
        if sc==ec: ws.cell(row=3,column=sc,value=label)
        else: ws.merge_cells(start_row=3,start_column=sc,end_row=3,end_column=ec); ws.cell(row=3,column=sc,value=label)
        for c in range(sc,ec+1):
            cell=ws.cell(row=3,column=c)
            cell.font=hf(10); cell.fill=fl(color); cell.alignment=al(); cell.border=bdr
    ws.row_dimensions[3].height=20

    # sub headers
    for ci,(h,color) in enumerate(zip(cols,col_fills),1):
        c=ws.cell(row=4,column=ci,value=h)
        c.font=hf(10); c.fill=fl(color); c.alignment=al(); c.border=bdr
    ws.row_dimensions[4].height=32

    # data
    tot=defaultdict(int)
    for i,doc in enumerate(docs):
        r=5+i
        bg='FFFFFF' if i%2==0 else 'F5F8FA'
        canvas_qty=defaultdict(int)
        for x in doc['items']:
            if x['type']=='canvas': canvas_qty[x['subtype']]+=x['qty']
        ft2=doc['_ft2']; f212=doc['_ft3_212']; f213=doc['_ft3_213']
        grand=doc['_ct']+ft2+f212+f213
        tot['ft2']+=ft2; tot['f212']+=f212; tot['f213']+=f213
        for sub in canvas_subs: tot[f'c_{sub}']+=canvas_qty.get(sub,0)

        row_vals=[th_date(doc['date']),doc['docId'],doc['customer'],doc.get('amphoe',''),doc.get('province','')]
        row_bgs=[bg]*5
        for sub in canvas_subs:
            ct=canvas_qty.get(sub,0); cc=calc_canvas(ct)
            row_vals+=[ct or '-',cc['boxes'] or '-',cc['rem'] or '-']
            row_bgs+=[CG if ct else bg]*3
        cf2=calc_foam200(ft2)
        row_vals+=[ft2 or '-',cf2['doz'] or '-',cf2['sacks'] or '-',cf2['rem_doz'] or '-']
        row_bgs+=[CF if ft2 else bg]*4
        if has212:
            c212=calc_foam212(f212)
            row_vals+=[f212 or '-',c212['boxes'] or '-',c212['rem_doz'] or '-']
            row_bgs+=[C212 if f212 else bg]*3
        if has213:
            c213=calc_foam212(f213)
            row_vals+=[f213 or '-',c213['boxes'] or '-',c213['rem_doz'] or '-']
            row_bgs+=[C213 if f213 else bg]*3
        row_vals+=[grand]; row_bgs+=['D6E4F0']

        for ci,(val,bgc) in enumerate(zip(row_vals,row_bgs),1):
            c=ws.cell(row=r,column=ci,value=val)
            c.fill=fl(bgc); c.font=nf(10,bold=(ci==NCOLS))
            c.alignment=al('left') if ci<=3 else al(); c.border=bdr
        ws.row_dimensions[r].height=22

    # summary
    sr=5+len(docs)
    ws.merge_cells(start_row=sr,start_column=1,end_row=sr,end_column=5)
    sum_vals={1:'รวมทั้งหมด'}
    ci=6
    for sub in canvas_subs:
        cc=calc_canvas(tot[f'c_{sub}'])
        sum_vals[ci]=tot[f'c_{sub}'] or '-'; sum_vals[ci+1]=cc['boxes'] or '-'; sum_vals[ci+2]=cc['rem'] or '-'
        ci+=3
    cf2=calc_foam200(tot['ft2'])
    sum_vals[ci]=tot['ft2'] or '-'; sum_vals[ci+1]=cf2['doz'] or '-'
    sum_vals[ci+2]=cf2['sacks'] or '-'; sum_vals[ci+3]=cf2['rem_doz'] or '-'; ci+=4
    if has212:
        c212=calc_foam212(tot['f212'])
        sum_vals[ci]=tot['f212'] or '-'; sum_vals[ci+1]=c212['boxes'] or '-'; sum_vals[ci+2]=c212['rem_doz'] or '-'; ci+=3
    if has213:
        c213=calc_foam212(tot['f213'])
        sum_vals[ci]=tot['f213'] or '-'; sum_vals[ci+1]=c213['boxes'] or '-'; sum_vals[ci+2]=c213['rem_doz'] or '-'; ci+=3
    sum_vals[NCOLS]=sum(tot.values())
    for ci in range(1,NCOLS+1):
        c=ws.cell(row=sr,column=ci)
        if ci in sum_vals: c.value=sum_vals[ci]
        c.font=hf(10); c.fill=fl(BLUE); c.alignment=al(); c.border=bdr
    ws.row_dimensions[sr].height=24

    # widths
    base_w=[12,14,24,12,14]
    for _ in canvas_subs: base_w+=[8,8,8]
    base_w+=[8,8,10,10]
    if has212: base_w+=[8,8,8]
    if has213: base_w+=[8,8,8]
    base_w+=[10]
    for ci,w in enumerate(base_w,1):
        ws.column_dimensions[get_column_letter(ci)].width=w

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ── UI ──
st.markdown("""
<div style="background:linear-gradient(135deg,#1B4F8A,#163D6E);color:white;border-radius:12px;padding:16px 24px;margin-bottom:20px">
<h1 style="font-size:20px;font-weight:700;margin:0">📦 คำนวณโหลดสินค้า IFO</h1>
<p style="font-size:13px;opacity:.7;margin:4px 0 0">อัปโหลด PDF → คำนวณอัตโนมัติ → ดาวน์โหลด Excel</p>
</div>""", unsafe_allow_html=True)

if 'ukey' not in st.session_state: st.session_state.ukey=0

c1,c2=st.columns([4,1])
with c1:
    uploaded=st.file_uploader("📂 ลาก PDF มาวาง (รองรับ PDF รวมหลาย IFO)",
        type=['pdf'],accept_multiple_files=True,key=f"up_{st.session_state.ukey}")
with c2:
    st.markdown('<br>',unsafe_allow_html=True)
    if st.button("🗑️ ล้างข้อมูล",use_container_width=True):
        st.session_state.ukey+=1; st.rerun()

if uploaded:
    docs=[]; errors=[]
    with st.spinner('🔍 กำลังอ่าน PDF...'):
        for f in uploaded:
            try:
                results=parse_pdf(f.read())
                for doc in results:
                    agg(doc); doc['_file']=f.name; docs.append(doc)
            except Exception as e:
                errors.append(f"{f.name}: {e}")

    for e in errors: st.error(f"❌ {e}")

    if docs:
        docs.sort(key=lambda x:x['docId'])
        tot_c=sum(d['_ct'] for d in docs)
        tot_f=sum(d['_ft2'] for d in docs)
        tot_212=sum(d['_ft3_212'] for d in docs)
        tot_213=sum(d['_ft3_213'] for d in docs)

        # metrics
        canvas_subs_all=defaultdict(int)
        for d in docs:
            for x in d['items']:
                if x['type']=='canvas': canvas_subs_all[x['subtype']]+=x['qty']

        metric_items=[("📄 เอกสาร",f"{len(docs)} IFO")]
        for sub,qty in sorted(canvas_subs_all.items()):
            metric_items.append((f"🟢 ผ้าใบ {sub}",f"{qty} คู่"))
        metric_items.append(("🔵 ฟองน้ำ 200",f"{tot_f} คู่"))
        if tot_212: metric_items.append(("🔵 ฟองน้ำ 212",f"{tot_212} คู่"))
        if tot_213: metric_items.append(("🔵 ฟองน้ำ 213",f"{tot_213} คู่"))

        cols=st.columns(len(metric_items))
        for i,(label,val) in enumerate(metric_items):
            cols[i].metric(label,val)

        st.divider()

        for doc in docs:
            with st.expander(f"📄 {doc['docId']} — {doc['customer']} ({th_date(doc['date'])})",expanded=True):
                c1,c2,c3=st.columns(3)
                c1.write(f"**ลูกค้า:** {doc['customer']}")
                c2.write(f"**อ.{doc.get('amphoe','')} จ.{doc.get('province','')}**")
                c3.write(f"**รวม:** {doc['_ct']+doc['_ft2']+doc['_ft3_212']+doc['_ft3_213']} คู่")
                if doc['_ct']:
                    cc=calc_canvas(doc['_ct'])
                    st.success(f"🟢 **ผ้าใบ:** {doc['_ct']} คู่ / {cc['boxes']} กล่อง" + (f" เศษ {cc['rem']} คู่" if cc['rem'] else ""))
                if doc['_ft2']:
                    cf=calc_foam200(doc['_ft2'])
                    st.info(f"🔵 **ฟองน้ำ 200:** {doc['_ft2']} คู่ / {cf['sacks']} กระสอบ เศษ {cf['rem_doz']} โหล")
                if doc['_ft3_212']:
                    c2=calc_foam212(doc['_ft3_212'])
                    st.info(f"🔵 **ฟองน้ำ 212:** {doc['_ft3_212']} คู่ / {c2['boxes']} กล่อง")
                if doc['_ft3_213']:
                    c3=calc_foam212(doc['_ft3_213'])
                    st.info(f"🔵 **ฟองน้ำ 213:** {doc['_ft3_213']} คู่ / {c3['boxes']} กล่อง")

        excel=build_excel(docs)
        fname=f"IFO_โหลดสินค้า_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        st.download_button("📥 ดาวน์โหลด Excel",data=excel,file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,type="primary")
    else:
        if not errors: st.warning("⚠️ ไม่พบข้อมูล IFO ในไฟล์ที่อัปโหลด")
