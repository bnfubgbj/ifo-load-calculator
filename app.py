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

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Sarabun', sans-serif; }
.header { background: linear-gradient(135deg,#1B4F8A,#163D6E); color:white; border-radius:12px; padding:18px 24px; margin-bottom:20px; }
.header h1 { font-size:22px; font-weight:700; margin:0; }
.header p  { font-size:13px; opacity:.7; margin:4px 0 0; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="header">
  <h1>📦 คำนวณโหลดสินค้า IFO</h1>
  <p>อัปโหลด PDF → คำนวณอัตโนมัติ → ดาวน์โหลด Excel</p>
</div>
""", unsafe_allow_html=True)

# ── CONSTANTS ─────────────────────────────────────────
BOX_CANVAS   = 12   # ผ้าใบ: 12 คู่/กล่อง
SACK_FOAM200 = 120  # ฟองน้ำ 200: 120 คู่/กระสอบ
PACK_FOAM212 = 24   # ฟองน้ำ 212/213: 24 คู่/กล่อง
DOZ          = 12   # 1 โหล = 12 คู่

# ── TEXT HELPERS ──────────────────────────────────────
def clean_cid(s):
    s = re.sub(r'\(cid:\d+\)', '', s)
    return re.sub(r'\s+', ' ', s).strip()

def normalize_thai(s):
    fixes = [
        (r'พรพมิ\s*ล','พรพิมล'),(r'จอมแจง้','จอมแจ้ง'),
        (r'บารเ์\s*บอร์?','บาร์เบอร์'),(r'พศิ\s*ษิ\s*ฐ์?','พิษฐ์'),
        (r'รา้น','ร้าน'),(r'เพชรบรุ\s*ี','เพชรบุรี'),
        (r'สระบรุ\s*ี','สระบุรี'),(r'แมฮ่\s*อ่\s*งสอน','แม่ฮ่องสอน'),
        (r'แมฮ่\s*อ่','แม่ฮ่องสอน'),(r'ทา่\s*ยาง','ท่ายาง'),
        (r'วหิ\s*ารแดง','วิหารแดง'),(r'ฟองนา้','ฟองน้ำ'),
        (r'ฟองนํา','ฟองน้ำ'),(r'ผา้\s*ใบ','ผ้าใบ'),
        (r'นา้\s*ตาล','น้ำตาล'),(r'นา้\s*เงนิ','น้ำเงิน'),
        (r'หนา้\s*ขาว','หน้าขาว'),(r'ลว้น','ล้วน'),
        (r'เขม้','เข้ม'),(r'ออ่น','อ่อน'),(r'พเิ\s*ศษ','พิเศษ'),
        (r'ดํา','ดำ'),(r'นํา','น้ำ'),(r'์$',''),(r'\s+',' '),
    ]
    for p,r in fixes: s = re.sub(p,r,s)
    return s.strip()

# ── PDF PARSE ─────────────────────────────────────────
def extract_lines(file_bytes):
    lines = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            words = page.extract_words(x_tolerance=3, y_tolerance=3)
            rows = {}
            for w in words:
                y = round(w['top']/5)*5
                rows.setdefault(y,[]).append(w['text'])
            for y in sorted(rows):
                l = clean_cid(' '.join(rows[y]))
                if l: lines.append(l)
    return lines

def detect_product_type(barcode, desc):
    """แยกประเภทสินค้าจาก barcode prefix"""
    pre3 = barcode[:3]
    pre2 = barcode[:2]
    # ผ้าใบ: 110, 111, 112, 113, 114 (รวม 205S/R, Zafari, 121Z)
    if pre2 in ('11',):  return 'canvas'
    if pre3 in ('110','111','112','113','114'): return 'canvas'
    # ฟองน้ำ 212: 122
    if pre3 == '122': return 'foam212'
    # ฟองน้ำ 213: 123
    if pre3 == '123': return 'foam212'
    # ฟองน้ำ 200: 120, 121
    if pre3 in ('120','121'): return 'foam200'
    # fallback จาก desc
    d = desc
    if re.search(r'ผ้าใบ|205[SR]|zafari|ซาฟารี', d, re.I): return 'canvas'
    if re.search(r'212|213', d): return 'foam212'
    if re.search(r'ฟองน้ำ|200', d): return 'foam200'
    return 'gift'

def get_product_subtype(barcode, desc):
    """ดึงชื่อรุ่นสินค้า"""
    d = normalize_thai(desc)
    pre3 = barcode[:3]
    if '205S' in d: return '205S'
    if '205R' in d: return '205R'
    if re.search(r'zafari|ซาฟารี|121Z', d, re.I): return 'Zafari'
    if pre3 == '123' or '213' in d: return '213'
    if pre3 == '122' or '212' in d: return '212'
    if pre3 in ('120','121') or '200' in d: return '200'
    # ผ้าใบอื่น
    m = re.search(r'(\d{3}[SR]?)', d)
    if m: return m.group(1)
    return d.split()[0] if d else 'อื่นๆ'

def is_gift(barcode, desc, price_str=''):
    """ตรวจว่าเป็นของแถมมั้ย — ราคา 0 หรือมีคำว่าแถม"""
    if re.search(r'แถม|gift|free', desc, re.I): return True
    return False

MONTH_MAP = {'ม.ค.':'01','ก.พ.':'02','มี.ค.':'03','เม.ย.':'04',
             'พ.ค.':'05','มิ.ย.':'06','ก.ค.':'07','ส.ค.':'08',
             'ก.ย.':'09','ต.ค.':'10','พ.ย.':'11','ธ.ค.':'12'}

def parse_pdf(file_bytes):
    result = {'docId':'','date':'','customer':'','province':'','amphoe':'','items':[]}
    lines = extract_lines(file_bytes)
    text = '\n'.join(lines)

    # docId
    m = re.search(r'IFO-\d+', text)
    if m: result['docId'] = m.group()

    # date
    pattern = r'(\d{1,2})\s+(' + '|'.join(re.escape(k) for k in MONTH_MAP) + r')\s+(\d{4})'
    m = re.search(pattern, text)
    if m:
        d,mo,y = m.group(1),m.group(2),int(m.group(3))
        if y > 2500: y -= 543
        result['date'] = f"{y}-{MONTH_MAP[mo]}-{d.zfill(2)}"

    # customer — line 8 รูปแบบ "ชอื ลกู คา้ : <ชื่อ>"
    for line in lines[:15]:
        # match "ชอื ลกู คา้ :" หรือ "ชื่อลูกค้า :"
        m = re.search(r'ช(?:อื|ื่อ)\s+ล(?:กู|ูก)\s+ค(?:า้|้า)\s*:\s*(.+?)(?:\s+ว(?:นั|ัน)\s*ท|$)', line)
        if not m:
            m = re.search(r'ช\S*\s+ล\S*\s+ค\S*\s*:\s*(.+?)(?:\s+ว\S*\s+ท|$)', line)
        if m:
            name = normalize_thai(m.group(1).strip())
            if len(name) > 1:
                result['customer'] = name
                break

    # province & amphoe
    for line in lines[:15]:
        mp = re.search(r'จ\.([ก-๙]+(?:\s+[ก-๙]+)?)', line)
        ma = re.search(r'อ\.([ก-๙]+(?:\s+[ก-๙]+)?)(?:\s+จ\.)?', line)
        if mp: result['province'] = normalize_thai(re.sub(r'\d+.*','',mp.group(1)).strip())
        if ma:
            amp = normalize_thai(re.sub(r'\d+.*','',ma.group(1)).strip())
            result['amphoe'] = re.sub(r'\s*จ$','',amp).strip()
        if result['province']: break

    # items
    for line in lines:
        if 'Z0001' in line or 'มัดจำ' in line: continue
        m = re.search(r'(\d{9})\s+(.+?)\s+(\d+)\s+คู่', line)
        if m:
            barcode = m.group(1)
            desc_raw = m.group(2).strip()
            qty = int(m.group(3))
            if qty <= 0: continue
            desc = normalize_thai(re.sub(r'\s+\d+(\.\d+)?(\s+\d+(\.\d+)?)*$','',desc_raw).strip())
            ptype = detect_product_type(barcode, desc)
            subtype = get_product_subtype(barcode, desc)
            gift = is_gift(barcode, desc)
            result['items'].append({'desc':desc,'type':ptype,'subtype':subtype,'qty':qty,'gift':gift})
    return result

# ── CALC HELPERS ──────────────────────────────────────

def th_date(s):
    if not s: return ''
    try:
        y,m,d = s.split('-')
        mn=['','ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.']
        return f"{int(d)} {mn[int(m)]} {int(y)+543}"
    except: return s

def _aggregate(doc):
    doc['_canvas']    = [x for x in doc['items'] if x['type']=='canvas'  and not x['gift']]
    doc['_foam200']   = [x for x in doc['items'] if x['type']=='foam200' and not x['gift']]
    doc['_foam212']   = [x for x in doc['items'] if x['type']=='foam212' and not x['gift']]
    doc['_gift_c']    = [x for x in doc['items'] if x['type']=='canvas'  and x['gift']]
    doc['_gift_f200'] = [x for x in doc['items'] if x['type']=='foam200' and x['gift']]
    doc['_gift_f212'] = [x for x in doc['items'] if x['type']=='foam212' and x['gift']]
    doc['_ct']   = sum(x['qty'] for x in doc['_canvas'])
    doc['_ft2']  = sum(x['qty'] for x in doc['_foam200'])
    doc['_ft3']  = sum(x['qty'] for x in doc['_foam212'])
    doc['_gct']  = sum(x['qty'] for x in doc['_gift_c'])
    doc['_gft2'] = sum(x['qty'] for x in doc['_gift_f200'])
    doc['_gft3'] = sum(x['qty'] for x in doc['_gift_f212'])

def calc_canvas(n):
    boxes, rem = divmod(n, BOX_CANVAS)
    return {'qty':n,'boxes':boxes,'rem':rem}

def calc_foam200(n):
    doz_total = n // DOZ
    rem_pairs = n % DOZ
    sacks, rem_doz = divmod(doz_total, 10)  # 10 โหล = 1 กระสอบ
    return {'qty':n,'doz':doz_total,'sacks':sacks,'rem_doz':rem_doz,'rem_pairs':rem_pairs}

def calc_foam212(n):
    packs, rem = divmod(n, PACK_FOAM212)
    rem_doz, rem_pairs = divmod(rem, DOZ)
    return {'qty':n,'packs':packs,'rem_doz':rem_doz,'rem_pairs':rem_pairs}

def fmt_canvas(c):
    s = f"{c['qty']} คู่ / {c['boxes']} กล่อง"
    if c['rem']: s += f" เศษ {c['rem']} คู่"
    return s

def fmt_foam200(f):
    s = f"{f['qty']} คู่ / {f['sacks']} กระสอบ"
    if f['rem_doz']: s += f" เศษ {f['rem_doz']} โหล"
    return s

def fmt_foam212(f):
    s = f"{f['qty']} คู่ / {f['packs']} กล่อง"
    if f['rem_doz']: s += f" เศษ {f['rem_doz']} โหล"
    return s

# ── BUILD EXCEL ───────────────────────────────────────
def build_excel(docs):
    wb = Workbook()

    # colors
    C_BLUE='1B4F8A'; C_WHITE='FFFFFF'
    C_CANVAS='D5E8D4'; C_CANVAS_H='2D6A27'
    C_FOAM200='DAE8FC'; C_FOAM200_H='1E3A8A'
    C_FOAM212='D0E4FF'; C_FOAM212_H='1E40AF'
    C_GIFT='FFF2CC'; C_GIFT_H='7D4E00'
    C_SUM='1B4F8A'; C_ALT='F5F8FA'

    thin = Side(style='thin', color='BBBBBB')
    bdr  = Border(left=thin,right=thin,top=thin,bottom=thin)

    def hf(sz=11,color=C_WHITE,bold=True):
        return Font(name='TH Sarabun New',bold=bold,color=color,size=sz)
    def nf(sz=11,bold=False,color='1E293B'):
        return Font(name='TH Sarabun New',bold=bold,size=sz,color=color)
    def fl(c): return PatternFill('solid',start_color=c,end_color=c)
    def al(h='center',v='center',wrap=True):
        return Alignment(horizontal=h,vertical=v,wrap_text=wrap)

    def write(ws,row,col,val,font=None,fill=None,align=None,border=None):
        c = ws.cell(row=row,column=col,value=val)
        if font:   c.font   = font
        if fill:   c.fill   = fill
        if align:  c.alignment = align
        if border: c.border = border
        return c

    def merge_write(ws,r1,c1,r2,c2,val,font=None,fill=None,align=None):
        ws.merge_cells(start_row=r1,start_column=c1,end_row=r2,end_column=c2)
        c = ws.cell(row=r1,column=c1,value=val)
        if font:  c.font  = font
        if fill:  c.fill  = fill
        if align: c.alignment = align
        # borders on all cells
        thin2 = Side(style='thin',color='BBBBBB')
        b = Border(left=thin2,right=thin2,top=thin2,bottom=thin2)
        for rr in range(r1,r2+1):
            for cc in range(c1,c2+1):
                ws.cell(row=rr,column=cc).border = b

    # ── aggregate data ──
    for doc in docs:
        _aggregate(doc)

    def th_date(s):
        if not s: return ''
        try:
            y,m,d = s.split('-')
            mn=['','ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.']
            return f"{int(d)} {mn[int(m)]} {int(y)+543}"
        except: return s

    # ════════════════════════════════════════════════
    # SHEET 1: สรุปรายเอกสาร
    # ════════════════════════════════════════════════
    ws = wb.active
    ws.title = 'สรุปรายเอกสาร'

    # col layout: A=วันที่ B=IFO C=ลูกค้า D=อ. E=จ.
    # F=ผ้าใบคู่ G=กล่อง H=เศษ  I=ของแถมผ้าใบคู่ J=กล่อง K=เศษ
    # L=ฟองน้ำ200คู่ M=โหล N=กระสอบ O=เศษโหล  P=ของแถม200คู่ Q=กระสอบ R=เศษโหล
    # S=ฟองน้ำ212/213คู่ T=กล่อง U=เศษโหล  V=ของแถม212คู่ W=กล่อง
    # X=รวม
    NCOLS = 24

    # title
    merge_write(ws,1,1,1,NCOLS,
        'สรุปโหลดสินค้า — บริษัท นันยางมาร์เก็ตติ้ง จำกัด',
        font=hf(14),fill=fl(C_BLUE),align=al())
    ws.row_dimensions[1].height = 30

    merge_write(ws,2,1,2,NCOLS,
        f'พิมพ์: {datetime.now().strftime("%d/%m/%Y %H:%M")}',
        font=nf(10),align=al('right'))
    ws.row_dimensions[2].height = 16

    # group row 3
    groups = [
        (1,5,'ข้อมูลเอกสาร',C_BLUE),
        (6,11,'ผ้าใบ (12 คู่/กล่อง)',C_CANVAS_H),
        (12,18,'ฟองน้ำ 200 (120 คู่/กระสอบ)',C_FOAM200_H),
        (19,23,'ฟองน้ำ 212/213 (24 คู่/กล่อง)',C_FOAM212_H),
        (24,24,'รวม',C_BLUE),
    ]
    for sc,ec,label,color in groups:
        merge_write(ws,3,sc,3,ec,label,
            font=hf(11),fill=fl(color),align=al())
    ws.row_dimensions[3].height = 22

    # sub headers row 4
    sub = ['วันที่','เลขที่IFO','ชื่อลูกค้า','อำเภอ','จังหวัด',
           'คู่','กล่อง','เศษคู่','ของแถม คู่','กล่อง','เศษคู่',
           'คู่','โหล','กระสอบ','เศษโหล','ของแถม คู่','กระสอบ','เศษโหล',
           'คู่','กล่อง','เศษโหล','ของแถม คู่','กล่อง',
           'รวมคู่']
    sub_colors = [C_BLUE]*5 + [C_CANVAS_H]*6 + [C_FOAM200_H]*7 + [C_FOAM212_H]*5 + [C_BLUE]
    for col,(h,color) in enumerate(zip(sub,sub_colors),1):
        write(ws,4,col,h,font=hf(10),fill=fl(color),align=al(),border=bdr)
    ws.row_dimensions[4].height = 36

    # data rows
    tot = defaultdict(int)
    for i,doc in enumerate(docs):
        r = 5+i
        ct,ft2,ft3 = doc['_ct'],doc['_ft2'],doc['_ft3']
        gct,gft2,gft3 = doc['_gct'],doc['_gft2'],doc['_gft3']
        cc=calc_canvas(ct); cf2=calc_foam200(ft2); cf3=calc_foam212(ft3)
        gcc=calc_canvas(gct); gcf2=calc_foam200(gft2); gcf3=calc_foam212(gft3)
        tot['ct']+=ct; tot['ft2']+=ft2; tot['ft3']+=ft3
        tot['gct']+=gct; tot['gft2']+=gft2; tot['gft3']+=gft3

        bg = 'FFFFFF' if i%2==0 else C_ALT
        row_vals = [
            th_date(doc['date']),doc['docId'],doc['customer'],
            doc.get('amphoe',''),doc.get('province',''),
            ct or '-',cc['boxes'] or '-',cc['rem'] or '-',
            gct or '-',gcc['boxes'] or '-',gcc['rem'] or '-',
            ft2 or '-',cf2['doz'] or '-',cf2['sacks'] or '-',cf2['rem_doz'] or '-',
            gft2 or '-',gcf2['sacks'] or '-',gcf2['rem_doz'] or '-',
            ft3 or '-',cf3['packs'] or '-',cf3['rem_doz'] or '-',
            gft3 or '-',gcf3['packs'] or '-',
            ct+ft2+ft3+gct+gft2+gft3
        ]
        row_bgs = [bg]*5 + \
            [C_CANVAS if ct else bg]*3 + [C_GIFT if gct else bg]*3 + \
            [C_FOAM200 if ft2 else bg]*4 + [C_GIFT if gft2 else bg]*3 + \
            [C_FOAM212 if ft3 else bg]*3 + [C_GIFT if gft3 else bg]*2 + \
            [C_ALT]
        for col,(val,bgc) in enumerate(zip(row_vals,row_bgs),1):
            write(ws,r,col,val,
                font=nf(11,bold=(col==NCOLS)),
                fill=fl(bgc),
                align=al('left' if col<=3 else 'center'),
                border=bdr)
        ws.row_dimensions[r].height = 22

    # summary row
    sr = 5+len(docs)
    merge_write(ws,sr,1,sr,5,'รวมทั้งหมด',font=hf(11),fill=fl(C_SUM),align=al())
    cc_t=calc_canvas(tot['ct']); cf2_t=calc_foam200(tot['ft2']); cf3_t=calc_foam212(tot['ft3'])
    gcc_t=calc_canvas(tot['gct']); gcf2_t=calc_foam200(tot['gft2']); gcf3_t=calc_foam212(tot['gft3'])
    sum_vals = {
        6:tot['ct'] or '-', 7:cc_t['boxes'] or '-', 8:cc_t['rem'] or '-',
        9:tot['gct'] or '-', 10:gcc_t['boxes'] or '-', 11:gcc_t['rem'] or '-',
        12:tot['ft2'] or '-', 13:cf2_t['doz'] or '-', 14:cf2_t['sacks'] or '-', 15:cf2_t['rem_doz'] or '-',
        16:tot['gft2'] or '-', 17:gcf2_t['sacks'] or '-', 18:gcf2_t['rem_doz'] or '-',
        19:tot['ft3'] or '-', 20:cf3_t['packs'] or '-', 21:cf3_t['rem_doz'] or '-',
        22:tot['gft3'] or '-', 23:gcf3_t['packs'] or '-',
        24:sum(tot.values())
    }
    for col in range(6,NCOLS+1):
        write(ws,sr,col,sum_vals.get(col,''),
            font=hf(11),fill=fl(C_SUM),align=al('center'),border=bdr)
    ws.row_dimensions[sr].height = 24

    # col widths
    widths = [12,14,24,12,14, 8,8,8, 10,8,8, 8,8,10,10, 10,10,10, 8,8,10, 10,8, 10]
    for col,w in enumerate(widths,1):
        ws.column_dimensions[get_column_letter(col)].width = w

    # ════════════════════════════════════════════════
    # SHEET 2: สรุปรวมทุก IFO
    # ════════════════════════════════════════════════
    ws2 = wb.create_sheet('สรุปรวม')

    # grand totals
    grand_ct  = sum(d['_ct']   for d in docs)
    grand_ft2 = sum(d['_ft2']  for d in docs)
    grand_ft3 = sum(d['_ft3']  for d in docs)
    grand_gct = sum(d['_gct']  for d in docs)
    grand_gft2= sum(d['_gft2'] for d in docs)
    grand_gft3= sum(d['_gft3'] for d in docs)

    gc=calc_canvas(grand_ct); gf2=calc_foam200(grand_ft2); gf3=calc_foam212(grand_ft3)
    ggc=calc_canvas(grand_gct); ggf2=calc_foam200(grand_gft2); ggf3=calc_foam212(grand_gft3)

    merge_write(ws2,1,1,1,8,
        f'สรุปรวมทุกเอกสาร ({len(docs)} IFO) — {datetime.now().strftime("%d/%m/%Y %H:%M")}',
        font=hf(13),fill=fl(C_BLUE),align=al())
    ws2.row_dimensions[1].height = 28

    # grand summary boxes row 3-4
    r = 3
    summary_blocks = []
    if grand_ct:   summary_blocks.append(('ผ้าใบ (ปกติ)',fmt_canvas(gc),C_CANVAS_H))
    if grand_gct:  summary_blocks.append(('ผ้าใบ (ของแถม)',fmt_canvas(ggc),C_GIFT_H))
    if grand_ft2:  summary_blocks.append(('ฟองน้ำ 200 (ปกติ)',fmt_foam200(gf2),C_FOAM200_H))
    if grand_gft2: summary_blocks.append(('ฟองน้ำ 200 (ของแถม)',fmt_foam200(ggf2),C_GIFT_H))
    if grand_ft3:  summary_blocks.append(('ฟองน้ำ 212/213 (ปกติ)',fmt_foam212(gf3),C_FOAM212_H))
    if grand_gft3: summary_blocks.append(('ฟองน้ำ 212/213 (ของแถม)',fmt_foam212(ggf3),C_GIFT_H))

    for idx,(label,val,color) in enumerate(summary_blocks):
        col = 1 + idx*2
        merge_write(ws2,3,col,3,col+1,label,font=hf(11),fill=fl(color),align=al())
        merge_write(ws2,4,col,4,col+1,val,font=nf(12,bold=True),fill=fl('F8FAFC'),align=al())
        ws2.column_dimensions[get_column_letter(col)].width = 14
        ws2.column_dimensions[get_column_letter(col+1)].width = 14
    ws2.row_dimensions[3].height = 22
    ws2.row_dimensions[4].height = 28

    # ── breakdown by subtype ──
    r = 6
    merge_write(ws2,r,1,r,8,'รายละเอียดตามชนิดสินค้า (รวมทุก IFO)',
        font=hf(12),fill=fl(C_BLUE),align=al())
    ws2.row_dimensions[r].height = 22

    r += 1
    det_headers = ['ชนิดสินค้า','ประเภท','รวม คู่','โหล','กล่อง/กระสอบ/กล่อง','เศษโหล','เศษคู่','หน่วย']
    det_colors  = [C_BLUE]*8
    for col,(h,color) in enumerate(zip(det_headers,det_colors),1):
        write(ws2,r,col,h,font=hf(11),fill=fl(color),align=al(),border=bdr)
    ws2.row_dimensions[r].height = 22
    r += 1

    # collect breakdown
    subtype_data = defaultdict(lambda: defaultdict(int))
    for doc in docs:
        for item in doc['items']:
            key = (item['subtype'], item['type'], item['gift'])
            subtype_data[key]['qty'] += item['qty']

    # sort: ปกติก่อน → ของแถมทีหลัง, แยกกลุ่ม canvas/foam200/foam212
    type_order = {'canvas':0,'foam200':1,'foam212':2}
    sorted_keys = sorted(subtype_data.keys(),
        key=lambda x: (type_order.get(x[1],9), int(x[2]), x[0]))

    last_gift = None
    for key in sorted_keys:
        subtype, ptype, gift = key
        qty = subtype_data[key]['qty']
        if qty == 0: continue

        # เว้นบรรทัดก่อนเริ่มของแถม
        if gift and last_gift == False:
            ws2.row_dimensions[r].height = 8
            r += 1

        last_gift = gift

        if ptype == 'canvas':
            c = calc_canvas(qty)
            unit = 'กล่อง'; pack_val = c['boxes']; rem_doz = '-'; rem_pair = c['rem'] or '-'
            color = C_GIFT if gift else C_CANVAS
        elif ptype == 'foam200':
            c = calc_foam200(qty)
            unit = 'กระสอบ'; pack_val = c['sacks']; rem_doz = c['rem_doz'] or '-'
            rem_pair = c['rem_pairs'] or '-'
            color = C_GIFT if gift else C_FOAM200
        else:
            c = calc_foam212(qty)
            unit = 'กล่อง'; pack_val = c['packs']; rem_doz = c['rem_doz'] or '-'
            rem_pair = c['rem_pairs'] or '-'
            color = C_GIFT if gift else C_FOAM212

        doz_total = qty // DOZ
        row_vals = [
            subtype,
            'ของแถม' if gift else 'ปกติ',
            qty, doz_total, pack_val, rem_doz, rem_pair, unit
        ]
        for col,val in enumerate(row_vals,1):
            write(ws2,r,col,val,
                font=nf(11,bold=(col==1)),
                fill=fl(color),
                align=al('left' if col<=2 else 'center'),
                border=bdr)
        ws2.row_dimensions[r].height = 20
        r += 1

    # col widths sheet2
    for col,w in enumerate([18,10,10,8,16,10,10,10],1):
        ws2.column_dimensions[get_column_letter(col)].width = w

    # ════════════════════════════════════════════════
    # SHEET 3: รายละเอียดทุกรายการ
    # ════════════════════════════════════════════════
    ws3 = wb.create_sheet('รายละเอียด')
    merge_write(ws3,1,1,1,8,'รายละเอียดสินค้าทุกรายการ',
        font=hf(13),fill=fl(C_BLUE),align=al())
    ws3.row_dimensions[1].height = 24

    dh = ['วันที่','เลขที่IFO','ชื่อลูกค้า','รายการสินค้า','ประเภท','คู่','โหล','กล่อง/กระสอบ/กล่อง']
    for col,h in enumerate(dh,1):
        write(ws3,2,col,h,font=hf(10),fill=fl(C_BLUE),align=al(),border=bdr)
    ws3.row_dimensions[2].height = 20

    TYPE_LABEL = {'canvas':'ผ้าใบ','foam200':'ฟองน้ำ 200','foam212':'ฟองน้ำ 212/213'}
    TYPE_COLOR = {'canvas':C_CANVAS,'foam200':C_FOAM200,'foam212':C_FOAM212}

    r3 = 3
    for doc in docs:
        for item in doc['items']:
            ptype = item['type']
            qty = item['qty']
            doz = qty // DOZ
            if ptype == 'canvas':
                c = calc_canvas(qty)
                load = f"{c['boxes']} กล่อง" + (f" เศษ {c['rem']} คู่" if c['rem'] else '')
            elif ptype == 'foam200':
                c = calc_foam200(qty)
                load = f"{c['sacks']} กระสอบ" + (f" เศษ {c['rem_doz']} โหล" if c['rem_doz'] else '')
            else:
                c = calc_foam212(qty)
                load = f"{c['packs']} กล่อง" + (f" เศษ {c['rem_doz']} โหล" if c['rem_doz'] else '')

            label = TYPE_LABEL.get(ptype,'อื่นๆ')
            if item['gift']: label += ' (ของแถม)'
            color = C_GIFT if item['gift'] else TYPE_COLOR.get(ptype,'FFFFFF')

            row_vals = [th_date(doc['date']),doc['docId'],doc['customer'],
                        item['desc'],label,qty,doz,load]
            for col,val in enumerate(row_vals,1):
                write(ws3,r3,col,val,
                    font=nf(10),fill=fl(color),
                    align=al('left' if col<=4 else 'center'),
                    border=bdr)
            ws3.row_dimensions[r3].height = 18
            r3 += 1

    for col,w in enumerate([12,14,24,36,16,8,8,20],1):
        ws3.column_dimensions[get_column_letter(col)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── UI ────────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "📂 ลาก PDF มาวาง หรือคลิกเพื่อเลือกไฟล์",
    type=['pdf'], accept_multiple_files=True,
    help="รองรับหลายไฟล์พร้อมกัน"
)

if uploaded_files:
    docs = []
    errors = []
    with st.spinner('🔍 กำลังอ่าน PDF...'):
        for f in uploaded_files:
            try:
                doc = parse_pdf(f.read())
                doc['_filename'] = f.name
                docs.append(doc)
            except Exception as e:
                errors.append(f"{f.name}: {e}")

    if errors:
        for e in errors: st.error(f"❌ {e}")

    if docs:
        docs.sort(key=lambda x: x['docId'] or '')
        for doc in docs:
            _aggregate(doc)

        # grand summary
        tot_all = sum(d['_ct']+d['_ft2']+d['_ft3']+d['_gct']+d['_gft2']+d['_gft3'] for d in docs)
        col1,col2,col3,col4 = st.columns(4)
        col1.metric("📄 เอกสาร", f"{len(docs)} IFO")

        # รวมผ้าใบแยกตามรุ่น
        from collections import defaultdict
        canvas_by_sub = defaultdict(int)
        for d in docs:
            for item in d['items']:
                if item['type'] == 'canvas':
                    canvas_by_sub[item['subtype']] += item['qty']
        canvas_summary = ' | '.join(f"{k} {v} คู่" for k,v in sorted(canvas_by_sub.items()))
        col2.metric("🟢 ผ้าใบ", canvas_summary or "0 คู่")
        col3.metric("🔵 ฟองน้ำ 200", f"{sum(d['_ft2']+d['_gft2'] for d in docs)} คู่")
        col4.metric("🔵 ฟองน้ำ 212/213", f"{sum(d['_ft3']+d['_gft3'] for d in docs)} คู่")

        st.divider()

        # per doc
        for doc in docs:
            with st.expander(f"📄 {doc['docId']} — {doc['customer']} ({th_date(doc['date'])})", expanded=True):
                c1,c2,c3 = st.columns(3)
                c1.write(f"**ลูกค้า:** {doc['customer']}")
                c2.write(f"**อ.{doc.get('amphoe','')} จ.{doc.get('province','')}**")
                c3.write(f"**รวม:** {doc['_ct']+doc['_ft2']+doc['_ft3']+doc['_gct']+doc['_gft2']+doc['_gft3']} คู่")

                if doc['_ct']:
                    c = calc_canvas(doc['_ct'])
                    st.success(f"🟢 **ผ้าใบ:** {fmt_canvas(c)}")
                if doc['_ft2']:
                    c = calc_foam200(doc['_ft2'])
                    st.info(f"🔵 **ฟองน้ำ 200:** {fmt_foam200(c)}")
                if doc['_ft3']:
                    c = calc_foam212(doc['_ft3'])
                    st.info(f"🔵 **ฟองน้ำ 212/213:** {fmt_foam212(c)}")
                if doc['_gct'] or doc['_gft2'] or doc['_gft3']:
                    gift_parts = []
                    if doc['_gct']:  gift_parts.append(f"ผ้าใบ {fmt_canvas(calc_canvas(doc['_gct']))}")
                    if doc['_gft2']: gift_parts.append(f"ฟองน้ำ200 {fmt_foam200(calc_foam200(doc['_gft2']))}")
                    if doc['_gft3']: gift_parts.append(f"ฟองน้ำ212/213 {fmt_foam212(calc_foam212(doc['_gft3']))}")
                    st.warning(f"🎁 **ของแถม:** {' | '.join(gift_parts)}")

        # download
        excel_buf = build_excel(docs)
        fname = f"IFO_โหลดสินค้า_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        st.download_button(
            label="📥 ดาวน์โหลด Excel",
            data=excel_buf, file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, type="primary"
        )


