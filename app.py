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

def norm(s):
    """แก้ตัวอักษรไทยที่ PDF ทำให้เพี้ยน"""
    fixes = [
        (r'รา้น','ร้าน'),(r'รา้ น','ร้าน'),(r'ราน้','ร้าน'),
        (r'บรษิ ทั','บริษัท'),(r'บรษิัท','บริษัท'),
        (r'จาํ กดั','จำกัด'),(r'จํากัด','จำกัด'),
        (r'ภเูก็ต','ภูเก็ต'),(r'ภเูก็ต','ภูเก็ต'),
        (r'เมอื ง','เมือง'),(r'เมอื ง','เมือง'),
        (r'ประจวบครีขี ันธ์','ประจวบคีรีขันธ์'),
        (r'ประจวบคีรีขันธ์','ประจวบคีรีขันธ์'),
        (r'นครสวรรค์','นครสวรรค์'),
        (r'กําแพงเพชร','กำแพงเพชร'),
        (r'สพุ รรณบรุ ี','สุพรรณบุรี'),
        (r'หนองคาย','หนองคาย'),
        (r'เชยงราย ี','เชียงราย'),(r'เชยี งราย','เชียงราย'),
        (r'ระยอง','ระยอง'),(r'สระบรุ ี','สระบุรี'),
        (r'ขาณวุรลักษบรุ ี','ขาณุวรลักษบุรี'),
        (r'คณุ ','คุณ '),(r'รา้น ','ร้าน'),
        (r'ซปุ เปอร์','ซุปเปอร์'),(r'ชป','ชอป'),
        (r'\s+',' '),
    ]
    for p,r in fixes: s = re.sub(p,r,s)
    return s.strip()

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
        # หา items โดย match barcode + desc + qty ทีละรายการ
        # format: <barcode9หลัก> <desc> <qty> คู่ <ราคา>
        item_pattern = re.compile(r'(?<![\d])((1[12]\d{7}))\s+(.+?)\s+(\d{1,4})\s+คู่')
        for m in item_pattern.finditer(full_text):
            bc, desc_raw, qty = m.group(1), m.group(3).strip(), int(m.group(4))
            if qty <= 0 or qty > 9999: continue
            if 'Z0001' in bc: continue
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

    # ── Colors & Styles ──
    BLUE='1B4F8A'; WHITE='FFFFFF'
    CG='E2EFDA'; CGH='375623'      # ผ้าใบ
    CF='DBEAFE'; CFH='1E3A8A'      # ฟองน้ำ 200
    C212='D0E4FF'; C212H='1E40AF'  # ฟองน้ำ 212
    C213='EDE9FE'; C213H='5B21B6'  # ฟองน้ำ 213
    CGIFT='FFF9C4'; CALT='F5F8FA'

    thin=Side(style='thin',color='BBBBBB')
    bdr=Border(left=thin,right=thin,top=thin,bottom=thin)

    FONT = 'TH Sarabun New'
    def hf(sz=11,color=WHITE,bold=True):
        return Font(name=FONT,bold=bold,color=color,size=sz)
    def nf(sz=11,bold=False,color='1E293B'):
        return Font(name=FONT,size=sz,bold=bold,color=color)
    def fl(c): return PatternFill('solid',start_color=c,end_color=c)
    def al(h='center',wrap=True):
        return Alignment(horizontal=h,vertical='center',wrap_text=wrap)

    def wcell(ws,r,c,val,font=None,fill=None,align=None,border=None):
        cell = ws.cell(row=r,column=c,value=val)
        if font:  cell.font=font
        if fill:  cell.fill=fill
        if align: cell.alignment=align
        if border:cell.border=border
        return cell

    def mwrite(ws,r1,c1,r2,c2,val,font=None,fill=None,align=None):
        if r1==r2 and c1==c2:
            ws.cell(row=r1,column=c1,value=val)
        else:
            ws.merge_cells(start_row=r1,start_column=c1,end_row=r2,end_column=c2)
            ws.cell(row=r1,column=c1,value=val)
        for rr in range(r1,r2+1):
            for cc in range(c1,c2+1):
                cell=ws.cell(row=rr,column=cc)
                if font:  cell.font=font
                if fill:  cell.fill=fill
                if align: cell.alignment=align
                cell.border=bdr

    # ── หา subtypes ──
    canvas_subs = sorted(set(x['subtype'] for d in docs for x in d['items'] if x['type']=='canvas'))
    has212 = any(x['subtype']=='212' for d in docs for x in d['items'] if x['type']=='foam212')
    has213 = any(x['subtype']=='213' for d in docs for x in d['items'] if x['type']=='foam212')

    # ══════════════════════════════════════════
    # SHEET 1: สรุปรายเอกสาร
    # ══════════════════════════════════════════
    ws1 = wb.active; ws1.title='สรุปรายเอกสาร'

    # build dynamic columns
    grp_cols = []  # (start, end, label, header_color, data_color)
    # info cols A-E
    grp_cols.append({'sc':1,'ec':5,'label':'ข้อมูลเอกสาร','hc':BLUE,'dc':BLUE,
                     'subs':['วันที่','เลขที่IFO','ชื่อลูกค้า','อำเภอ','จังหวัด']})
    col=6
    for sub in canvas_subs:
        grp_cols.append({'sc':col,'ec':col+2,'label':f'ผ้าใบ {sub} (12 คู่/กล่อง)','hc':CGH,'dc':CG,
                         'subs':['คู่','กล่อง','เศษคู่']})
        col+=3
    grp_cols.append({'sc':col,'ec':col+3,'label':'ฟองน้ำ 200 (120 คู่/กระสอบ)','hc':CFH,'dc':CF,
                     'subs':['คู่','โหล','กระสอบ','เศษโหล']})
    col+=4
    if has212:
        grp_cols.append({'sc':col,'ec':col+2,'label':'ฟองน้ำ 212 (24 คู่/กล่อง)','hc':C212H,'dc':C212,
                         'subs':['คู่','กล่อง','เศษโหล']})
        col+=3
    if has213:
        grp_cols.append({'sc':col,'ec':col+2,'label':'ฟองน้ำ 213 (24 คู่/กล่อง)','hc':C213H,'dc':C213,
                         'subs':['คู่','กล่อง','เศษโหล']})
        col+=3
    grp_cols.append({'sc':col,'ec':col,'label':'รวม','hc':BLUE,'dc':BLUE,'subs':['รวมคู่']})
    NCOLS=col

    # Row 1: title
    mwrite(ws1,1,1,1,NCOLS,'สรุปโหลดสินค้า — บริษัท นันยางมาร์เก็ตติ้ง จำกัด',
           font=hf(14),fill=fl(BLUE),align=al())
    ws1.row_dimensions[1].height=28
    # Row 2: date
    mwrite(ws1,2,1,2,NCOLS,f'พิมพ์: {datetime.now().strftime("%d/%m/%Y %H:%M")}',
           font=nf(10),align=al('right'))
    ws1.row_dimensions[2].height=16
    # Row 3: group headers
    for g in grp_cols:
        mwrite(ws1,3,g['sc'],3,g['ec'],g['label'],font=hf(11),fill=fl(g['hc']),align=al())
    ws1.row_dimensions[3].height=22
    # Row 4: sub headers
    col_idx=1
    col_dc_map={}
    for g in grp_cols:
        for si,sub in enumerate(g['subs']):
            wcell(ws1,4,col_idx,sub,font=hf(10),fill=fl(g['hc']),align=al(),border=bdr)
            col_dc_map[col_idx]=g['dc']
            col_idx+=1
    ws1.row_dimensions[4].height=34

    # Data rows
    tot=defaultdict(int)
    for i,doc in enumerate(docs):
        r=5+i
        bg='FFFFFF' if i%2==0 else CALT
        canvas_qty=defaultdict(int)
        for x in doc['items']:
            if x['type']=='canvas': canvas_qty[x['subtype']]+=x['qty']
        ft2=doc['_ft2']; f212=doc['_ft3_212']; f213=doc['_ft3_213']
        grand=doc['_ct']+ft2+f212+f213
        for sub in canvas_subs: tot[f'c_{sub}']+=canvas_qty.get(sub,0)
        tot['ft2']+=ft2; tot['f212']+=f212; tot['f213']+=f213

        row_vals=[th_date(doc['date']),doc['docId'],doc['customer'],
                  doc.get('amphoe',''),doc.get('province','')]
        row_bgs=[bg]*5
        for sub in canvas_subs:
            ct=canvas_qty.get(sub,0); cc=calc_canvas(ct)
            row_vals+=[ct or '-',cc['boxes'] or '-',cc['rem'] or '-']
            row_bgs+=[CG if ct else bg]*3
        cf2=calc_foam200(ft2)
        row_vals+=[ft2 or '-',cf2['doz'] or '-',cf2['sacks'] or '-',cf2['rem_doz'] or '-']
        row_bgs+=[CF if ft2 else bg]*4
        if has212:
            c2=calc_foam212(f212)
            row_vals+=[f212 or '-',c2['boxes'] or '-',c2['rem_doz'] or '-']
            row_bgs+=[C212 if f212 else bg]*3
        if has213:
            c3=calc_foam212(f213)
            row_vals+=[f213 or '-',c3['boxes'] or '-',c3['rem_doz'] or '-']
            row_bgs+=[C213 if f213 else bg]*3
        row_vals+=[grand]; row_bgs+=['D6E4F0']

        for ci,(val,bgc) in enumerate(zip(row_vals,row_bgs),1):
            wcell(ws1,r,ci,val,
                  font=nf(11,bold=(ci==NCOLS)),
                  fill=fl(bgc),
                  align=al('left' if ci<=3 else 'center'),
                  border=bdr)
        ws1.row_dimensions[r].height=22

    # Summary row
    sr=5+len(docs)
    mwrite(ws1,sr,1,sr,5,'รวมทั้งหมด',font=hf(11),fill=fl(BLUE),align=al())
    ci=6
    for sub in canvas_subs:
        cc=calc_canvas(tot[f'c_{sub}'])
        for v in [tot[f'c_{sub}'] or '-',cc['boxes'] or '-',cc['rem'] or '-']:
            wcell(ws1,sr,ci,v,font=hf(11),fill=fl(BLUE),align=al('center'),border=bdr); ci+=1
    cf2=calc_foam200(tot['ft2'])
    for v in [tot['ft2'] or '-',cf2['doz'] or '-',cf2['sacks'] or '-',cf2['rem_doz'] or '-']:
        wcell(ws1,sr,ci,v,font=hf(11),fill=fl(BLUE),align=al('center'),border=bdr); ci+=1
    if has212:
        c2=calc_foam212(tot['f212'])
        for v in [tot['f212'] or '-',c2['boxes'] or '-',c2['rem_doz'] or '-']:
            wcell(ws1,sr,ci,v,font=hf(11),fill=fl(BLUE),align=al('center'),border=bdr); ci+=1
    if has213:
        c3=calc_foam212(tot['f213'])
        for v in [tot['f213'] or '-',c3['boxes'] or '-',c3['rem_doz'] or '-']:
            wcell(ws1,sr,ci,v,font=hf(11),fill=fl(BLUE),align=al('center'),border=bdr); ci+=1
    wcell(ws1,sr,ci,sum(tot.values()),font=hf(11),fill=fl(BLUE),align=al('center'),border=bdr)
    ws1.row_dimensions[sr].height=24

    # col widths sheet1
    base_w=[12,14,26,12,14]
    for _ in canvas_subs: base_w+=[8,8,8]
    base_w+=[8,8,10,10]
    if has212: base_w+=[8,8,8]
    if has213: base_w+=[8,8,8]
    base_w+=[10]
    for ci,w in enumerate(base_w,1):
        ws1.column_dimensions[get_column_letter(ci)].width=w

    # ══════════════════════════════════════════
    # SHEET 2: สรุปรวม
    # ══════════════════════════════════════════
    ws2=wb.create_sheet('สรุปรวม')

    mwrite(ws2,1,1,1,8,f'สรุปรวมทุกเอกสาร ({len(docs)} IFO) — {datetime.now().strftime("%d/%m/%Y %H:%M")}',
           font=hf(13),fill=fl(BLUE),align=al())
    ws2.row_dimensions[1].height=26

    # summary boxes row 3-4
    blocks=[]
    for sub in canvas_subs:
        ct=tot[f'c_{sub}']; cc=calc_canvas(ct)
        blocks.append((f'ผ้าใบ {sub} (ปกติ)',f'{ct} คู่ / {cc["boxes"]} กล่อง',CGH,CG))
    if tot['ft2']:
        cf2=calc_foam200(tot['ft2'])
        txt=f'{tot["ft2"]} คู่ / {cf2["sacks"]} กระสอบ'
        if cf2['rem_doz']: txt+=f' เศษ {cf2["rem_doz"]} โหล'
        blocks.append(('ฟองน้ำ 200 (ปกติ)',txt,CFH,CF))
    if tot['f212']:
        c2=calc_foam212(tot['f212'])
        blocks.append(('ฟองน้ำ 212 (ปกติ)',f'{tot["f212"]} คู่ / {c2["boxes"]} กล่อง',C212H,C212))
    if tot['f213']:
        c3=calc_foam212(tot['f213'])
        blocks.append(('ฟองน้ำ 213 (ปกติ)',f'{tot["f213"]} คู่ / {c3["boxes"]} กล่อง',C213H,C213))

    col=1
    for label,val,hc,dc in blocks:
        mwrite(ws2,3,col,3,col+1,label,font=hf(11),fill=fl(hc),align=al())
        mwrite(ws2,4,col,4,col+1,val,font=nf(12,bold=True),fill=fl(dc),align=al())
        ws2.column_dimensions[get_column_letter(col)].width=16
        ws2.column_dimensions[get_column_letter(col+1)].width=16
        col+=2
    ws2.row_dimensions[3].height=22; ws2.row_dimensions[4].height=26

    # breakdown table
    r=6
    mwrite(ws2,r,1,r,8,'รายละเอียดตามชนิดสินค้า (รวมทุก IFO)',font=hf(12),fill=fl(BLUE),align=al())
    ws2.row_dimensions[r].height=22; r+=1

    det_hdrs=['ชนิดสินค้า','ประเภท','รวม คู่','โหล','กล่อง/กระสอบ','เศษโหล','เศษคู่','หน่วย']
    for ci,h in enumerate(det_hdrs,1):
        wcell(ws2,r,ci,h,font=hf(11),fill=fl(BLUE),align=al(),border=bdr)
    ws2.row_dimensions[r].height=22; r+=1

    TYPE_UNIT={'canvas':'กล่อง','foam200':'กระสอบ','foam212':'กล่อง'}
    subtype_data=defaultdict(lambda:defaultdict(int))
    for doc in docs:
        for x in doc['items']:
            subtype_data[(x['subtype'],x['type'],x['gift'])]['qty']+=x['qty']

    sorted_keys=sorted(subtype_data.keys(),key=lambda x:({'canvas':0,'foam200':1,'foam212':2}.get(x[1],3),int(x[2]),x[0]))
    last_gift=None
    for key in sorted_keys:
        sub,ptype,gift=key
        qty=subtype_data[key]['qty']
        if qty==0: continue
        if gift and last_gift==False:
            ws2.row_dimensions[r].height=6; r+=1
        last_gift=gift
        doz=qty//12
        if ptype=='canvas':
            cc=calc_canvas(qty); pack=cc['boxes']; rem_doz='-'; rem_pair=cc['rem'] or '-'; unit='กล่อง'; dc=CGIFT if gift else CG
        elif ptype=='foam200':
            cf=calc_foam200(qty); pack=cf['sacks']; rem_doz=cf['rem_doz'] or '-'; rem_pair=cf['rem_pairs'] if 'rem_pairs' in cf else '-'; unit='กระสอบ'; dc=CGIFT if gift else CF
        else:
            cf=calc_foam212(qty); pack=cf['boxes']; rem_doz=cf['rem_doz'] or '-'; rem_pair=cf['rem'] if 'rem' in cf else '-'; unit='กล่อง'
            dc=(CGIFT if gift else (C212 if sub=='212' else C213))
        row_v=[sub,'ของแถม' if gift else 'ปกติ',qty,doz,pack,rem_doz,rem_pair,unit]
        for ci,v in enumerate(row_v,1):
            wcell(ws2,r,ci,v,font=nf(11,bold=(ci==1)),fill=fl(dc),align=al('left' if ci<=2 else 'center'),border=bdr)
        ws2.row_dimensions[r].height=20; r+=1

    for ci,w in enumerate([16,10,10,8,14,10,10,10],1):
        ws2.column_dimensions[get_column_letter(ci)].width=w

    # ══════════════════════════════════════════
    # SHEET 3: รายละเอียด
    # ══════════════════════════════════════════
    ws3=wb.create_sheet('รายละเอียด')
    mwrite(ws3,1,1,1,8,'รายละเอียดสินค้าทุกรายการ',font=hf(13),fill=fl(BLUE),align=al())
    ws3.row_dimensions[1].height=24

    dh=['วันที่','เลขที่IFO','ชื่อลูกค้า','รายการสินค้า','ประเภท','คู่','โหล','กล่อง/กระสอบ/กล่อง']
    for ci,h in enumerate(dh,1):
        wcell(ws3,2,ci,h,font=hf(11),fill=fl(BLUE),align=al(),border=bdr)
    ws3.row_dimensions[2].height=22

    TYPE_LABEL={'canvas':'ผ้าใบ','foam200':'ฟองน้ำ 200','foam212':'ฟองน้ำ 212/213'}
    TYPE_COLOR={'canvas':CG,'foam200':CF,'foam212':C212}

    r3=3
    for doc in docs:
        for x in doc['items']:
            qty=x['qty']; doz=qty//12; ptype=x['type']
            if ptype=='canvas':
                cc=calc_canvas(qty); load=f'{cc["boxes"]} กล่อง'+(f' เศษ {cc["rem"]} คู่' if cc['rem'] else '')
            elif ptype=='foam200':
                cf=calc_foam200(qty); load=f'{cf["sacks"]} กระสอบ'+(f' เศษ {cf["rem_doz"]} โหล' if cf['rem_doz'] else '')
            else:
                cf=calc_foam212(qty); load=f'{cf["boxes"]} กล่อง'+(f' เศษ {cf["rem_doz"]} โหล' if cf['rem_doz'] else '')
            label=TYPE_LABEL.get(ptype,'อื่นๆ')
            dc=CGIFT if x['gift'] else (C213 if x['subtype']=='213' else TYPE_COLOR.get(ptype,'FFFFFF'))
            row_v=[th_date(doc['date']),doc['docId'],doc['customer'],x['desc'],label,qty,doz,load]
            for ci,v in enumerate(row_v,1):
                wcell(ws3,r3,ci,v,font=nf(10),fill=fl(dc),align=al('left' if ci<=4 else 'center'),border=bdr)
            ws3.row_dimensions[r3].height=18; r3+=1

    for ci,w in enumerate([12,14,26,36,14,8,8,20],1):
        ws3.column_dimensions[get_column_letter(ci)].width=w

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
