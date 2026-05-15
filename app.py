import streamlit as st
import pdfplumber
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

st.set_page_config(
    page_title="คำนวณโหลดสินค้า IFO",
    page_icon="📦",
    layout="centered"
)

# ── CSS ───────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'IBM Plex Sans Thai', sans-serif; }

.main-header {
    background: linear-gradient(135deg, #1B4F8A, #163D6E);
    color: white; border-radius: 12px; padding: 20px 24px;
    margin-bottom: 20px; display: flex; align-items: center; gap: 14px;
}
.main-header h1 { font-size: 20px; font-weight: 700; margin: 0; }
.main-header p  { font-size: 13px; opacity: 0.7; margin: 4px 0 0; }

.doc-card {
    background: white; border: 1px solid #DDE3ED;
    border-radius: 12px; padding: 16px; margin-bottom: 14px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.07);
}
.doc-title { font-size: 16px; font-weight: 700; color: #1B4F8A; margin-bottom: 2px; }
.doc-sub   { font-size: 13px; color: #64748B; margin-bottom: 12px; }

.metric-row { display: flex; gap: 10px; flex-wrap: wrap; margin: 8px 0; }
.metric-box {
    flex: 1; min-width: 90px; background: #F8FAFC;
    border: 1px solid #DDE3ED; border-radius: 8px;
    padding: 10px 12px; text-align: center;
}
.metric-label { font-size: 11px; color: #94A3B8; font-weight: 600;
                text-transform: uppercase; letter-spacing: .05em; margin-bottom: 3px; }
.metric-value { font-size: 22px; font-weight: 700; color: #1E293B; font-family: monospace; line-height: 1; }
.metric-unit  { font-size: 12px; color: #64748B; }
.metric-sub   { font-size: 11px; color: #16A34A; margin-top: 3px; }

.type-label { font-size: 13px; font-weight: 700; margin: 10px 0 6px; }
.lbl-canvas { color: #166534; }
.lbl-foam   { color: #1E40AF; }
.lbl-gift   { color: #854D0E; }

.total-bar {
    background: #1B4F8A; color: white; border-radius: 8px;
    padding: 10px 16px; display: flex; justify-content: space-between;
    align-items: center; margin-top: 10px; font-weight: 600;
}

.grand-card {
    background: linear-gradient(135deg, #1B4F8A, #1E40AF);
    color: white; border-radius: 12px; padding: 18px;
    margin-bottom: 16px;
}
.grand-card h3 { font-size: 14px; opacity: 0.8; margin-bottom: 12px; }

.chip {
    display: inline-block; font-size: 11px; font-weight: 600;
    padding: 2px 9px; border-radius: 20px; margin-right: 4px;
}
.chip-c { background: #DCFCE7; color: #166534; }
.chip-f { background: #DBEAFE; color: #1E40AF; }
.chip-g { background: #FEF9C3; color: #854D0E; }

.stDataFrame { border-radius: 8px; overflow: hidden; }
</style>
""", unsafe_allow_html=True)

# ── HEADER ────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <div style="font-size:32px">📦</div>
  <div>
    <h1>คำนวณโหลดสินค้า IFO</h1>
    <p>อัปโหลด PDF → คำนวณโหล / ลัง / กระสอบ → ดาวน์โหลด Excel</p>
  </div>
</div>
""", unsafe_allow_html=True)

# ── PARSE ─────────────────────────────────────────────
MONTH_MAP = {
    'ม.ค.':'01','ก.พ.':'02','มี.ค.':'03','เม.ย.':'04',
    'พ.ค.':'05','มิ.ย.':'06','ก.ค.':'07','ส.ค.':'08',
    'ก.ย.':'09','ต.ค.':'10','พ.ย.':'11','ธ.ค.':'12'
}

def parse_pdf(file_bytes):
    result = {'docId':'', 'date':'', 'customer':'', 'items':[]}
    text = ''
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text += (page.extract_text() or '') + '\n'

    # docId
    m = re.search(r'IFO-\d+', text)
    if m: result['docId'] = m.group()

    # date
    pattern = r'(\d{1,2})\s+(' + '|'.join(re.escape(k) for k in MONTH_MAP) + r')\s+(\d{4})'
    m = re.search(pattern, text)
    if m:
        d, mo, y = m.group(1), m.group(2), int(m.group(3))
        if y > 2500: y -= 543
        result['date'] = f"{y}-{MONTH_MAP[mo]}-{d.zfill(2)}"

    # customer — ดึงชื่อลูกค้าจาก PDF โดยหาจาก keyword แล้วดึงเฉพาะชื่อ
    lines = [l.strip() for l in text.split('\n') if l.strip()]

    # วิธีที่ 1: หาบรรทัดที่มี keyword ชื่อคน/ร้าน แล้วดึงเฉพาะส่วนชื่อ
    cust_keywords = ['ร้าน','น.ส.','นาย','นาง','หจก','บจก','ห้าง','บริษัท']
    for line in lines[:50]:
        for kw in cust_keywords:
            if kw in line:
                # ตัดข้อมูลที่ไม่เกี่ยวออก เช่น วันที่ เลขที่
                clean = re.sub(r'ชื่อ.*?:|รหัส.*?:|วันที่.*?:|เลขที่.*?:', '', line)
                clean = re.sub(r'\d{2}/\d{2}/\d{4}|\d{4}-\d{2}-\d{2}', '', clean)
                clean = clean.strip()
                if len(clean) > 2:
                    result['customer'] = clean
                    break
        if result['customer']:
            break

    # วิธีที่ 2: fallback หาชื่อสั้นๆ ภาษาไทยล้วน ไม่มีตัวเลข
    if not result['customer']:
        for line in lines[:30]:
            if (re.match(r'^[ก-๙\s\.\(\)/]{3,25}$', line)
                    and 'IFO' not in line
                    and 'บริษัท' not in line
                    and 'นันยาง' not in line
                    and not re.search(r'\d', line)):
                result['customer'] = line
                break

    # items
    for line in lines:
        if 'Z0001' in line or 'มัดจำ' in line: continue
        m = re.search(r'(\d{9})\s+(.+?)\s+(\d+)\s+คู่', line)
        if m:
            barcode, desc, qty = m.group(1), m.group(2).strip(), int(m.group(3))
            if qty > 0:
                t = detect_type(barcode, desc)
                desc_clean = re.sub(r'\s+\d+(\.\d+)?(\s+\d+(\.\d+)?)*$', '', desc).strip()
                result['items'].append({'desc': desc_clean, 'type': t, 'qty': qty})
    return result, text

def detect_type(barcode, desc):
    txt = barcode + desc
    if re.search(r'ผ้าใบ|205[SR]', txt): return 'canvas'
    if re.search(r'ฟองน้ำ|200|212|213', txt): return 'foam'
    return 'gift'

# ── CALC ──────────────────────────────────────────────
def foam_calc(n):
    if n == 0: return {'doz':0,'rp':0,'txt':'-','big':0}
    doz, rp = divmod(n, 12)
    big, rem = divmod(n, 120)
    sm, lf = divmod(rem, 12)
    txt = f'{big} กระสอบใหญ่'
    if sm: txt += f' {sm} กระสอบ'
    if lf: txt += f' 1 กระสอบ({lf} คู่)'
    return {'doz':doz,'rp':rp,'big':big,'sm':sm,'lf':lf,'txt':txt}

def canvas_calc(n):
    lang, rem = divmod(n, 12)
    return {'lang':lang,'rem':rem,'doz':lang}

def format_th_date(s):
    if not s: return ''
    try:
        y, m, d = s.split('-')
        mn = ['','ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.']
        return f"{int(d)} {mn[int(m)]} {int(y)+543}"
    except: return s

# ── EXCEL ─────────────────────────────────────────────
def build_excel(docs):
    wb = Workbook()
    BLUE='1B4F8A'; WHITE='FFFFFF'
    CANVAS_BG='E2EFDA'; FOAM_BG='DBEAFE'; GIFT_BG='FEF9C3'; SUM_BG='D6E4F0'
    thin = Side(style='thin', color='CCCCCC')
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)

    def hf(sz=11): return Font(name='TH Sarabun New', bold=True, color=WHITE, size=sz)
    def nf(bold=False, sz=10): return Font(name='TH Sarabun New', bold=bold, size=sz)
    def fl(c): return PatternFill('solid', start_color=c, end_color=c)
    def al(h='left'): return Alignment(horizontal=h, vertical='center', wrap_text=True)

    # ── Sheet 1: สรุป ──
    ws = wb.active
    ws.title = 'สรุปโหลดสินค้า'
    # col layout: A=วันที่ B=เลขที่ C=ลูกค้า D=ผ้าใบ(คู่) E=ผ้าใบ(ลัง) F=ฟองน้ำ(คู่) G=ฟองน้ำ(โหล) H=ฟองน้ำ(กระสอบ) I=ของแถม(คู่) J=รวม
    NCOLS = 10
    ws.merge_cells('A1:J1')
    ws['A1'] = 'สรุปโหลดสินค้า — บริษัท นันยางมาร์เก็ตติ้ง จำกัด'
    ws['A1'].font = hf(14); ws['A1'].fill = fl(BLUE)
    ws['A1'].alignment = al('center'); ws.row_dimensions[1].height = 28

    ws.merge_cells('A2:J2')
    ws['A2'] = f'พิมพ์: {datetime.now().strftime("%d/%m/%Y %H:%M")}'
    ws['A2'].font = nf(); ws['A2'].alignment = al('right')
    ws.row_dimensions[2].height = 16

    # Row 3: group headers
    ws.merge_cells('A3:C3')
    ws.merge_cells('D3:E3')
    ws.merge_cells('F3:H3')
    ws.merge_cells('I3:I3')
    ws.merge_cells('J3:J3')
    grp_hdrs = {1:'', 4:'ผ้าใบ', 6:'ฟองน้ำ', 9:'ของแถม', 10:'รวม (คู่)'}
    grp_fills = {1:BLUE, 4:'2D6A27', 6:'1E3A8A', 9:'78350F', 10:BLUE}
    for col in range(1, NCOLS+1):
        c = ws.cell(row=3, column=col)
        for k in sorted(grp_hdrs.keys(), reverse=True):
            if col >= k:
                c.value = grp_hdrs[k] if col == k else ''
                c.fill = fl(grp_fills[k])
                break
        c.font = hf(11); c.alignment = al('center'); c.border = bdr
    ws.row_dimensions[3].height = 20

    # Row 4: sub headers
    sub_hdrs = ['วันที่','เลขที่เอกสาร','ชื่อลูกค้า','คู่','ลัง','คู่','โหล','กระสอบ','คู่','รวม (คู่)']
    sub_fills = [BLUE,BLUE,BLUE,'2D6A27','2D6A27','1E3A8A','1E3A8A','1E3A8A','78350F',BLUE]
    for col, (h, bg) in enumerate(zip(sub_hdrs, sub_fills), 1):
        c = ws.cell(row=4, column=col, value=h)
        c.font = hf(10); c.fill = fl(bg)
        c.alignment = al('center'); c.border = bdr
    ws.row_dimensions[4].height = 20

    tot_c = tot_f = tot_g = 0
    for i, doc in enumerate(docs):
        r = 5 + i
        ci = [x for x in doc['items'] if x['type']=='canvas']
        fi = [x for x in doc['items'] if x['type']=='foam']
        gi = [x for x in doc['items'] if x['type']=='gift']
        ct = sum(x['qty'] for x in ci)
        ft = sum(x['qty'] for x in fi)
        gt = sum(x['qty'] for x in gi)
        cf_ = canvas_calc(ct); ff_ = foam_calc(ft); gf_ = canvas_calc(gt)
        tot_c += ct; tot_f += ft; tot_g += gt

        rem_txt = f'เศษ {cf_["rem"]} คู่' if cf_['rem'] else ''
        sack_txt = ff_['txt'] if ft else '-'

        bg = 'F7FAFB' if i % 2 == 0 else 'FFFFFF'
        row_vals = [
            format_th_date(doc['date']), doc['docId'], doc['customer'],
            ct if ct else '-', f'{cf_["lang"]} ลัง' + (f' ({rem_txt})' if rem_txt else '') if ct else '-',
            ft if ft else '-', f'{ff_["doz"]} โหล' if ft else '-', sack_txt,
            gt if gt else '-',
            ct+ft+gt
        ]
        row_bgs = [bg,bg,bg, CANVAS_BG,CANVAS_BG, FOAM_BG,FOAM_BG,FOAM_BG, GIFT_BG if gt else bg, SUM_BG]
        for col, (val, bg2) in enumerate(zip(row_vals, row_bgs), 1):
            c = ws.cell(row=r, column=col, value=val)
            c.fill = fl(bg2); c.font = nf(bold=(col==10))
            c.alignment = al('left') if col <= 3 else al('center')
            c.border = bdr
        ws.row_dimensions[r].height = 28

    # summary row
    sr = 5 + len(docs)
    ff_t = foam_calc(tot_f); cf_t = canvas_calc(tot_c)
    ws.merge_cells(f'A{sr}:C{sr}')
    sum_data = {
        1: 'รวมทั้งหมด',
        4: tot_c if tot_c else '-',
        5: f'{cf_t["lang"]} ลัง',
        6: tot_f if tot_f else '-',
        7: f'{ff_t["doz"]} โหล',
        8: ff_t["txt"] if tot_f else '-',
        9: tot_g if tot_g else '-',
        10: tot_c+tot_f+tot_g
    }
    for col in range(1, NCOLS+1):
        c = ws.cell(row=sr, column=col)
        if col in sum_data:
            c.value = sum_data[col]
        c.font = hf(10); c.fill = fl(BLUE)
        c.alignment = al('center'); c.border = bdr
    ws.row_dimensions[sr].height = 28

    for col, w in enumerate([12,16,26,10,14,10,12,26,10,12], 1):
        ws.column_dimensions[get_column_letter(col)].width = w

    # ── Sheet 2: รายละเอียด ──
    ws2 = wb.create_sheet('รายละเอียด')
    ws2.merge_cells('A1:H1')
    ws2['A1'] = 'รายละเอียดสินค้าแยกรายเอกสาร'
    ws2['A1'].font = hf(13); ws2['A1'].fill = fl(BLUE)
    ws2['A1'].alignment = al('center'); ws2.row_dimensions[1].height = 24

    dh = ['วันที่','เลขที่เอกสาร','ชื่อลูกค้า','รายการ','ประเภท','คู่','โหล','ลัง/กระสอบ']
    for col, h in enumerate(dh, 1):
        c = ws2.cell(row=2, column=col, value=h)
        c.font = hf(10); c.fill = fl(BLUE)
        c.alignment = al('center'); c.border = bdr
    ws2.row_dimensions[2].height = 20

    TYPE_LABEL = {'canvas':'ผ้าใบ','foam':'ฟองน้ำ','gift':'ของแถม'}
    TYPE_COLOR = {'canvas':CANVAS_BG,'foam':FOAM_BG,'gift':GIFT_BG}
    r2 = 3
    for doc in docs:
        for item in doc['items']:
            doz, rem = divmod(item['qty'], 12)
            if item['type'] == 'foam':
                load_txt = foam_calc(item['qty'])['txt']
            else:
                cc = canvas_calc(item['qty'])
                load_txt = f'{cc["lang"]} ลัง' + (f' เศษ {cc["rem"]} คู่' if cc['rem'] else '')
            vals = [format_th_date(doc['date']), doc['docId'], doc['customer'],
                    item['desc'], TYPE_LABEL.get(item['type'],''),
                    item['qty'], f'{doz} โหล' + (f' เศษ {rem}' if rem else ''), load_txt]
            bg = TYPE_COLOR.get(item['type'],'FFFFFF')
            for col, val in enumerate(vals, 1):
                c = ws2.cell(row=r2, column=col, value=val)
                c.fill = fl(bg); c.font = nf()
                c.alignment = al('left') if col <= 4 else al('center')
                c.border = bdr
            r2 += 1

    for col, w in enumerate([14,16,28,36,10,10,16,28], 1):
        ws2.column_dimensions[get_column_letter(col)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── UI ────────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "📂 ลาก PDF มาวาง หรือคลิกเพื่อเลือกไฟล์",
    type=['pdf'],
    accept_multiple_files=True,
    help="รองรับหลายไฟล์พร้อมกัน"
)

if uploaded_files:
    docs = []
    errors = []

    with st.spinner('🔍 กำลังอ่าน PDF...'):
        for f in uploaded_files:
            try:
                doc, _ = parse_pdf(f.read())
                doc['_filename'] = f.name
                docs.append(doc)
            except Exception as e:
                errors.append(f"{f.name}: {str(e)}")

    if errors:
        for e in errors:
            st.error(f"❌ {e}")

    if docs:
        # sort by date
        docs.sort(key=lambda x: x['date'] or '')

        # totals
        tot_c = tot_f = tot_g = 0
        for doc in docs:
            tot_c += sum(x['qty'] for x in doc['items'] if x['type']=='canvas')
            tot_f += sum(x['qty'] for x in doc['items'] if x['type']=='foam')
            tot_g += sum(x['qty'] for x in doc['items'] if x['type']=='gift')

        # grand summary
        if len(docs) > 1:
            ff_ = foam_calc(tot_f); cf_ = canvas_calc(tot_c)
            st.markdown(f"""
            <div class="grand-card">
              <h3>📊 สรุปรวม {len(docs)} เอกสาร</h3>
              <div style="display:flex;gap:10px;flex-wrap:wrap">
                <div style="flex:1;min-width:100px;background:rgba(255,255,255,.15);border-radius:8px;padding:10px;text-align:center">
                  <div style="font-size:11px;opacity:.7;margin-bottom:4px">ผ้าใบ</div>
                  <div style="font-size:22px;font-weight:700;font-family:monospace">{tot_c}</div>
                  <div style="font-size:11px;opacity:.75">{cf_['lang']} ลัง</div>
                </div>
                <div style="flex:1;min-width:100px;background:rgba(255,255,255,.15);border-radius:8px;padding:10px;text-align:center">
                  <div style="font-size:11px;opacity:.7;margin-bottom:4px">ฟองน้ำ</div>
                  <div style="font-size:22px;font-weight:700;font-family:monospace">{tot_f}</div>
                  <div style="font-size:11px;opacity:.75">{ff_['doz']} โหล<br>{ff_['txt']}</div>
                </div>
                <div style="flex:1;min-width:100px;background:rgba(255,255,255,.15);border-radius:8px;padding:10px;text-align:center">
                  <div style="font-size:11px;opacity:.7;margin-bottom:4px">รวมทั้งสิ้น</div>
                  <div style="font-size:22px;font-weight:700;font-family:monospace">{tot_c+tot_f+tot_g}</div>
                  <div style="font-size:11px;opacity:.75">คู่</div>
                </div>
              </div>
            </div>
            """, unsafe_allow_html=True)

        # per doc
        for doc in docs:
            ci = [x for x in doc['items'] if x['type']=='canvas']
            fi = [x for x in doc['items'] if x['type']=='foam']
            gi = [x for x in doc['items'] if x['type']=='gift']
            ct = sum(x['qty'] for x in ci)
            ft = sum(x['qty'] for x in fi)
            gt = sum(x['qty'] for x in gi)
            cf_ = canvas_calc(ct); ff_ = foam_calc(ft); gf_ = canvas_calc(gt)

            chips = ''
            if ct: chips += f'<span class="chip chip-c">ผ้าใบ {ct} คู่</span>'
            if ft: chips += f'<span class="chip chip-f">ฟองน้ำ {ft} คู่</span>'
            if gt: chips += f'<span class="chip chip-g">ของแถม {gt} คู่</span>'

            st.markdown(f"""
            <div class="doc-card">
              <div class="doc-title">{doc['docId'] or doc['_filename']}</div>
              <div class="doc-sub">{format_th_date(doc['date'])} · {doc['customer'] or '—'}</div>
              <div>{chips}</div>
            """, unsafe_allow_html=True)

            if ct > 0:
                st.markdown(f'<div class="type-label lbl-canvas">🟢 ผ้าใบ</div>', unsafe_allow_html=True)
                col1, col2, col3 = st.columns(3)
                col1.metric("รวม", f"{ct} คู่")
                col2.metric("โหล", f"{cf_['doz']} โหล", delta=f"เศษ {cf_['rem']} คู่" if cf_['rem'] else None)
                col3.metric("ลัง", f"{cf_['lang']} ลัง", delta=f"เศษ {cf_['rem']} คู่" if cf_['rem'] else None)

            if ft > 0:
                st.markdown(f'<div class="type-label lbl-foam">🔵 ฟองน้ำ</div>', unsafe_allow_html=True)
                col1, col2, col3 = st.columns(3)
                col1.metric("รวม", f"{ft} คู่")
                col2.metric("โหล", f"{ff_['doz']} โหล", delta=f"เศษ {ff_['rp']} คู่" if ff_['rp'] else None)
                col3.metric("กระสอบ", ff_['txt'])

            if gt > 0:
                st.markdown(f'<div class="type-label lbl-gift">🎁 ของแถม</div>', unsafe_allow_html=True)
                col1, col2, col3 = st.columns(3)
                col1.metric("รวม", f"{gt} คู่")
                col2.metric("โหล", f"{gf_['doz']} โหล")
                col3.metric("ลัง", f"{gf_['lang']} ลัง")

            if not doc['items']:
                st.warning("⚠️ ไม่พบรายการสินค้าในเอกสารนี้")

            st.markdown("</div>", unsafe_allow_html=True)
            st.markdown("---")

        # download excel
        excel_buf = build_excel(docs)
        fname = f"IFO_โหลดสินค้า_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        st.download_button(
            label="📥 ดาวน์โหลด Excel",
            data=excel_buf,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )

else:
    st.markdown("""
    <div style="text-align:center;padding:48px 20px;color:#94A3B8">
      <div style="font-size:48px;margin-bottom:12px">📄</div>
      <div style="font-size:15px;font-weight:600;color:#64748B">ลาก PDF มาวางที่นี่</div>
      <div style="font-size:13px;margin-top:6px">รองรับหลายไฟล์พร้อมกัน</div>
    </div>
    """, unsafe_allow_html=True)
