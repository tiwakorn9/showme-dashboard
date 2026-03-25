# ============================================================
#  ShowMe Thailand — Business Dashboard
#  ไฟล์หลัก: app.py
#
#  วิธีใช้งาน:
#  1. เปิด Terminal ใน VSCode
#  2. พิมพ์: pip install flask pandas openpyxl
#  3. พิมพ์: python app.py
#  4. เปิด browser ไปที่: http://localhost:5050
# ============================================================

# === นำเข้า Library ===
from flask import Flask, render_template, request, jsonify
# Flask    = framework สำหรับสร้าง Web App ด้วย Python
# render_template = แสดงไฟล์ HTML
# request  = รับข้อมูลที่ส่งมาจาก browser
# jsonify  = แปลงข้อมูล Python เป็น JSON ส่งกลับ browser

import pandas as pd
# pandas = library จัดการข้อมูลตาราง เหมือน Excel แต่เร็วกว่ามาก

import warnings
warnings.filterwarnings('ignore')  # ซ่อน warning ที่ไม่จำเป็น

# สร้าง Flask app
# __name__ = ชื่อไฟล์ปัจจุบัน (app.py)
app = Flask(__name__)

# ============================================================
# ส่วนที่ 1: แปลงวันที่ภาษาไทย
# ============================================================
# BigSeller export วันที่เป็นภาษาไทย เช่น "25 ก.พ. 2026 12:30"
# ต้องแปลงให้ Python อ่านได้

TH_MONTHS = {
    'ม.ค.':'01', 'ก.พ.':'02', 'มี.ค.':'03', 'เม.ย.':'04',
    'พ.ค.':'05', 'มิ.ย.':'06', 'ก.ค.':'07', 'ส.ค.':'08',
    'ก.ย.':'09', 'ต.ค.':'10', 'พ.ย.':'11', 'ธ.ค.':'12'
}

def parse_thai_date(text):
    """
    แปลงวันที่ภาษาไทยเป็น Python Timestamp
    
    ตัวอย่าง:
    Input:  '25 ก.พ. 2026 12:30'
    Output: Timestamp('2026-02-25 12:30:00')
    """
    try:
        s = str(text)
        # แทนที่เดือนไทยด้วยตัวเลข
        for thai, num in TH_MONTHS.items():
            s = s.replace(thai, num)
        parts = s.strip().split()
        day, mon, year = parts[0], parts[1], int(parts[2])
        # ถ้าปีเป็น พ.ศ. ให้แปลงเป็น ค.ศ.
        if year > 2500:
            year -= 543
        time_part = parts[3] if len(parts) > 3 else '00:00'
        return pd.Timestamp(f"{year}-{mon}-{day} {time_part}")
    except:
        return pd.NaT  # Not a Time = ไม่สามารถแปลงได้


# ============================================================
# ส่วนที่ 2: ฟังก์ชันวิเคราะห์ข้อมูล (หัวใจของโปรแกรม)
# ============================================================

def analyze_data(files):
    """
    รับ dict ของไฟล์ Excel แล้ววิเคราะห์ข้อมูลทั้งหมด
    คืนค่าเป็น dict ของผลลัพธ์
    
    files = {
        'profit': <file object>,  # ไฟล์กำไรคำสั่งซื้อ
        'sku':    <file object>,  # ไฟล์รายงาน SKU
    }
    """
    result = {}  # dict เก็บผลลัพธ์ทั้งหมด

    # ----------------------------------------
    # 2.1 อ่านและประมวลผลไฟล์กำไรคำสั่งซื้อ
    # ----------------------------------------
    if 'profit' in files:
        try:
            # อ่าน Excel โดยไม่ใช้ header อัตโนมัติ
            df_raw = pd.read_excel(files['profit'], header=None)

            # กำหนด column names จากแถวแรก
            df_raw.columns = df_raw.iloc[0]

            # ตัดแถว header ออก เริ่มข้อมูลจริงที่แถว 2
            df_raw = df_raw.iloc[2:].reset_index(drop=True)

            # ตั้งชื่อ column ให้ใช้งานง่าย
            df_raw.columns = [
                'order_id', 'shop', 'revenue', 'sku', 'qty', 'gift',
                'cost', 'profit', 'margin', 'sale_price', 'ship_buyer',
                'discount', 'commission', 'txn_fee', 'service_fee',
                'ship_seller', 'marketing', 'refund', 'platform_fee',
                'order_time', 'confirm_time', 'pay_time', 'update_time',
                'finish_time', 'status', 'receive_pay_time', 'item_id'
            ]

            # แปลง column ตัวเลข
            for col in ['profit', 'qty', 'sale_price', 'cost']:
                df_raw[col] = pd.to_numeric(df_raw[col], errors='coerce')

            # กรองเฉพาะแถวที่มีข้อมูลจริง
            df = df_raw[
                df_raw['sku'].notna() &
                df_raw['profit'].notna() &
                ~df_raw['sku'].astype(str).str.contains('\n', na=False)
            ].copy()

            # --- คำนวณ KPI หลัก ---
            total_profit  = float(df['profit'].sum())
            total_revenue = float(df['sale_price'].sum())

            result['summary'] = {
                'total_qty':     int(df['qty'].sum()),
                'total_profit':  round(total_profit, 0),
                'total_orders':  int(df['order_id'].nunique()),
                'avg_margin':    round(total_profit / total_revenue * 100, 1),
                'total_revenue': round(total_revenue, 0),
            }

            # --- วิเคราะห์ SKU ---
            # groupby = จัดกลุ่มตาม SKU แล้วคำนวณสถิติ
            sku = df.groupby('sku').agg(
                qty     = ('qty', 'sum'),
                profit  = ('profit', 'sum'),
                revenue = ('sale_price', 'sum')
            ).reset_index()

            # คำนวณ Margin และกำไรต่อชิ้น
            sku['margin_pct']      = (sku['profit'] / sku['revenue'] * 100).round(2)
            sku['profit_per_unit'] = (sku['profit'] / sku['qty']).round(2)
            sku['max_cpa']         = (sku['profit_per_unit'] * 0.5).round(2)
            # max_cpa = ค่าโฆษณาสูงสุดต่อออเดอร์ที่ยังคุ้มทุน

            # --- วิเคราะห์ช่องทาง ---
            ch = df.groupby('shop').agg(
                orders  = ('order_id', 'count'),
                qty     = ('qty', 'sum'),
                profit  = ('profit', 'sum'),
                revenue = ('sale_price', 'sum')
            ).reset_index()

            ch['profit_per_order'] = (ch['profit'] / ch['orders']).round(2)
            ch['pct'] = (ch['qty'] / ch['qty'].sum() * 100).round(1)
            result['channels'] = ch.sort_values('qty', ascending=False).to_dict('records')

            # --- วิเคราะห์รายวัน ---
            df['date_parsed'] = df['order_time'].apply(parse_thai_date)
            df['date_only']   = df['date_parsed'].dt.date

            daily = df.groupby('date_only').agg(
                qty     = ('qty', 'sum'),
                profit  = ('profit', 'sum'),
                revenue = ('sale_price', 'sum')
            ).reset_index()

            daily['margin']   = (daily['profit'] / daily['revenue'] * 100).round(1)
            daily['date_str'] = daily['date_only'].astype(str)
            daily = daily.sort_values('date_only')

            # แปลงเป็น list ของ dict เพื่อส่งไป JavaScript
            result['daily'] = daily[['date_str', 'qty', 'profit', 'margin']].round(2).to_dict('records')

            # เก็บ DataFrame ไว้ใช้ merge กับ SKU report
            result['_sku_df'] = sku

        except Exception as e:
            result['error_profit'] = str(e)

    # ----------------------------------------
    # 2.2 อ่านและประมวลผลไฟล์รายงาน SKU
    # ----------------------------------------
    if 'sku' in files:
        try:
            sr = pd.read_excel(files['sku'])

            # กรองแถวสรุปออก
            sr = sr[sr['ชื่อSKU'] != 'ทั้งหมด'].copy()

            # แปลง column สต็อก
            sr['stock']       = pd.to_numeric(sr['สต็อกพร้อมขาย'],               errors='coerce').fillna(0)
            sr['days_stock']  = pd.to_numeric(sr['จำนวนวันที่พร้อมขาย'],          errors='coerce').fillna(0)
            sr['daily_sales'] = pd.to_numeric(sr['เฉลี่ยรายวันการขาย Stock-Out'], errors='coerce').fillna(0)
            sr = sr.rename(columns={'ชื่อSKU': 'sku'})

            result['_sr_df'] = sr

        except Exception as e:
            result['error_sku'] = str(e)

    # ----------------------------------------
    # 2.3 Merge ข้อมูลสองไฟล์เข้าด้วยกัน
    # ----------------------------------------
    sku = result.pop('_sku_df', None)  # ดึง DataFrame จาก result
    sr  = result.pop('_sr_df', None)

    if sku is not None and sr is not None:
        # merge = รวมตารางโดยใช้ column 'sku' เป็นตัวเชื่อม
        # how='left' = เอาทุกแถวจาก sku แม้ไม่มีใน sr
        m = sku.merge(sr[['sku', 'stock', 'days_stock', 'daily_sales']], on='sku', how='left')
        m[['stock', 'days_stock', 'daily_sales']] = m[['stock', 'days_stock', 'daily_sales']].fillna(0)

        result['summary']['total_sku']      = len(m)
        result['summary']['stock_critical'] = int((m['days_stock'] < 7).sum())

        # Top 10 ขายดี — เรียงจากมากไปน้อยตาม qty
        result['top10_sales'] = m.nlargest(10, 'qty')[
            ['sku', 'qty', 'profit', 'margin_pct', 'profit_per_unit', 'stock', 'days_stock']
        ].round(2).to_dict('records')

        # Top 10 กำไร — เรียงจากมากไปน้อยตาม profit
        result['top10_profit'] = m.nlargest(10, 'profit')[
            ['sku', 'qty', 'profit', 'margin_pct', 'profit_per_unit']
        ].round(2).to_dict('records')

        # Margin ต่ำแต่ขายดี
        result['low_margin'] = m[
            (m['margin_pct'] < 27) &      # margin ต่ำกว่า 27%
            (m['qty'] >= 30) &             # ขายได้อย่างน้อย 30 ชิ้น
            (m['profit_per_unit'] < 25)    # กำไรต่อชิ้นต่ำกว่า 25 บาท
        ].nlargest(10, 'qty')[
            ['sku', 'qty', 'profit', 'margin_pct', 'profit_per_unit']
        ].round(2).to_dict('records')

        # แนะนำยิง Ads
        result['ads'] = m[
            (m['stock'] >= 100) &         # สต็อกเพียงพอ
            (m['days_stock'] >= 20) &      # เหลืออย่างน้อย 20 วัน
            (m['margin_pct'] >= 25) &      # margin ดี
            (m['qty'] >= 30)               # ขายได้สม่ำเสมอ
        ].nlargest(10, 'qty')[
            ['sku', 'qty', 'margin_pct', 'max_cpa', 'stock', 'days_stock']
        ].round(2).to_dict('records')

        # ควรหยุดขาย
        result['stop'] = m[
            (m['qty'] <= 3) &             # ขายได้น้อยมาก
            (m['stock'] > 500)             # แต่สต็อกเยอะ
        ].nlargest(10, 'stock')[
            ['sku', 'qty', 'margin_pct', 'stock', 'days_stock']
        ].round(2).to_dict('records')

        # ควรสต็อกเพิ่ม
        result['restock'] = m[
            (m['qty'] >= 30) &            # ขายดี
            (m['days_stock'] < 45)         # สต็อกเหลือน้อย
        ].nlargest(10, 'qty')[
            ['sku', 'qty', 'stock', 'days_stock', 'daily_sales']
        ].round(2).to_dict('records')

    elif sku is not None:
        result['summary']['total_sku'] = len(sku)
        result['top10_sales']  = sku.nlargest(10, 'qty')[['sku', 'qty', 'profit', 'margin_pct', 'profit_per_unit']].round(2).to_dict('records')
        result['top10_profit'] = sku.nlargest(10, 'profit')[['sku', 'qty', 'profit', 'margin_pct', 'profit_per_unit']].round(2).to_dict('records')

    return result


# ============================================================
# ส่วนที่ 3: Routes (เส้นทาง URL)
# ============================================================

# Route หลัก — แสดงหน้าแรก
@app.route('/')
def index():
    """
    เมื่อเปิด http://localhost:5050
    Flask จะ render ไฟล์ templates/index.html
    """
    return render_template('index.html')


# Route รับไฟล์และวิเคราะห์
@app.route('/analyze', methods=['POST'])
def run_analysis():
    """
    รับ POST request พร้อมไฟล์ Excel
    วิเคราะห์และส่งผลลัพธ์กลับเป็น JSON
    
    methods=['POST'] = รับเฉพาะ POST request
    """
    try:
        # รับไฟล์จาก form
        files = {key: request.files[key].stream for key in request.files}

        # วิเคราะห์ข้อมูล
        result = analyze_data(files)

        # ส่งกลับเป็น JSON
        return jsonify({'success': True, 'data': result})

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


# ============================================================
# ส่วนที่ 4: เริ่มโปรแกรม
# ============================================================

# if __name__ == '__main__' = รันเฉพาะเมื่อเรียกไฟล์นี้โดยตรง
# ไม่รันเมื่อ import จากไฟล์อื่น
if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5050))
    print("\n" + "=" * 50)
    print("  ShowMe Thailand — Business Dashboard")
    print(f"  เปิด browser: http://localhost:{port}")
    print("=" * 50 + "\n")
    app.run(debug=False, port=port, host='0.0.0.0')
