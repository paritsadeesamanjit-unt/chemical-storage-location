import streamlit as st
import pandas as pd
import numpy as np
import io

# ---------------------------------------------------------
# Page Configuration
# ---------------------------------------------------------
st.set_page_config(
    page_title="ระบบวิเคราะห์ประวัติการรับเข้าสารเคมี (Chemical Stock & All Receipts)",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------------------------------------------------
# Custom Styling
# ---------------------------------------------------------
st.markdown("""
<style>
    .metric-box {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 12px 18px;
        border-left: 5px solid #28a745;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    .stDataFrame {
        border-radius: 8px;
        overflow: hidden;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# Helper Functions: Data Processing
# ---------------------------------------------------------
@st.cache_data
def load_stock_data(uploaded_file):
    """โหลดข้อมูลไฟล์สต๊อกหลัก อ่านทุก Sheet อัตโนมัติ"""
    try:
        excel_obj = pd.ExcelFile(uploaded_file)
        dfs = []
        for sheet in excel_obj.sheet_names:
            df_s = pd.read_excel(uploaded_file, sheet_name=sheet)
            df_s.columns = [str(c).strip() for c in df_s.columns]
            df_s['Source_Sheet'] = sheet
            dfs.append(df_s)
            
        combined_df = pd.concat(dfs, ignore_index=True)
        
        rename_map = {
            'Material': 'Material Num',
            'Material No': 'Material Num',
            'Material Number': 'Material Num',
            'รหัสวัสดุ': 'Material Num',
            'Material Description': 'Material Desc',
            'ชื่อวัสดุ': 'Material Desc',
            'Storage Location': 'Storage location'
        }
        combined_df.rename(columns=rename_map, inplace=True)
        
        if 'Material Num' in combined_df.columns:
            combined_df['Material Num'] = combined_df['Material Num'].astype(str).str.strip()
            
        if 'Unrestricted' in combined_df.columns:
            combined_df['Unrestricted'] = pd.to_numeric(combined_df['Unrestricted'], errors='coerce').fillna(0)
            
        return combined_df
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการโหลดไฟล์สต๊อก: {e}")
        return None


@st.cache_data
def load_movement_data(uploaded_file):
    """โหลดข้อมูลไฟล์การเคลื่อนไหวรับเข้า (Goods Receipt)"""
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [str(c).strip() for c in df.columns]
        
        rename_map = {
            'Material Num': 'Material',
            'Material Number': 'Material',
            'รหัสวัสดุ': 'Material',
            'Qty in unit of entry': 'GR_Qty',
            'Posting Date': 'Posting_Date',
            'Document Date': 'Document_Date',
            'Purchase order': 'PO_Number',
            'Material Document': 'Mat_Document'
        }
        df.rename(columns=rename_map, inplace=True)
        
        if 'Material' in df.columns:
            df['Material'] = df['Material'].astype(str).str.strip()
            
        if 'GR_Qty' in df.columns:
            df['GR_Qty'] = pd.to_numeric(df['GR_Qty'], errors='coerce').fillna(0)
            
        # แปลงวันที่
        for date_col in ['Posting_Date', 'Document_Date', 'Entry Date']:
            if date_col in df.columns:
                df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
                
        # ปรับเลข PO ให้เป็นตัวเลขไม่มีทศนิยม
        if 'PO_Number' in df.columns:
            df['PO_Number'] = df['PO_Number'].fillna('-').astype(str).str.replace(r'\.0$', '', regex=True)
            
        return df
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการโหลดไฟล์ Movement: {e}")
        return None


def process_all_receipts_data(df_stock, df_movement):
    """
    เชื่อมโยงข้อมูลสต๊อกกับประวัติการรับเข้าทั้งหมด (All Receipts)
    คืนค่าทั้ง:
    1. รายการรับเข้าทั้งหมดแบบละเอียด (1 บรรทัดต่อ 1 การรับเข้า)
    2. สรุปรวมประวัติทุก PO/วันที่ ในบรรทัดเดียว (Grouped by Material)
    3. ข้อมูลสรุปภาพรวม (Summary)
    """
    chem_codes = set(df_stock['Material Num'].dropna().unique())
    df_gr_chem = df_movement[df_movement['Material'].isin(chem_codes)].copy()
    
    date_col = 'Posting_Date' if 'Posting_Date' in df_gr_chem.columns else 'Document_Date'
    
    # -------------------------------------------------------------
    # 1. รายการรับเข้าทั้งหมด (All Transactions: 1 บรรทัดต่อ 1 การรับเข้า)
    # -------------------------------------------------------------
    # นำข้อมูล Master Stock ไปผูกเข้ากับ Transaction ทุกรายการ
    merged_all = pd.merge(
        df_stock,
        df_gr_chem,
        left_on='Material Num',
        right_on='Material',
        how='left'
    )
    
    # เรียงลำดับตาม รหัสสารเคมี และ วันที่รับเข้าล่าสุดลงไปหาเก่าสุด
    merged_all.sort_values(by=['Material Num', date_col, 'Mat_Document'], ascending=[True, False, False], inplace=True)
    
    # คำนวณลำดับรอบที่รับเข้า (เช่น รับครั้งที่ 1 จาก 4 ครั้ง)
    merged_all['Receipt_Seq'] = merged_all.groupby('Material Num').cumcount() + 1
    total_counts = merged_all.groupby('Material Num')['Receipt_Seq'].transform('count')
    merged_all['Receipt_Total'] = total_counts
    merged_all['รอบที่รับ'] = "รอบที่ " + merged_all['Receipt_Seq'].astype(str) + " / " + merged_all['Receipt_Total'].astype(str)
    
    # -------------------------------------------------------------
    # 2. มุมมองจัดกลุ่ม: รวมทุกประวัติรับเข้าไว้ในแถวเดียว (Grouped All History)
    # -------------------------------------------------------------
    def format_all_receipts(group):
        g_sorted = group.sort_values(by=date_col, ascending=False)
        receipt_items = []
        for _, r in g_sorted.iterrows():
            if pd.notna(r[date_col]):
                d_str = r[date_col].strftime('%Y-%m-%d')
                po = r.get('PO_Number', '-')
                qty = r.get('GR_Qty', 0)
                receipt_items.append(f"• {d_str} | PO: {po} | จำนวน: {qty:,.2f}")
        return "\n".join(receipt_items) if receipt_items else "ไม่มีข้อมูลรับเข้า"

    grouped_history = df_gr_chem.groupby('Material').apply(format_all_receipts).reset_index(name='ประวัติการรับเข้าทั้งหมด')
    
    # คำนวณยอดรวม
    summary_agg = df_gr_chem.groupby('Material').agg(
        Total_GR_Qty=('GR_Qty', 'sum'),
        Receipt_Count=('GR_Qty', 'count'),
        Latest_Date=(date_col, 'max')
    ).reset_index()
    
    df_grouped_view = pd.merge(df_stock, summary_agg, left_on='Material Num', right_on='Material', how='left')
    df_grouped_view = pd.merge(df_grouped_view, grouped_history, left_on='Material Num', right_on='Material', how='left')
    df_grouped_view.drop(columns=['Material_x', 'Material_y'], inplace=True, errors='ignore')
    
    return merged_all, df_grouped_view, df_gr_chem, date_col


def convert_df_to_excel(df):
    """แปลง DataFrame เป็นไฟล์ Excel ในหน่วยความจำเพื่อดาวน์โหลด"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()


# ---------------------------------------------------------
# Sidebar: Upload & Filters
# ---------------------------------------------------------
st.sidebar.title("🧪 ตัวจัดการข้อมูลสารเคมี")
st.sidebar.markdown("อัปโหลดไฟล์ข้อมูลเพื่อประมวลผล:")

stock_file = st.sidebar.file_uploader("1. ไฟล์สต๊อกปัจจุบัน (Excel)", type=["xlsx", "xls"], key="stock_file")
movement_file = st.sidebar.file_uploader("2. ไฟล์ประวัติการรับเข้า EXPORT (Excel)", type=["xlsx", "xls"], key="mov_file")

st.sidebar.markdown("---")
use_sample = st.sidebar.checkbox("📂 ใช้ข้อมูลในระบบอัตโนมัติ (Default Files)", value=(stock_file is None and movement_file is None))

df_stock = None
df_movement = None

if stock_file is not None:
    df_stock = load_stock_data(stock_file)
elif use_sample:
    try:
        df_stock = load_stock_data("Chemical Storage Location.XLSX")
    except:
        pass

if movement_file is not None:
    df_movement = load_movement_data(movement_file)
elif use_sample:
    try:
        df_movement = load_movement_data("EXPORT1.XLSX")
    except:
        pass


# ---------------------------------------------------------
# Main Page Content
# ---------------------------------------------------------
st.title("📊 รายงานข้อมูลประวัติการรับเข้าสารเคมีทั้งหมด")
st.caption("แสดงข้อมูลวันที่รับเข้า, เลขที่ PO และจำนวนที่รับเข้าครบทุกรายการ (All Receiving Records)")

if df_stock is None or df_movement is None:
    st.info("👋 กรุณาอัปโหลดไฟล์สต๊อกและไฟล์การเคลื่อนไหวทางแถบเมนูด้านซ้ายเพื่อเริ่มประมวลผล")
    st.stop()

# ประมวลผลข้อมูล
merged_all, df_grouped_view, df_gr_chem, date_col = process_all_receipts_data(df_stock, df_movement)

# สรุปตัวเลขสถิติภาพรวมด้านบน
c1, c2, c3, c4 = st.columns(4)
c1.metric("จำนวนสารเคมีในสต๊อก", f"{len(df_stock):,} รายการ")
c2.metric("ยอดสต๊อกคงเหลือรวม", f"{df_stock['Unrestricted'].sum():,.2f}")
c3.metric("จำนวนครั้งที่รับเข้าทั้งหมด", f"{len(df_gr_chem):,} ครั้ง")
c4.metric("ยอดรับเข้ารวมทั้งหมด", f"{df_gr_chem['GR_Qty'].sum():,.2f}")

st.markdown("---")

# ---------------------------------------------------------
# ส่วนเลือกมุมมองการแสดงผล (Display Modes)
# ---------------------------------------------------------
st.subheader("📑 ตารางข้อมูลการรับเข้าสารเคมี")

view_mode = st.radio(
    "เลือกรูปแบบการแสดงข้อมูล:",
    [
        "1. แสดงประวัติการรับเข้าทุกรายการ (1 บรรทัดต่อ 1 การรับเข้า - รายละเอียดครบทุก PO และวันที่)",
        "2. แสดงประวัติรวมทุกการรับเข้าในแถวเดียว (1 บรรทัดต่อ 1 สารเคมี รวมประวัติ PO ทั้งหมด)"
    ],
    horizontal=True
)

# กล่องค้นหาและตัวกรองข้อมูล
f1, f2, f3 = st.columns([1.2, 1.5, 1.3])
with f1:
    loc_options = ["ทั้งหมด"] + sorted(list(df_stock['Storage location'].dropna().unique()))
    selected_loc = st.selectbox("กรองตามสถานที่จัดเก็บ:", loc_options)
with f2:
    search_kw = st.text_input("ค้นหารหัสสารเคมี / ชื่อสารเคมี:", placeholder="เช่น T11-, BO-7790...")
with f3:
    po_search = st.text_input("ค้นหาเลขที่ PO:", placeholder="เช่น 6700195641...")

# ---------------------------------------------------------
# รูปแบบที่ 1: รายการรับเข้าทั้งหมด (All Transactions)
# ---------------------------------------------------------
if "1. แสดงประวัติการรับเข้าทุกรายการ" in view_mode:
    filtered_df = merged_all.copy()
    
    if selected_loc != "ทั้งหมด":
        filtered_df = filtered_df[filtered_df['Storage location'] == selected_loc]
    if search_kw:
        filtered_df = filtered_df[
            filtered_df['Material Num'].str.contains(search_kw, case=False, na=False) |
            filtered_df['Material Desc'].str.contains(search_kw, case=False, na=False)
        ]
    if po_search:
        filtered_df = filtered_df[filtered_df['PO_Number'].str.contains(po_search, na=False)]

    # จัดเตรียมคอลัมน์แสดงผล
    display_cols = [
        'Material Num', 'Material Desc', 'Storage location', 'Unrestricted', 'Unit',
        'รอบที่รับ', date_col, 'PO_Number', 'GR_Qty', 'Mat_Document', 'Batch', 'Vendor'
    ]
    actual_cols = [c for c in display_cols if c in filtered_df.columns]
    
    show_df = filtered_df[actual_cols].copy()
    if date_col in show_df.columns:
        show_df[date_col] = show_df[date_col].dt.strftime('%Y-%m-%d').fillna('-')
        
    show_df.rename(columns={
        'Material Num': 'รหัสสารเคมี',
        'Material Desc': 'ชื่อสารเคมี',
        'Storage location': 'สถานที่จัดเก็บ',
        'Unrestricted': 'สต๊อกคงเหลือปัจจุบัน',
        'Unit': 'หน่วย',
        'รอบที่รับ': 'รอบที่รับ',
        date_col: 'วันที่รับเข้า',
        'PO_Number': 'เลขที่ PO',
        'GR_Qty': 'จำนวนที่รับเข้า',
        'Mat_Document': 'เลขที่เอกสารรับ (Mat Doc)',
        'Batch': 'ล็อต (Batch)',
        'Vendor': 'รหัสผู้ขาย (Vendor)'
    }, inplace=True)

    st.markdown(f"**แสดงผลทั้งหมด: `{len(show_df):,}` รายการรับเข้า** (เรียงลำดับจากวันที่รับล่าสุดไปหาเก่าสุด)")
    st.dataframe(show_df, use_container_width=True, height=520)

    # ปุ่มดาวน์โหลด Excel
    excel_data = convert_df_to_excel(show_df)
    st.download_button(
        label="📥 ดาวน์โหลดประวัติการรับเข้าทั้งหมดเป็น Excel",
        data=excel_data,
        file_name="Chemical_All_Goods_Receipts.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ---------------------------------------------------------
# รูปแบบที่ 2: สรุปใน 1 แถวต่อสารเคมี (Grouped All Receipts in One Row)
# ---------------------------------------------------------
else:
    filtered_grouped = df_grouped_view.copy()
    
    if selected_loc != "ทั้งหมด":
        filtered_grouped = filtered_grouped[filtered_grouped['Storage location'] == selected_loc]
    if search_kw:
        filtered_grouped = filtered_grouped[
            filtered_grouped['Material Num'].str.contains(search_kw, case=False, na=False) |
            filtered_grouped['Material Desc'].str.contains(search_kw, case=False, na=False)
        ]
    if po_search:
        filtered_grouped = filtered_grouped[filtered_grouped['ประวัติการรับเข้าทั้งหมด'].str.contains(po_search, na=False)]

    disp_grouped = filtered_grouped[[
        'Material Num', 'Material Desc', 'Storage location', 'Unit', 
        'Unrestricted', 'Total_GR_Qty', 'Receipt_Count', 'ประวัติการรับเข้าทั้งหมด'
    ]].copy()
    
    disp_grouped.rename(columns={
        'Material Num': 'รหัสสารเคมี',
        'Material Desc': 'ชื่อสารเคมี',
        'Storage location': 'สถานที่จัดเก็บ',
        'Unit': 'หน่วย',
        'Unrestricted': 'สต๊อกปัจจุบัน',
        'Total_GR_Qty': 'ยอดรับสะสมรวม',
        'Receipt_Count': 'จำนวนครั้งที่รับ',
        'ประวัติการรับเข้าทั้งหมด': 'ประวัติการรับเข้าทั้งหมด (วันที่ | PO | จำนวน)'
    }, inplace=True)

    st.markdown(f"**แสดงผลทั้งหมด: `{len(disp_grouped):,}` ชนิดสารเคมี**")
    st.dataframe(disp_grouped, use_container_width=True, height=520)

    excel_grouped = convert_df_to_excel(disp_grouped)
    st.download_button(
        label="📥 ดาวน์โหลดตารางสรุปเป็น Excel",
        data=excel_grouped,
        file_name="Chemical_Grouped_All_Receipts.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
