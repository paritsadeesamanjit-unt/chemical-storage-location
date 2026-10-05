import streamlit as st
import pandas as pd
import numpy as np
import io

# ---------------------------------------------------------
# Page Configuration
# ---------------------------------------------------------
st.set_page_config(
    page_title="ระบบวิเคราะห์การรับเข้าสารเคมีและสต๊อก (Chemical Inventory & GR Analytics)",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------------------------------------------------
# Custom Styling
# ---------------------------------------------------------
st.markdown("""
<style>
    .metric-card {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 15px;
        border-left: 5px solid #1f77b4;
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
            # ตรวจสอบชื่อคอลัมน์มาตรฐาน
            df_s.columns = [str(c).strip() for c in df_s.columns]
            df_s['Source_Sheet'] = sheet
            dfs.append(df_s)
            
        combined_df = pd.concat(dfs, ignore_index=True)
        
        # จัดการชื่อคอลัมน์ให้อยู่ในมาตรฐานเดียวกัน
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
        
        # ปรับมาตรฐานชื่อคอลัมน์
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
                
        # แปลง PO Number ให้แสดงผลเป็นรหัสตัวเลขชัดเจน
        if 'PO_Number' in df.columns:
            df['PO_Number'] = df['PO_Number'].fillna('-').astype(str).str.replace(r'\.0$', '', regex=True)
            
        return df
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการโหลดไฟล์ Movement: {e}")
        return None


def calculate_analytics(df_stock, df_movement):
    """คำนวณเชื่อมโยงข้อมูลระหว่างสต๊อกและประวัติการรับเข้า"""
    # 1. กรองเฉพาะรายการสารเคมีที่มีอยู่ในไฟล์สต๊อกหลัก
    chem_codes = set(df_stock['Material Num'].dropna().unique())
    df_gr_chem = df_movement[df_movement['Material'].isin(chem_codes)].copy()
    
    # เรียงลำดับตามรหัสและวันที่รับเข้าล่าสุด
    date_col = 'Posting_Date' if 'Posting_Date' in df_gr_chem.columns else 'Document_Date'
    df_gr_chem.sort_values(by=['Material', date_col, 'Mat_Document'], ascending=[True, True, True], inplace=True)
    
    # 2. คำนวณค่า Aggregate สรุปรวมต่อรหัสวัสดุ
    summary_agg = df_gr_chem.groupby('Material').agg(
        Total_GR_Qty=('GR_Qty', 'sum'),
        Receipt_Count=('GR_Qty', 'count'),
        Latest_GR_Date=(date_col, 'max'),
        Earliest_GR_Date=(date_col, 'min')
    ).reset_index()
    
    # 3. ดึงข้อมูลครั้งล่าสุด (Latest Receipt Detail: PO ล่าสุด, จำนวนล่าสุด, เลขเอกสารล่าสุด)
    latest_tx = df_gr_chem.groupby('Material').last().reset_index()
    latest_subset = latest_tx[['Material', 'PO_Number', 'GR_Qty', 'Mat_Document']].copy()
    latest_subset.columns = ['Material', 'Latest_PO', 'Latest_GR_Qty', 'Latest_Mat_Doc']
    
    # รวมข้อมูลสรุปล่าสุดเข้าด้วยกัน
    gr_analytics = pd.merge(summary_agg, latest_subset, on='Material', how='left')
    
    # 4. นำไป Merge กับ Master Stock Data
    master_summary = pd.merge(df_stock, gr_analytics, left_on='Material Num', right_on='Material', how='left')
    master_summary.drop(columns=['Material'], inplace=True, errors='ignore')
    
    # จัดการค่า Null สำหรับตัวที่ยังไม่เคยมีการรับเข้า
    master_summary['Total_GR_Qty'] = master_summary['Total_GR_Qty'].fillna(0)
    master_summary['Receipt_Count'] = master_summary['Receipt_Count'].fillna(0).astype(int)
    master_summary['Latest_PO'] = master_summary['Latest_PO'].fillna('-')
    master_summary['Latest_GR_Qty'] = master_summary['Latest_GR_Qty'].fillna(0)
    
    # 5. ตารางประวัติรับเข้าแบบละเอียดทั้งหมด (Detail Transaction)
    # นำข้อมูลจาก Master Stock เช่น Location และ Stock คงเหลือปัจจุบัน มารวมใน Transaction
    stock_info = df_stock[['Material Num', 'Material Desc', 'Storage location', 'Unit', 'Unrestricted']].drop_duplicates(subset=['Material Num'])
    detail_merged = pd.merge(df_gr_chem, stock_info, left_on='Material', right_on='Material Num', how='left')
    
    return master_summary, detail_merged, df_gr_chem


def convert_df_to_excel(df):
    """แปลง DataFrame เป็นไฟล์ Excel ในหน่วยความจำเพื่อดาวน์โหลด"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
    return output.getvalue()


# ---------------------------------------------------------
# Sidebar: Upload & Settings
# ---------------------------------------------------------
st.sidebar.title("🧪 ตัวจัดการข้อมูลสารเคมี")
st.sidebar.markdown("อัปโหลดไฟล์ข้อมูลเพื่อประมวลผล:")

stock_file = st.sidebar.file_uploader("1. ไฟล์สต๊อกปัจจุบัน (Excel)", type=["xlsx", "xls"], key="stock_file")
movement_file = st.sidebar.file_uploader("2. ไฟล์ประวัติการรับเข้า EXPORT (Excel)", type=["xlsx", "xls"], key="mov_file")

st.sidebar.markdown("---")
use_sample = st.sidebar.checkbox("📂 ใช้ข้อมูลในระบบอัตโนมัติ (Default Files)", value=(stock_file is None and movement_file is None))

# ตรวจสอบการโหลดไฟล์
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
# Main Application Content
# ---------------------------------------------------------
st.title("📊 ระบบวิเคราะห์ประวัติการรับเข้าและสต๊อกสารเคมี")
st.caption("Chemical Inventory & Goods Receipt (PO / Receiving History) Analytics System")

if df_stock is None or df_movement is None:
    st.info("👋 กรุณาอัปโหลดไฟล์ **1. ไฟล์สต๊อกสารเคมีปัจจุบัน** และ **2. ไฟล์การเคลื่อนไหวรับเข้า (EXPORT)** ทางแถบเมนูด้านซ้ายเพื่อเริ่มการวิเคราะห์")
    st.stop()

# ประมวลผลข้อมูล
master_summary, detail_merged, df_gr_chem = calculate_analytics(df_stock, df_movement)

# KPI Cards
total_items = len(master_summary)
total_current_stock = master_summary['Unrestricted'].sum()
total_received_qty = master_summary['Total_GR_Qty'].sum()
total_gr_records = len(df_gr_chem)
unique_pos = df_gr_chem['PO_Number'].nunique() if 'PO_Number' in df_gr_chem.columns else 0

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("จำนวนสารเคมีทั้งหมด", f"{total_items:,} รายการ")
col2.metric("ยอดสต๊อกคงเหลือรวม", f"{total_current_stock:,.2f}")
col3.metric("ยอดรับเข้าสะสมทั้งหมด", f"{total_received_qty:,.2f}")
col4.metric("ครั้งที่รับเข้าทั้งหมด", f"{total_gr_records:,} ครั้ง")
col5.metric("จำนวน PO ทั้งหมด", f"{unique_pos:,} ใบ")

st.markdown("---")

# Navigation Tabs
tab1, tab2, tab3, tab4 = st.tabs([
    "📋 สรุปสต๊อก & การรับเข้าล่าสุด (Summary)",
    "📑 ประวัติการรับเข้าทั้งหมด (All Receipts Detail)",
    "🔍 เจาะลึกรายสารเคมี (Item Drilldown)",
    "📦 วิเคราะห์ตามใบสั่งซื้อ (PO Analysis)"
])

# ---------------------------------------------------------
# TAB 1: สรุปภาพรวมและข้อมูลการรับเข้าล่าสุด (1 แถวต่อ 1 สารเคมี)
# ---------------------------------------------------------
with tab1:
    st.subheader("📋 ภาพรวมสต๊อกสารเคมีและข้อมูลการรับเข้าล่าสุด")
    st.markdown("ตารางนี้เชื่อมโยงสารเคมีทุกตัวกับยอดรับรวมสะสม พร้อม **วันที่รับเข้าล่าสุด**, **เลขที่ PO ล่าสุด** และ **จำนวนที่รับเข้าล่าสุด**")
    
    # ตัวกรองข้อมูล
    f_col1, f_col2 = st.columns([1, 2])
    with f_col1:
        loc_options = ["ทั้งหมด"] + sorted(list(master_summary['Storage location'].dropna().unique()))
        selected_loc = st.selectbox("กรองตามพื้นที่จัดเก็บ (Storage Location):", loc_options)
    with f_col2:
        search_kw = st.text_input("ค้นหารหัสหรือชื่อสารเคมี:", placeholder="พิมพ์รหัส เช่น T12- หรือชื่อสารเคมี...")

    filtered_summary = master_summary.copy()
    if selected_loc != "ทั้งหมด":
        filtered_summary = filtered_summary[filtered_summary['Storage location'] == selected_loc]
    if search_kw:
        filtered_summary = filtered_summary[
            filtered_summary['Material Num'].str.contains(search_kw, case=False, na=False) |
            filtered_summary['Material Desc'].str.contains(search_kw, case=False, na=False)
        ]

    # เตรียมตารางแสดงผลสวยงาม
    display_cols = [
        'Material Num', 'Material Desc', 'Storage location', 'Unit', 
        'Unrestricted', 'Total_GR_Qty', 'Receipt_Count', 
        'Latest_GR_Date', 'Latest_PO', 'Latest_GR_Qty'
    ]
    
    # แปลง Format วันที่ให้ดูง่าย
    display_df = filtered_summary[display_cols].copy()
    display_df['Latest_GR_Date'] = display_df['Latest_GR_Date'].dt.strftime('%Y-%m-%d').fillna('-')
    display_df.rename(columns={
        'Material Num': 'รหัสสารเคมี (Material)',
        'Material Desc': 'ชื่อสารเคมี (Description)',
        'Storage location': 'สถานที่จัดเก็บ',
        'Unit': 'หน่วย',
        'Unrestricted': 'สต๊อกปัจจุบัน',
        'Total_GR_Qty': 'ยอดรับสะสมรวม',
        'Receipt_Count': 'จำนวนครั้งที่รับ',
        'Latest_GR_Date': 'วันที่รับล่าสุด',
        'Latest_PO': 'PO ล่าสุด',
        'Latest_GR_Qty': 'จำนวนที่รับล่าสุด'
    }, inplace=True)

    st.dataframe(display_df, use_container_width=True, height=450)
    
    # ปุ่มดาวน์โหลด Excel
    excel_summary = convert_df_to_excel(display_df)
    st.download_button(
        label="📥 ดาวน์โหลดตารางสรุปเป็น Excel",
        data=excel_summary,
        file_name="Chemical_Stock_and_Latest_GR_Summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ---------------------------------------------------------
# TAB 2: ประวัติการรับเข้าทั้งหมด (All Transaction Details)
# ---------------------------------------------------------
with tab2:
    st.subheader("📑 ประวัติการรับเข้าสารเคมีทุกรายการ (Goods Receipt Transactions)")
    st.markdown("แสดงรายการการรับเข้าแยกตามแต่ละครั้ง (1 Material มีหลายการรับเข้า) พร้อมเลขที่ PO และจำนวนที่รับ")

    # ตัวกรองในแท็บ 2
    c1, c2, c3 = st.columns([1, 1, 1])
    with c1:
        po_filter = st.text_input("ค้นหาเลขที่ PO:", placeholder="เช่น 6700195641...")
    with c2:
        mat_filter = st.text_input("ค้นหารหัสสารเคมี:", placeholder="เช่น T11-90240202...", key="t2_mat")
    with c3:
        if 'Posting_Date' in detail_merged.columns and not detail_merged['Posting_Date'].isna().all():
            min_date = detail_merged['Posting_Date'].min().date()
            max_date = detail_merged['Posting_Date'].max().date()
            date_range = st.date_input("ช่วงวันที่รับเข้า (Posting Date):", [min_date, max_date])
        else:
            date_range = None

    filtered_detail = detail_merged.copy()
    if po_filter:
        filtered_detail = filtered_detail[filtered_detail['PO_Number'].str.contains(po_filter, na=False)]
    if mat_filter:
        filtered_detail = filtered_detail[
            filtered_detail['Material'].str.contains(mat_filter, case=False, na=False) |
            filtered_detail['Material Desc'].str.contains(mat_filter, case=False, na=False)
        ]
    if date_range and len(date_range) == 2:
        start_d, end_d = date_range
        filtered_detail = filtered_detail[
            (filtered_detail['Posting_Date'].dt.date >= start_d) & 
            (filtered_detail['Posting_Date'].dt.date <= end_d)
        ]

    # เลือกคอลัมน์สำคัญมาแสดงผล
    show_cols = [
        'Posting_Date', 'Purchase order' if 'Purchase order' in filtered_detail.columns else 'PO_Number',
        'Material', 'Material Desc', 'GR_Qty', 'Unit of Entry', 'Storage location', 
        'Mat_Document', 'Batch', 'Vendor'
    ]
    actual_cols = [c for c in show_cols if c in filtered_detail.columns]
    
    st.markdown(f"**พบข้อมูลทั้งหมด: {len(filtered_detail):,} รายการ**")
    st.dataframe(filtered_detail[actual_cols].sort_values(by='Posting_Date', ascending=False), use_container_width=True, height=450)

    excel_detail = convert_df_to_excel(filtered_detail[actual_cols])
    st.download_button(
        label="📥 ดาวน์โหลดประวัติการรับเข้าทั้งหมดเป็น Excel",
        data=excel_detail,
        file_name="Chemical_All_GR_History.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ---------------------------------------------------------
# TAB 3: เจาะลึกรายสารเคมี (Item Drilldown & Receiving Timeline)
# ---------------------------------------------------------
with tab3:
    st.subheader("🔍 เจาะลึกข้อมูลและประวัติการรับเข้าของสารเคมีรายตัว")
    
    chem_list = sorted(master_summary['Material Num'].unique())
    selected_chem = st.selectbox(
        "เลือกรหัสสารเคมีที่ต้องการตรวจสอบ:", 
        chem_list,
        format_func=lambda x: f"{x} - {master_summary[master_summary['Material Num'] == x]['Material Desc'].values[0]}"
    )
    
    chem_info = master_summary[master_summary['Material Num'] == selected_chem].iloc[0]
    chem_history = detail_merged[detail_merged['Material'] == selected_chem].sort_values(by='Posting_Date', ascending=False)
    
    # การ์ดแสดงข้อมูลสรุปของสารตัวที่เลือก
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("สต๊อกปัจจุบัน", f"{chem_info['Unrestricted']:,.2f} {chem_info['Unit']}")
    m2.metric("ยอดรับเข้ารวมทั้งหมด", f"{chem_info['Total_GR_Qty']:,.2f} {chem_info['Unit']}")
    m3.metric("จำนวนรอบที่รับเข้า", f"{len(chem_history)} ครั้ง")
    m4.metric("รับเข้าล่าสุด", str(chem_info['Latest_GR_Date'])[:10] if pd.notna(chem_info['Latest_GR_Date']) else "-")

    st.markdown("#### ประวัติการรับเข้าสารเคมีรายการนี้แยกตาม PO และวันที่")
    if not chem_history.empty:
        # แสดงตารางประวัติของสารเคมีตัวนี้
        sub_cols = ['Posting_Date', 'PO_Number', 'GR_Qty', 'Unit of Entry', 'Mat_Document', 'Batch', 'Vendor']
        valid_sub = [c for c in sub_cols if c in chem_history.columns]
        st.dataframe(chem_history[valid_sub], use_container_width=True)
        
        # กราฟแท่งแสดงแนวโน้มจำนวนที่รับเข้าในแต่ละวันที่
        chart_data = chem_history.groupby('Posting_Date')['GR_Qty'].sum().reset_index()
        chart_data.sort_values(by='Posting_Date', inplace=True)
        chart_data.set_index('Posting_Date', inplace=True)
        st.markdown("##### กราฟแสดงจำนวนที่รับเข้าในแต่ละช่วงเวลา:")
        st.bar_chart(chart_data)
    else:
        st.warning("ไม่พบประวัติการรับเข้าสำหรับสารเคมีรายการนี้")

# ---------------------------------------------------------
# TAB 4: วิเคราะห์ตามใบสั่งซื้อ (PO Analysis)
# ---------------------------------------------------------
with tab4:
    st.subheader("📦 สรุปการรับเข้าสารเคมีจัดกลุ่มตามเลขที่ใบสั่งซื้อ (PO Analysis)")
    
    po_summary = df_gr_chem.groupby('PO_Number').agg(
        Total_Qty=('GR_Qty', 'sum'),
        Item_Count=('Material', 'nunique'),
        Receipt_Records=('GR_Qty', 'count'),
        Latest_Date=('Posting_Date', 'max'),
        Vendor=('Vendor', lambda x: x.iloc[0] if len(x) > 0 else '-')
    ).reset_index().sort_values(by='Latest_Date', ascending=False)
    
    po_summary['Latest_Date'] = po_summary['Latest_Date'].dt.strftime('%Y-%m-%d')
    po_summary.rename(columns={
        'PO_Number': 'เลขที่ PO',
        'Total_Qty': 'จำนวนรับรวม',
        'Item_Count': 'จำนวนชนิดสารเคมี',
        'Receipt_Records': 'จำนวนรายการรับ',
        'Latest_Date': 'วันที่รับเข้าล่าสุด',
        'Vendor': 'รหัสผู้ขาย (Vendor)'
    }, inplace=True)
    
    st.dataframe(po_summary, use_container_width=True, height=450)
