# -*- coding: utf-8 -*-
"""
Dashboard Công Suất Sản Xuất
Tính toán theo logic Excel CS máy_SX1 và CS máy_SX2
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import numpy as np
import os

# ============= CẤU HÌNH =============
st.set_page_config(
    page_title="Công Suất Sản Xuất",
    page_icon="📊",
    layout="wide"
)

CONFIG = {
    'google_credentials': 'api-agent-471608-912673253587.json',
    'google_sheet_url': 'https://docs.google.com/spreadsheets/d/1F2NzTR50kXzGx9Pc5KdBwwqnIRXGvViPv6mgw8YMNW0/edit',
    'lathe_machines': ['48', '50', '51', '52', '54', '55', '56', '57', '58', '59', '60', '61']
}

# ============= FUNCTIONS =============

@st.cache_resource
def authenticate_google_sheets():
    """Xác thực Google Sheets"""
    try:
        scopes = ['https://www.googleapis.com/auth/spreadsheets']
        
        # Thử đọc từ Streamlit Secrets (cho môi trường Cloud)
        try:
            if "gcp_service_account" in st.secrets:
                creds = Credentials.from_service_account_info(
                    st.secrets["gcp_service_account"],
                    scopes=scopes
                )
                return gspread.authorize(creds)
            else:
                # Debug: List available keys to help user check config
                st.error("❌ Lỗi: Không tìm thấy 'gcp_service_account' trong Secrets.")
                st.write(f"🔍 Các keys hiện có trong Secrets: {list(st.secrets.keys())}")
                st.info("💡 Vui lòng kiểm tra lại tên header trong Secrets phải là [gcp_service_account]")
                return None
            
        # Nếu không có secret, đọc từ file JSON (môi trường Local)
        if os.path.exists(CONFIG['google_credentials']):
            creds = Credentials.from_service_account_file(
                CONFIG['google_credentials'],
                scopes=scopes
            )
            return gspread.authorize(creds)
        else:
            st.error("❌ Lỗi: Không tìm thấy 'gcp_service_account' trong Secrets và không có file JSON cục bộ.")
            st.info("💡 Vui lòng vào Settings -> Secrets trên Streamlit Cloud và dán cấu hình TOML vào.")
            return None
    except Exception as e:
        st.error(f"❌ Lỗi xác thực: {e}")
        return None

@st.cache_data(ttl=300)
def read_phtcv_data():
    """Đọc dữ liệu PHTCV từ Google Sheets"""
    try:
        client = authenticate_google_sheets()
        if not client:
            return None
        
        spreadsheet = client.open_by_url(CONFIG['google_sheet_url'])
        worksheet = spreadsheet.worksheet('PHTCV')
        data = worksheet.get_all_values()
        
        if data and len(data) > 1:
            df = pd.DataFrame(data[1:], columns=data[0])
            df = df.dropna(axis=0, how='all')
            
            # Convert time columns to numeric
            time_cols = ['tgcb', 'chạy thử', 'gá lắp', 'gia công', 'dừng', 'dừng khác', 'sửa']
            for col in time_cols:
                if col in df.columns:
                    df[col] = pd.to_numeric(
                        df[col].astype(str).str.replace(',', '.'),
                        errors='coerce'
                    ).fillna(0)
            
            # Parse date column
            if 'ngày tháng' in df.columns:
                df['date_parsed'] = pd.to_datetime(df['ngày tháng'], format='%d/%m/%Y', errors='coerce')
            
            return df
        return pd.DataFrame()
    except Exception as e:
        st.error(f"❌ Lỗi đọc PHTCV: {e}")
        return None

@st.cache_data(ttl=300)
def read_machine_list():
    """Đọc danh sách máy từ Google Sheets"""
    try:
        client = authenticate_google_sheets()
        if not client:
            return []
        
        spreadsheet = client.open_by_url(CONFIG['google_sheet_url'])
        worksheet = spreadsheet.worksheet('machine_list')
        data = worksheet.get_all_values()
        
        if data and len(data) > 1:
            df = pd.DataFrame(data[1:], columns=data[0])
            # Assume first column contains machine numbers
            machine_col = df.columns[0]
            machines = df[machine_col].dropna().astype(str).tolist()
            return [m.strip() for m in machines if m.strip()]
        return []
    except Exception as e:
        st.warning(f"⚠️ Không thể đọc machine_list: {e}")
        return []

def calculate_capacity_by_type(df, machine_type='all'):
    """
    Tính công suất theo loại máy (tiện/phay/tất cả)
    Logic: Nhân với sl thực tế: gia công, gá lắp
    Các thành phần khác (chuẩn bị, chạy thử, dừng, sửa) SUM trực tiếp
    """
    lathe_machines = CONFIG['lathe_machines']
    
    # Filter by machine type
    if machine_type == 'lathe':
        df_filtered = df[df['số máy'].isin(lathe_machines)].copy()
    elif machine_type == 'milling':
        df_filtered = df[~df['số máy'].isin(lathe_machines)].copy()
    else:  # all
        df_filtered = df.copy()
    
    # Convert sl thực tế to numeric
    df_filtered['sl_thuc_te'] = pd.to_numeric(
        df_filtered['sl thực tế'].astype(str).str.replace(',', '.'),
        errors='coerce'
    ).fillna(1)
    
    # Calculate totals:
    # - Nhân với sl thực tế: gia công, gá lắp
    # - SUM trực tiếp: chuẩn bị, chạy thử, dừng, dừng khác, sửa
    # - QUAN TRỌNG: Loại bỏ thời gian dừng/dừng khác >= 420, 660, 630 (thời gian ca)
    
    time_tgcb = df_filtered['tgcb'].sum()
    time_chay_thu = df_filtered['chạy thử'].sum()
    time_ga_lap = (df_filtered['gá lắp'] * df_filtered['sl_thuc_te']).sum()
    time_gia_cong = (df_filtered['gia công'] * df_filtered['sl_thuc_te']).sum()
    
    # Filter out shift times (420, 660, 630) from stop times
    SHIFT_TIMES = [420, 630, 660]
    df_filtered_dung = df_filtered[~df_filtered['dừng'].isin(SHIFT_TIMES)].copy()
    time_dung = df_filtered_dung['dừng'].sum()
    
    if 'dừng khác' in df_filtered.columns:
        df_filtered_dung_khac = df_filtered[~df_filtered['dừng khác'].isin(SHIFT_TIMES)].copy()
        time_dung_khac = df_filtered_dung_khac['dừng khác'].sum()
    else:
        time_dung_khac = 0
    
    time_sua = df_filtered['sửa'].sum()
    
    # Total time = sum of ALL components
    total_time = time_tgcb + time_chay_thu + time_ga_lap + time_gia_cong + time_dung + time_dung_khac + time_sua
    
    if total_time == 0:
        return None
    
    # Calculate percentages
    result = {
        'total_time': total_time,
        'time_tgcb': time_tgcb,
        'time_chay_thu': time_chay_thu,
        'time_ga_lap': time_ga_lap,
        'time_gia_cong': time_gia_cong,
        'time_dung': time_dung,
        'time_dung_khac': time_dung_khac,
        'time_sua': time_sua,
        
        # Percentages: (component / total) * 100
        'pct_tgcb': (time_tgcb / total_time * 100),
        'pct_chay_thu': (time_chay_thu / total_time * 100),
        'pct_ga_lap': (time_ga_lap / total_time * 100),
        'pct_gia_cong': (time_gia_cong / total_time * 100),
        'pct_dung': (time_dung / total_time * 100),
        'pct_dung_khac': (time_dung_khac / total_time * 100),
        'pct_sua': (time_sua / total_time * 100),
    }
    
    return result

def create_stacked_bar_chart(data_dict, title):
    """
    Tạo biểu đồ xếp chồng theo thứ tự Excel
    Thứ tự từ dưới lên: Gia công → Gá lắp → Chạy thử → Chuẩn bị → Dừng → Sửa hàng → Dừng khác
    """
    categories = list(data_dict.keys())
    
    # Prepare data - matching Excel order
    gia_cong = [data_dict[cat]['pct_gia_cong'] for cat in categories]
    ga_lap = [data_dict[cat]['pct_ga_lap'] for cat in categories]
    chay_thu = [data_dict[cat]['pct_chay_thu'] for cat in categories]
    tgcb = [data_dict[cat]['pct_tgcb'] for cat in categories]
    dung = [data_dict[cat]['pct_dung'] for cat in categories]
    sua = [data_dict[cat]['pct_sua'] for cat in categories]
    dung_khac = [data_dict[cat]['pct_dung_khac'] for cat in categories]
    
    fig = go.Figure()
    
    # 1. Gia công (green - bottom)
    fig.add_trace(go.Bar(
        name='Tỷ lệ thời gian gia công',
        x=categories,
        y=gia_cong,
        marker_color='#92D050',
        text=[f'{data_dict[cat]["time_gia_cong"]:.0f}<br>{v:.0f}%' for cat, v in zip(categories, gia_cong)],
        textposition='inside',
        textfont=dict(size=16, color='black'),  # Increased font size
        hovertemplate='<b>Gia công</b><br>%{y:.1f}%<br>%{customdata[0]:.0f} phút<extra></extra>',
        customdata=[[data_dict[cat]["time_gia_cong"]] for cat in categories]
    ))
    
    # 2. Gá lắp (gray)
    fig.add_trace(go.Bar(
        name='Tỷ lệ thời gian gá lắp',
        x=categories,
        y=ga_lap,
        marker_color='#A6A6A6',
        text=[f'{data_dict[cat]["time_ga_lap"]:.0f}\u003cbr\u003e{v:.0f}%' if v > 3 else '' for cat, v in zip(categories, ga_lap)],
        textposition='inside',
        textfont=dict(size=16, color='black'),  # Increased font size
        hovertemplate='<b>Gá lắp</b><br>%{y:.1f}%<br>%{customdata[0]:.0f} phút<extra></extra>',
        customdata=[[data_dict[cat]["time_ga_lap"]] for cat in categories]
    ))
    
    # 3. Chạy thử (light blue - matching Excel)
    fig.add_trace(go.Bar(
        name='Tỷ lệ thời gian chạy thử',
        x=categories,
        y=chay_thu,
        marker_color='#9DC3E6',  # Light blue like Excel
        text=[f'{data_dict[cat]["time_chay_thu"]:.0f}\u003cbr\u003e{v:.0f}%' if v > 3 else '' for cat, v in zip(categories, chay_thu)],
        textposition='inside',
        textfont=dict(size=14, color='black'),  # Increased font size
        hovertemplate='\u003cb\u003eChạy thử\u003c/b\u003e\u003cbr\u003e%{y:.1f}%\u003cbr\u003e%{customdata[0]:.0f} phút\u003cextra\u003e\u003c/extra\u003e',
        customdata=[[data_dict[cat]["time_chay_thu"]] for cat in categories]
    ))
    
    # 4. Chuẩn bị (yellow - matching Excel)
    fig.add_trace(go.Bar(
        name='Tỷ lệ thời gian chuẩn bị',
        x=categories,
        y=tgcb,
        marker_color='#FFD966',  # Yellow like Excel
        text=[f'{data_dict[cat]["time_tgcb"]:.0f}\u003cbr\u003e{v:.0f}%' if v > 3 else '' for cat, v in zip(categories, tgcb)],
        textposition='inside',
        textfont=dict(size=14, color='black'),  # Increased font size
        hovertemplate='\u003cb\u003eChuẩn bị\u003c/b\u003e\u003cbr\u003e%{y:.1f}%\u003cbr\u003e%{customdata[0]:.0f} phút\u003cextra\u003e\u003c/extra\u003e',
        customdata=[[data_dict[cat]["time_tgcb"]] for cat in categories]
    ))
    
    # 5. Dừng (red - matching Excel)
    fig.add_trace(go.Bar(
        name='Tỷ lệ thời gian dừng',
        x=categories,
        y=dung,
        marker_color='#FF0000',  # Red color like Excel
        text=[f'{data_dict[cat]["time_dung"]:.0f}\u003cbr\u003e{v:.0f}%' if v > 3 else '' for cat, v in zip(categories, dung)],
        textposition='inside',
        textfont=dict(size=16, color='white'),  # Increased font size
        hovertemplate='\u003cb\u003eDừng\u003c/b\u003e\u003cbr\u003e%{y:.1f}%\u003cbr\u003e%{customdata[0]:.0f} phút\u003cextra\u003e\u003c/extra\u003e',
        customdata=[[data_dict[cat]["time_dung"]] for cat in categories]
    ))
    
    # 6. Sửa hàng (orange - matching Excel)
    fig.add_trace(go.Bar(
        name='Tỷ lệ thời gian sửa hàng',
        x=categories,
        y=sua,
        marker_color='#FFC000',  # Orange like Excel
        text=[f'{data_dict[cat]["time_sua"]:.0f}\u003cbr\u003e{v:.0f}%' if v > 2 else '' for cat, v in zip(categories, sua)],
        textposition='inside',
        textfont=dict(size=14, color='black'),  # Increased font size
        hovertemplate='\u003cb\u003eSửa hàng\u003c/b\u003e\u003cbr\u003e%{y:.1f}%\u003cbr\u003e%{customdata[0]:.0f} phút\u003cextra\u003e\u003c/extra\u003e',
        customdata=[[data_dict[cat]["time_sua"]] for cat in categories]
    ))
    
    # 7. Dừng khác (dark red - matching Excel)
    fig.add_trace(go.Bar(
        name='Tỷ lệ thời gian dừng khác',
        x=categories,
        y=dung_khac,
        marker_color='#C00000',  # Dark red color like Excel
        text=[f'{data_dict[cat]["time_dung_khac"]:.0f}\u003cbr\u003e{v:.0f}%' if v > 2 else '' for cat, v in zip(categories, dung_khac)],
        textposition='inside',
        textfont=dict(size=14, color='white'),  # Increased font size
        hovertemplate='\u003cb\u003eDừng khác\u003c/b\u003e\u003cbr\u003e%{y:.1f}%\u003cbr\u003e%{customdata[0]:.0f} phút\u003cextra\u003e\u003c/extra\u003e',
        customdata=[[data_dict[cat]["time_dung_khac"]] for cat in categories]
    ))
    
    fig.update_layout(
        title=dict(text=title, font=dict(size=20)),
        barmode='stack',
        barnorm='percent',
        yaxis=dict(
            title=dict(text='Tỷ lệ %', font=dict(size=16)),  # Correct syntax
            range=[0, 100],
            tickfont=dict(size=14)
        ),
        xaxis=dict(title='', tickfont=dict(size=14)),
        height=600,
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.2,
            xanchor="center",
            x=0.5,
            font=dict(size=14)
        ),
        hovermode='closest'
    )
    
    return fig

def calculate_machine_counts(df, machine_type, dept_name):
    """
    Tính số máy chạy (có thời gian gia công > 0)
    """
    lathe_machines = CONFIG['lathe_machines']
    
    # Filter by machine type
    if machine_type == 'lathe':
        df_filtered = df[df['số máy'].isin(lathe_machines)].copy()
    elif machine_type == 'milling':
        df_filtered = df[~df['số máy'].isin(lathe_machines)].copy()
    else:
        df_filtered = df.copy()
    
    # Count unique machines with processing time > 0
    machines_with_production = set()
    for machine in df_filtered['số máy'].unique():
        df_machine = df_filtered[df_filtered['số máy'] == machine].copy()
        time_gia_cong = (df_machine['gia công'] * pd.to_numeric(
            df_machine['sl thực tế'].astype(str).str.replace(',', '.'),
            errors='coerce'
        ).fillna(1)).sum()
        
        if time_gia_cong > 0:
            machines_with_production.add(machine)
    
    return len(machines_with_production)

def create_machine_time_count_chart(data_dict, count_dict):
    """
    Tạo biểu đồ bar với annotations (bong bóng ghi chú) cho số máy
    data_dict: {'Tiện SX1': time, 'Tiện SX2': time, 'Phay SX1': time, 'Phay SX2': time}
    count_dict: {'Tiện SX1': count, 'Tiện SX2': count, 'Phay SX1': count, 'Phay SX2': count}
    """
    categories = list(data_dict.keys())
    times = list(data_dict.values())
    counts = list(count_dict.values())
    
    # Calculate percentages
    total_time = sum(times)
    percentages = [(t / total_time * 100) if total_time > 0 else 0 for t in times]
    
    # Colors matching Excel
    colors = {
        'Tiện SX1': '#C5E0B4',
        'Tiện SX2': '#A9D08E',
        'Phay SX1': '#00B0F0',
        'Phay SX2': '#0070C0'
    }
    
    fig = go.Figure()
    
    # Add bars for processing time
    for i, cat in enumerate(categories):
        fig.add_trace(go.Bar(
            name=f'Thời gian gia công máy {cat.lower()}',
            x=[cat],
            y=[percentages[i]],
            marker_color=colors.get(cat, '#999999'),
            text=f'{times[i]:.0f}<br>{percentages[i]:.0f}%',
            textposition='inside',
            textfont=dict(size=14, color='black'),
            showlegend=True
        ))
    
    # Add annotations (callouts) for machine counts above bars
    annotations = []
    for i, cat in enumerate(categories):
        annotations.append(dict(
            x=cat,
            y=percentages[i] + 5,  # Position above bar
            text=f'Số máy {cat.lower()} chạy {cat.split()[1]}<br>{counts[i]}',
            showarrow=True,
            arrowhead=2,
            arrowsize=1,
            arrowwidth=1,
            arrowcolor='#666666',
            ax=0,
            ay=-40,
            font=dict(size=10, color='black'),
            bgcolor='white',
            bordercolor='#666666',
            borderwidth=1,
            borderpad=4
        ))
    
    # Update layout
    fig.update_layout(
        title=dict(text='Thời gian + Số máy chạy phay + tiện 2 ca SX', font=dict(size=20)),
        xaxis=dict(title='', tickfont=dict(size=14)),
        yaxis=dict(
            title=dict(text='Tỷ lệ %', font=dict(size=14)),
            range=[0, 80],  # Adjusted for annotations
            tickfont=dict(size=12)
        ),
        height=500,
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.2,
            xanchor="center",
            x=0.5,
            font=dict(size=11)
        ),
        annotations=annotations,
        hovermode='x unified'
    )
    
    return fig


def main():
    st.title("📊 BIỂU ĐỒ TỔNG CÔNG SUẤT MÁY LẺ")
    
    # Sidebar
    with st.sidebar:
        st.header("⚙️ Cài đặt")
        
        if st.button("🔄 Làm mới dữ liệu"):
            st.cache_data.clear()
            st.rerun()
        
        st.markdown("---")
        st.info(f"📅 {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    
    # Load data
    with st.spinner("Đang tải dữ liệu PHTCV..."):
        df_phtcv = read_phtcv_data()
    
    if df_phtcv is None or df_phtcv.empty:
        st.error("❌ Không thể tải dữ liệu PHTCV")
        return
    
    # Filters - Month and Date filter with Excel export
    col_filter1, col_filter2, col_export = st.columns([1, 1, 1])
    
    with col_filter1:
        # Get available months
        if 'date_parsed' in df_phtcv.columns:
            df_phtcv['year_month'] = df_phtcv['date_parsed'].dt.to_period('M')
            available_months = df_phtcv['year_month'].dropna().unique()
            available_months = sorted(available_months, reverse=True)
            
            if len(available_months) > 0:
                month_options = ['Tất cả'] + [str(m) for m in available_months]
                selected_month = st.selectbox(
                    "Chọn tháng:",
                    options=month_options,
                    index=0
                )
            else:
                selected_month = 'Tất cả'
        else:
            selected_month = 'Tất cả'
    
    with col_filter2:
        # Get available dates (filtered by month if selected)
        if 'date_parsed' in df_phtcv.columns:
            if selected_month != 'Tất cả':
                df_month_filtered = df_phtcv[df_phtcv['year_month'] == pd.Period(selected_month)].copy()
                available_dates = df_month_filtered['date_parsed'].dropna().dt.date.unique()
            else:
                available_dates = df_phtcv['date_parsed'].dropna().dt.date.unique()
            
            available_dates = sorted(available_dates, reverse=True)
            
            if len(available_dates) > 0:
                selected_date = st.selectbox(
                    "Chọn ngày:",
                    options=['Tất cả'] + [d.strftime('%d/%m/%Y') for d in available_dates],
                    index=0
                )
            else:
                selected_date = 'Tất cả'
                st.warning("Không tìm thấy ngày trong dữ liệu")
        else:
            selected_date = 'Tất cả'
    
    # Filter by month and date
    df_filtered = df_phtcv.copy()
    
    if selected_month != 'Tất cả' and 'year_month' in df_filtered.columns:
        df_filtered = df_filtered[df_filtered['year_month'] == pd.Period(selected_month)].copy()
        st.info(f"📅 Hiển thị dữ liệu tháng: {selected_month}")
    
    if selected_date != 'Tất cả' and 'date_parsed' in df_filtered.columns:
        filter_date = pd.to_datetime(selected_date, format='%d/%m/%Y').date()
        df_filtered = df_filtered[df_filtered['date_parsed'].dt.date == filter_date].copy()
        st.info(f"📅 Hiển thị dữ liệu ngày: {selected_date}")
    
    # Excel Export Button
    with col_export:
        st.write("")  # Spacing
        st.write("")  # Spacing
        if not df_filtered.empty:
            # Prepare export data
            export_df = df_filtered.copy()
            
            # Remove helper columns
            cols_to_remove = ['date_parsed', 'year_month', 'sl_thuc_te']
            export_df = export_df.drop(columns=[col for col in cols_to_remove if col in export_df.columns])
            
            # Convert to Excel
            from io import BytesIO
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                export_df.to_excel(writer, index=False, sheet_name='Công suất')
            
            excel_data = output.getvalue()
            
            # Determine filename
            if selected_month != 'Tất cả':
                filename = f"cong_suat_{selected_month}.xlsx"
            elif selected_date != 'Tất cả':
                filename = f"cong_suat_{selected_date.replace('/', '-')}.xlsx"
            else:
                filename = "cong_suat_tat_ca.xlsx"
            
            st.download_button(
                label="📥 Xuất Excel",
                data=excel_data,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    
    # Calculate capacity for both departments first
    departments = ['Sản xuất 1', 'Sản xuất 2']
    dept_capacities = {}
    
    for dept in departments:
        df_dept = df_filtered[df_filtered['bộ phận'] == dept].copy()
        if not df_dept.empty:
            cap_total = calculate_capacity_by_type(df_dept, 'all')
            if cap_total:
                dept_capacities[dept] = cap_total
    
    # ========== BIỂU ĐỒ TỔNG SX1 VÀ SX2 ==========
    if len(dept_capacities) >= 2:
        st.markdown("---")
        st.header("📊 SO SÁNH CÔNG SUẤT TỔNG SX1 VÀ SX2")
        
        # Display summary metrics
        col1, col2 = st.columns(2)
        
        with col1:
            sx1_cap = dept_capacities.get('Sản xuất 1', {})
            if sx1_cap:
                st.metric("🏭 Sản xuất 1 - Tổng thời gian", f"{sx1_cap['total_time']:.0f} phút")
                st.metric("Tỷ lệ gia công", f"{sx1_cap['pct_gia_cong']:.1f}%")
        
        with col2:
            sx2_cap = dept_capacities.get('Sản xuất 2', {})
            if sx2_cap:
                st.metric("🏭 Sản xuất 2 - Tổng thời gian", f"{sx2_cap['total_time']:.0f} phút")
                st.metric("Tỷ lệ gia công", f"{sx2_cap['pct_gia_cong']:.1f}%")
        
        # Create combined chart
        combined_data = {
            'SẢN XUẤT 1': dept_capacities['Sản xuất 1'],
            'SẢN XUẤT 2': dept_capacities['Sản xuất 2']
        }
        
        fig_combined = create_stacked_bar_chart(combined_data, "BIỂU ĐỒ SO SÁNH CÔNG SUẤT TỔNG - SX1 VÀ SX2")
        st.plotly_chart(fig_combined, use_container_width=True)
        
        # Show comparison table
        with st.expander("📋 Xem bảng so sánh chi tiết"):
            comparison_df = pd.DataFrame({
                'Phân xưởng': ['Sản xuất 1', 'Sản xuất 2'],
                'Tổng phút': [sx1_cap['total_time'], sx2_cap['total_time']],
                'Gia công (%)': [f"{sx1_cap['pct_gia_cong']:.0f}%", f"{sx2_cap['pct_gia_cong']:.0f}%"],
                'Gá lắp (%)': [f"{sx1_cap['pct_ga_lap']:.0f}%", f"{sx2_cap['pct_ga_lap']:.0f}%"],
                'Chạy thử (%)': [f"{sx1_cap['pct_chay_thu']:.0f}%", f"{sx2_cap['pct_chay_thu']:.0f}%"],
                'Dừng (%)': [f"{sx1_cap['pct_dung']:.0f}%", f"{sx2_cap['pct_dung']:.0f}%"],
                'Dừng khác (%)': [f"{sx1_cap['pct_dung_khac']:.0f}%", f"{sx2_cap['pct_dung_khac']:.0f}%"],
                'Sửa (%)': [f"{sx1_cap['pct_sua']:.0f}%", f"{sx2_cap['pct_sua']:.0f}%"],
            })
            st.dataframe(comparison_df, use_container_width=True)
        
        # ========== METRIC BOXES: THỜI GIAN + SỐ MÁY CHẠY ==========
        st.markdown("---")
        st.header("📊 THỜI GIAN + SỐ MÁY CHẠY PHAY + TIỆN 2 CA SX")
        
        # Calculate data for all 4 categories
        time_data = {}
        count_data = {}
        
        for dept in departments:
            df_dept_chart = df_filtered[df_filtered['bộ phận'] == dept].copy()
            if not df_dept_chart.empty:
                # Calculate lathe data
                lathe_cap = calculate_capacity_by_type(df_dept_chart, 'lathe')
                if lathe_cap:
                    dept_short = 'SX1' if dept == 'Sản xuất 1' else 'SX2'
                    time_data[f'Tiện {dept_short}'] = lathe_cap['time_gia_cong']
                    count_data[f'Tiện {dept_short}'] = calculate_machine_counts(df_dept_chart, 'lathe', dept)
                
                # Calculate milling data
                milling_cap = calculate_capacity_by_type(df_dept_chart, 'milling')
                if milling_cap:
                    dept_short = 'SX1' if dept == 'Sản xuất 1' else 'SX2'
                    time_data[f'Phay {dept_short}'] = milling_cap['time_gia_cong']
                    count_data[f'Phay {dept_short}'] = calculate_machine_counts(df_dept_chart, 'milling', dept)
        
        # Display as metric boxes - showing separate lathe and milling counts
        if time_data and count_data:
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                # Box 1: Số máy tiện SX1
                count_sx1_tien = count_data.get('Tiện SX1', 0)
                st.metric("Số máy tiện chạy SX1", f"{count_sx1_tien}")
            
            with col2:
                # Box 2: Số máy tiện SX2
                count_sx2_tien = count_data.get('Tiện SX2', 0)
                st.metric("Số máy tiện chạy SX2", f"{count_sx2_tien}")
            
            with col3:
                # Box 3: Số máy phay SX1
                count_sx1_phay = count_data.get('Phay SX1', 0)
                st.metric("Số máy phay chạy SX1", f"{count_sx1_phay}")
            
            with col4:
                # Box 4: Số máy phay SX2
                count_sx2_phay = count_data.get('Phay SX2', 0)
                st.metric("Số máy phay chạy SX2", f"{count_sx2_phay}")

            

        else:
            st.warning("Không đủ dữ liệu để hiển thị")
    
    
    # ========== CHI TIẾT TỪNG PHÂN XƯỞNG ==========
    st.markdown("---")
    st.header("📋 CHI TIẾT TỪNG PHÂN XƯỞNG")
    
    for dept in departments:
        st.markdown("---")
        st.subheader(f"Công Suất {dept}")
        
        # Filter by department
        df_dept = df_filtered[df_filtered['bộ phận'] == dept].copy()
        
        if df_dept.empty:
            st.warning(f"Không có dữ liệu cho {dept}")
            continue
        
        # Calculate capacity for 3 categories
        cap_lathe = calculate_capacity_by_type(df_dept, 'lathe')
        cap_milling = calculate_capacity_by_type(df_dept, 'milling')
        cap_total = calculate_capacity_by_type(df_dept, 'all')
        
        if not cap_lathe or not cap_milling or not cap_total:
            st.error(f"Không thể tính toán công suất cho {dept}")
            continue
        
        # Display metrics
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Tổng CS Máy Tiện", f"{cap_lathe['total_time']:.0f} phút")
            st.metric("Tỷ lệ gia công", f"{cap_lathe['pct_gia_cong']:.0f}%")
        
        with col2:
            st.metric("Tổng CS Máy Phay", f"{cap_milling['total_time']:.0f} phút")
            st.metric("Tỷ lệ gia công", f"{cap_milling['pct_gia_cong']:.0f}%")
        
        with col3:
            st.metric("Tổng Cộng", f"{cap_total['total_time']:.0f} phút")
            st.metric("Tỷ lệ gia công", f"{cap_total['pct_gia_cong']:.0f}%")
        
        # Create chart
        data_dict = {
            'TỔNG CS MÁY TIỆN': cap_lathe,
            'TỔNG CS MÁY PHAY': cap_milling,
            'TỔNG CỘNG': cap_total
        }
        
        fig = create_stacked_bar_chart(data_dict, f"BIỂU ĐỒ TỔNG CÔNG SUẤT MÁY - {dept}")
        st.plotly_chart(fig, use_container_width=True)
        
        # Show detailed analysis in tabs
        st.markdown("### 📋 Phân tích chi tiết")
        
        tab1, tab2, tab3, tab4 = st.tabs(["⚠️ Máy dừng > 10%", "🔧 Máy gá lắp > 10%", "⏱️ Máy chuẩn bị > 10%", "🛑 Máy dừng 100%"])
        
        # Calculate machine-level statistics
        df_dept['sl_thuc_te'] = pd.to_numeric(
            df_dept['sl thực tế'].astype(str).str.replace(',', '.'),
            errors='coerce'
        ).fillna(1)
        
        # Group by machine
        machine_stats = []
        for machine in df_dept['số máy'].unique():
            df_machine = df_dept[df_dept['số máy'] == machine].copy()
            
            # Calculate times
            time_tgcb = df_machine['tgcb'].sum()
            time_chay_thu = df_machine['chạy thử'].sum()
            time_ga_lap = (df_machine['gá lắp'] * df_machine['sl_thuc_te']).sum()
            time_gia_cong = (df_machine['gia công'] * df_machine['sl_thuc_te']).sum()
            
            SHIFT_TIMES = [420, 630, 660]
            df_dung_filtered = df_machine[~df_machine['dừng'].isin(SHIFT_TIMES)].copy()
            time_dung = df_dung_filtered['dừng'].sum()
            
            if 'dừng khác' in df_machine.columns:
                df_dung_khac_filtered = df_machine[~df_machine['dừng khác'].isin(SHIFT_TIMES)].copy()
                time_dung_khac = df_dung_khac_filtered['dừng khác'].sum()
            else:
                time_dung_khac = 0
            
            time_sua = df_machine['sửa'].sum()
            total_time = time_tgcb + time_chay_thu + time_ga_lap + time_gia_cong + time_dung + time_dung_khac + time_sua
            
            if total_time > 0:
                # Get explanations
                explanations = df_machine['giải trình'].dropna().tolist() if 'giải trình' in df_machine.columns else []
                explanation_text = ', '.join([str(e) for e in explanations if str(e).strip() != ''])
                
                machine_stats.append({
                    'số máy': machine,
                    'machine_num': int(machine) if machine.isdigit() else 9999,  # For sorting
                    'total_time': total_time,
                    'time_dung': time_dung,
                    'time_dung_khac': time_dung_khac,
                    'time_ga_lap': time_ga_lap,
                    'time_tgcb': time_tgcb,
                    'pct_dung': (time_dung / total_time * 100),
                    'pct_dung_khac': (time_dung_khac / total_time * 100),
                    'pct_ga_lap': (time_ga_lap / total_time * 100),
                    'pct_tgcb': (time_tgcb / total_time * 100),
                    'pct_total_stop': ((time_dung + time_dung_khac) / total_time * 100),
                    'explanation': explanation_text
                })
        
        df_machine_stats = pd.DataFrame(machine_stats)
        
        # TAB 1: Máy dừng > 10%
        with tab1:
            if not df_machine_stats.empty:
                df_stop = df_machine_stats[df_machine_stats['pct_total_stop'] > 10].copy()
                if not df_stop.empty:
                    df_stop = df_stop.sort_values('machine_num')  # Sort by numeric value
                    df_stop_display = df_stop[['số máy', 'pct_total_stop', 'explanation']].copy()
                    df_stop_display.columns = ['Số máy', 'Tỷ lệ % dừng', 'Lý do']
                    df_stop_display['Tỷ lệ % dừng'] = df_stop_display['Tỷ lệ % dừng'].round(1)
                    st.dataframe(df_stop_display, use_container_width=True, hide_index=True)
                else:
                    st.success("✅ Không có máy nào có tỷ lệ dừng > 10%")
            else:
                st.warning("Không có dữ liệu")
        
        # TAB 2: Máy gá lắp > 10%
        with tab2:
            if not df_machine_stats.empty:
                df_ga_lap = df_machine_stats[df_machine_stats['pct_ga_lap'] > 10].copy()
                if not df_ga_lap.empty:
                    df_ga_lap = df_ga_lap.sort_values('machine_num')  # Sort by numeric value
                    df_ga_lap_display = df_ga_lap[['số máy', 'explanation']].copy()
                    df_ga_lap_display.columns = ['Số máy', 'Lý do']
                    st.dataframe(df_ga_lap_display, use_container_width=True, hide_index=True)
                else:
                    st.success("✅ Không có máy nào có tỷ lệ gá lắp > 10%")
            else:
                st.warning("Không có dữ liệu")
        
        # TAB 3: Máy chuẩn bị > 10%
        with tab3:
            if not df_machine_stats.empty:
                df_tgcb = df_machine_stats[df_machine_stats['pct_tgcb'] > 10].copy()
                if not df_tgcb.empty:
                    df_tgcb_display = df_tgcb[['số máy']].copy()
                    df_tgcb_display.columns = ['Số máy']
                    st.dataframe(df_tgcb_display, use_container_width=True, hide_index=True)
                else:
                    st.success("✅ Không có máy nào có tỷ lệ chuẩn bị > 10%")
            else:
                st.warning("Không có dữ liệu")

        # TAB 3: Máy chuẩn bị > 10%
        with tab3:
            if not df_machine_stats.empty:
                df_tgcb = df_machine_stats[df_machine_stats['pct_tgcb'] > 10].copy()
                if not df_tgcb.empty:
                    df_tgcb = df_tgcb.sort_values('machine_num')  # Sort by numeric value
                    df_tgcb_display = df_tgcb[['số máy']].copy()
                    df_tgcb_display.columns = ['Số máy']
                    st.dataframe(df_tgcb_display, use_container_width=True, hide_index=True)
                else:
                    st.success("✅ Không có máy nào có tỷ lệ chuẩn bị > 10%")
            else:
                st.warning("Không có dữ liệu")
        
        # TAB 4: Máy dừng 100%
        with tab4:
            # Get full machine list from Google Sheets
            all_machines = read_machine_list()
            
            # Get machines in current department data
            machines_in_data = df_dept['số máy'].unique().tolist()
            
            # CONDITION 1: Machines NOT in data
            machines_not_in_data = [m for m in all_machines if m not in machines_in_data]
            
            # CONDITION 2 AND 3: Machines with stop time >= shift times AND all production columns empty
            stopped_machines_with_data = []
            SHIFT_TIMES = [420, 630, 660]
            
            for machine in machines_in_data:
                df_machine = df_dept[df_dept['số máy'] == machine].copy()
                
                # Check if machine has stop time >= any shift time
                max_dung = df_machine['dừng'].max()
                max_dung_khac = df_machine['dừng khác'].max() if 'dừng khác' in df_machine.columns else 0
                
                has_shift_stop = (max_dung >= 420) or (max_dung_khac >= 420)
                
                # Check if all production columns are empty/zero
                time_tgcb = df_machine['tgcb'].sum()
                time_chay_thu = df_machine['chạy thử'].sum()
                time_ga_lap = df_machine['gá lắp'].sum()
                time_gia_cong = df_machine['gia công'].sum()
                
                has_no_production = (time_tgcb == 0 and time_chay_thu == 0 and 
                                    time_ga_lap == 0 and time_gia_cong == 0)
                
                # Condition 2 AND 3
                if has_shift_stop and has_no_production:
                    stopped_machines_with_data.append(machine)
            
            # FINAL: Condition 1 OR (Condition 2 AND 3)
            # Sort numerically by converting to int
            all_stopped_machines = sorted(
                set(machines_not_in_data + stopped_machines_with_data),
                key=lambda x: int(x) if x.isdigit() else float('inf')
            )
            total_stopped = len(all_stopped_machines)
            
            # Display results
            if total_stopped > 0:
                # Create list of machines with categories
                machine_data = []
                
                # Add Total row first
                machine_data.append({'Số máy': 'Total', 'Trạng thái': f'{total_stopped} máy'})
                
                # Add individual machines
                for machine in all_stopped_machines:
                    if machine in machines_not_in_data:
                        category = "Không có dữ liệu"
                    else:
                        category = "Dừng toàn bộ ca"
                    machine_data.append({'Số máy': machine, 'Trạng thái': category})
                
                df_stopped = pd.DataFrame(machine_data)
                
                # Display with scrolling
                st.dataframe(
                    df_stopped,
                    use_container_width=True,
                    hide_index=True,
                    height=400  # Fixed height to enable scrolling
                )
            else:
                st.success("✅ Không có máy nào dừng 100%")


if __name__ == "__main__":
    main()
