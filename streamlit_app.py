import io
import openpyxl
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# 1. ตั้งค่า Page Config
st.set_page_config(
    page_title="Datapaq NB3 BTM",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. บังคับ Dark Mode CSS + จัดการตารางให้อยู่ตรงกลางอย่างสมบูรณ์
st.markdown("""
    <style>
        /* ซ่อนแถบขาว Header ด้านบน */
        header[data-testid="stHeader"] {
            background-color: transparent !important;
            display: none !important;
        }
        [data-testid="stToolbar"] {
            display: none !important;
        }
        
        /* ตั้งค่าพื้นหลัง Dark Mode */
        html, body, .stApp, [data-testid="stAppViewContainer"] {
            background-color: #0e1117 !important;
            color: #ffffff !important;
        }
        [data-testid="stSidebar"] {
            background-color: #161b22 !important;
        }
        .stMarkdown, h1, h2, h3, p, span, label {
            color: #ffffff !important;
        }

        /* ปุ่มเคลียร์ข้อมูลใน Sidebar */
        [data-testid="stSidebar"] div.stButton > button {
            background-color: #21262d !important;
            color: #ffffff !important;
            border: 1px solid #F0B90B !important;
            font-weight: bold !important;
            width: 100% !important;
            padding: 8px 16px !important;
        }
        [data-testid="stSidebar"] div.stButton > button:hover {
            background-color: #F0B90B !important;
            color: #000000 !important;
        }

        /* กล่อง File Uploader */
        [data-testid="stFileUploader"] {
            background-color: #161b22 !important;
            border: 1.5px solid #F0B90B !important;
            border-radius: 8px !important;
            padding: 10px !important;
        }
        [data-testid="stFileUploader"] section {
            background-color: #1c2128 !important;
            border: 1px dashed #F0B90B !important;
            border-radius: 6px !important;
        }
        [data-testid="stFileUploader"] section div, 
        [data-testid="stFileUploader"] section span,
        [data-testid="stFileUploader"] section small {
            color: #e6edf3 !important;
        }

        /* การ์ดไฟล์ที่อัปโหลดแล้ว */
        [data-testid="stFileUploaderFileData"],
        [data-testid="stFileUploaderFileData"] > div,
        [data-testid="stFileUploaderFile"] {
            background-color: #21262d !important;
            border: 1px solid #F0B90B !important;
            border-radius: 6px !important;
        }
        [data-testid="stFileUploaderFileData"] *,
        [data-testid="stFileUploaderFile"] * {
            color: #ffffff !important;
            font-weight: bold !important;
        }

        /* สไตล์กล่องแสดง Header Metadata แบบ Raw Header */
        .raw-header-box {
            background-color: #161b22;
            border: 1px solid #30363d;
            border-left: 4px solid #F0B90B;
            border-radius: 6px;
            padding: 12px 18px;
            font-family: 'Courier New', Courier, monospace;
            font-size: 14px;
            color: #e6edf3;
            margin-bottom: 15px;
            line-height: 1.6;
        }
        .raw-header-key {
            color: #58a6ff;
            font-weight: bold;
        }
        .raw-header-val {
            color: #D29922;
            font-weight: bold;
        }

        /* ปรับแถบ Expander */
        [data-testid="stExpander"] {
            background-color: #161b22 !important;
            border: 1px solid #30363d !important;
            border-radius: 8px !important;
        }
        [data-testid="stExpander"] details summary {
            background-color: #21262d !important;
            color: #ffffff !important;
            border-radius: 8px !important;
        }
        [data-testid="stExpander"] details summary * {
            color: #ffffff !important;
        }

        /* ปรับแต่งตาราง Dataframe & บังคับจัดข้อความ/ตัวเลขทุกช่องให้อยู่ตรงกลางทั้งหมด */
        [data-testid="stDataFrame"], [data-testid="stTable"] {
            background-color: #161b22 !important;
            border: 1px solid #30363d !important;
            border-radius: 8px !important;
        }
        div[data-testid="stDataFrame"] div[role="grid"] {
            background-color: #161b22 !important;
            color: #ffffff !important;
        }
        div[data-testid="stDataFrame"] div[role="columnheader"] {
            background-color: #21262d !important;
            color: #ffffff !important;
            justify-content: center !important;
            text-align: center !important;
        }
        div[data-testid="stDataFrame"] div[role="columnheader"] * {
            justify-content: center !important;
            text-align: center !important;
        }
        div[data-testid="stDataFrame"] div[role="gridcell"] {
            justify-content: center !important;
            text-align: center !important;
            display: flex !important;
            align-items: center !important;
        }
        div[data-testid="stDataFrame"] div[role="gridcell"] * {
            text-align: center !important;
        }
        
        [data-testid="stTable"] th, [data-testid="stTable"] td {
            text-align: center !important;
            vertical-align: middle !important;
        }

        /* ปรับแต่งปุ่มดาวน์โหลด Excel */
        div.stDownloadButton > button {
            background-color: #21262d !important;
            border: 1.5px solid #F0B90B !important;
            border-radius: 6px !important;
            padding: 8px 16px !important;
            transition: all 0.2s ease-in-out;
        }
        div.stDownloadButton > button, 
        div.stDownloadButton > button *,
        div.stDownloadButton > button p,
        div.stDownloadButton > button span {
            color: #ffffff !important;
            font-weight: bold !important;
            font-size: 15px !important;
        }
        div.stDownloadButton > button:hover {
            background-color: #F0B90B !important;
            border-color: #F0B90B !important;
        }
        div.stDownloadButton > button:hover,
        div.stDownloadButton > button:hover *,
        div.stDownloadButton > button:hover p,
        div.stDownloadButton > button:hover span {
            color: #000000 !important;
        }
    </style>
""", unsafe_allow_html=True)

# 3. แสดงชื่อโปรแกรมหลัก
st.title("🏭 Datapaq NB3 BTM")

# 4. ฟังก์ชันแปลงวินาทีเป็นรูปแบบ mm:ss หรือ hh:mm:ss
def format_seconds_to_time(total_seconds):
    if pd.isna(total_seconds) or total_seconds <= 0:
        return "00:00:00"
    
    total_sec = int(round(total_seconds + 1e-5))
    hours = total_sec // 3600
    minutes = (total_sec % 3600) // 60
    seconds = total_sec % 60
    
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

def time_str_to_seconds(t_str):
    parts = t_str.split(":")
    if len(parts) == 3:
        return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
    return 0

def hex_to_rgba(hex_str, opacity=0.25):
    hex_str = hex_str.lstrip('#')
    r = int(hex_str[0:2], 16)
    g = int(hex_str[2:4], 16)
    b = int(hex_str[4:6], 16)
    return f"rgba({r}, {g}, {b}, {opacity})"

# 📌 ฟังก์ชันปรับช่วงเวลาโซนสีให้แมตช์กับแกนเวลากราฟโดยไม่หลุดขอบ
def safe_zones_for_plotly(zones_data, df_chart):
    if df_chart.empty:
        return zones_data
    
    max_sec = df_chart["ElapsedSeconds"].iloc[-1]
    max_time_str = df_chart["Time (HH:MM:SS)"].iloc[-1]
    
    safe_zones = []
    for z in zones_data:
        z_start = time_str_to_seconds(z["Start Time"])
        z_end = time_str_to_seconds(z["End Time"])
        
        if z_start > max_sec:
            continue
            
        z_copy = dict(z)
        if z_end > max_sec:
            z_copy["End Time"] = max_time_str
            
        safe_zones.append(z_copy)
        
    return safe_zones

# 📌 ฟังก์ชันกำหนดตำแหน่ง Probe ตาม #title
def get_probe_location(p_num, title_meta):
    title_upper = str(title_meta).upper() if title_meta else ""
    
    # 1. เงื่อนไข M48
    if "M48" in title_upper:
        if p_num in [1, 2, 3, 4]:
            return "Bottom"
        else:
            return "Top"
            
    # 2. เงื่อนไข N173
    elif "N173" in title_upper:
        if "TYPE2" in title_upper or "ALT" in title_upper or "V2" in title_upper:
            if p_num in [1, 5]: return "L"
            elif p_num in [3, 7]: return "ML"
            elif p_num in [4, 6]: return "MR"
            else: return "R"
        else:
            if p_num in [1, 7]: return "L"
            elif p_num in [3, 5]: return "ML"
            elif p_num in [2, 8]: return "MR"
            else: return "R"
                
    # 3. กรณีอื่นๆ (ค่าเริ่มต้น BTM)
    else:
        if p_num in [1, 2]: return "L"
        elif p_num in [3, 4]: return "ML"
        elif p_num in [5, 6]: return "MR"
        else: return "R"

# 5. ฟังก์ชันอ่านไฟล์ CSV แบบตัดเวลาติดลบและ Auto-Trim ปรับเวลาเข้าเตา
def parse_single_file(uploaded_file):
    uploaded_file.seek(0)
    raw_bytes = uploaded_file.read()
    
    text_content = None
    for enc in ['utf-8', 'cp932', 'shift_jis', 'tis-620', 'latin1']:
        try:
            text_content = raw_bytes.decode(enc)
            break
        except Exception:
            continue
            
    if text_content is None:
        text_content = raw_bytes.decode('utf-8', errors='ignore')

    lines = text_content.splitlines()
    
    probe_labels = {}
    data_rows = []
    
    metadata = {
        "paqfile start date": "-",
        "paqfile start time": "-",
        "title": "-",
        "operator": "-",
        "product": "-"
    }

    for line in lines:
        line_str = line.strip()
        if not line_str:
            continue
        
        if line_str.startswith("#"):
            line_clean = line_str.lstrip("#").strip()
            if "=" in line_clean:
                key, val = [p.strip() for p in line_clean.split("=", 1)]
                val = val.rstrip(",")
                
                if key.lower() == "title":
                    metadata["title"] = val
                elif key.lower() == "paqfile start date":
                    metadata["paqfile start date"] = val
                elif key.lower() == "paqfile start time":
                    metadata["paqfile start time"] = val
                elif key.lower() == "operator":
                    metadata["operator"] = val
                elif key.lower() == "product":
                    metadata["product"] = val
                elif key.isdigit():
                    ch_num = int(key)
                    probe_labels[ch_num] = val
        else:
            parts = [p.strip() for p in line_str.split(",") if p.strip() != ""]
            if len(parts) >= 10:
                try:
                    time_str = parts[0].strip()
                    
                    if time_str.startswith("-"):
                        continue
                        
                    t_parts = time_str.split(":")
                    if len(t_parts) == 3:
                        elapsed_sec = int(t_parts[0]) * 3600 + int(t_parts[1]) * 60 + int(t_parts[2])
                    else:
                        elapsed_sec = len(data_rows)
                    
                    dist_val = float(parts[1])
                    probe_vals = [float(p) for p in parts[2:10]]
                    
                    data_rows.append({
                        "elapsed_sec": elapsed_sec,
                        "time_str": time_str,
                        "dist_val": dist_val,
                        "probes": probe_vals
                    })
                except (ValueError, IndexError):
                    continue

    if not data_rows:
        return pd.DataFrame(), metadata

    parsed_data = []
    for row in data_rows:
        row_dict = {
            "ElapsedSeconds": row["elapsed_sec"],
            "Time (HH:MM:SS)": row["time_str"],
            "Distance (m)": round(row["dist_val"], 2)
        }
        for i in range(1, 9):
            col_label = f"Probe #{i}"
            if i in probe_labels:
                lbl = probe_labels[i]
                col_label = f"Probe #{i}: {lbl[:15]}..." if len(lbl) > 15 else f"Probe #{i}: {lbl}"
            row_dict[col_label] = row["probes"][i-1]
            
        parsed_data.append(row_dict)

    df_res = pd.DataFrame(parsed_data)
    df_res = df_res.drop_duplicates(subset=["ElapsedSeconds"]).sort_values("ElapsedSeconds").reset_index(drop=True)
    
    # 📌 AUTO-TRIM LOGIC: ตัดเส้นแบนราบอุณหภูมิห้องออก (เริ่มนับ 0 เมื่ออุณหภูมิเริ่มไต่ระดับ)
    probe_cols = [c for c in df_res.columns if c.startswith("Probe #")]
    max_temps = df_res[probe_cols].max(axis=1)
    
    if not max_temps.empty and max_temps.iloc[0] < 45.0:
        valid_idx = max_temps[max_temps >= 45.0].index
        if len(valid_idx) > 0:
            s_idx = max(df_res.index[0], valid_idx[0] - 3)
            df_res = df_res.loc[s_idx:].copy()
            first_sec = df_res["ElapsedSeconds"].iloc[0]
            df_res["ElapsedSeconds"] = df_res["ElapsedSeconds"] - first_sec
            
            def sec_to_hhmmss(s):
                m, sec = divmod(s, 60)
                h, m = divmod(m, 60)
                return f"{h:02d}:{m:02d}:{sec:02d}"
                
            df_res["Time (HH:MM:SS)"] = df_res["ElapsedSeconds"].apply(sec_to_hhmmss)
    
    df_res = df_res.reset_index(drop=True)
    return df_res, metadata

def to_excel_bytes(dataframe, summary_dataframe=None, fig_plotly=None):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if summary_dataframe is not None and not summary_dataframe.empty:
            summary_dataframe.to_excel(writer, sheet_name='Parameter Summary')
            
        df_export = dataframe.copy()
        df_export.to_excel(writer, index=False, sheet_name='Raw Log Data')

    if fig_plotly is not None:
        try:
            img_bytes = fig_plotly.to_image(format="png", width=1200, height=550)
            img_buf = io.BytesIO(img_bytes)
            
            wb = openpyxl.load_workbook(output)
            ws = wb['Parameter Summary'] if 'Parameter Summary' in wb.sheetnames else wb.active
            
            img = openpyxl.drawing.image.Image(img_buf)
            img.anchor = 'A12'
            ws.add_image(img)
            
            output = io.BytesIO()
            wb.save(output)
        except Exception:
            pass

    output.seek(0)
    return output.getvalue()

# 6. เมนู Sidebar
st.sidebar.header("📁 เมนูอัปโหลดข้อมูล")

if st.sidebar.button("🧹 เคลียร์ข้อมูลไฟล์เก่าทั้งหมด"):
    st.cache_data.clear()
    st.rerun()

uploaded_file = st.sidebar.file_uploader(
    "อัปโหลดไฟล์ CSV (.csv)", 
    type=["csv"],
    accept_multiple_files=False
)

# 7. แสดงผล Header Metadata + กราฟพร้อมโซนเวลา
if uploaded_file:
    df, metadata = parse_single_file(uploaded_file)
    
    if df.empty:
        st.error("⚠️ ไม่สามารถอ่านข้อมูลจากไฟล์ที่อัปโหลดได้ กรุณาตรวจสอบว่าเป็นไฟล์ CSV จาก Datapaq หรือไม่")
    else:
        st.sidebar.success(f"โหลดไฟล์ {uploaded_file.name} สำเร็จ ({len(df)} แถว)")

        st.sidebar.markdown("---")
        st.sidebar.header("🎛️ Dynamic Controls")

        color_shading_mode = st.sidebar.radio(
            "เลือกโหมดแสดงสี:",
            ["แสดงสีตามโซน (By Zone)", "แสดงสีตามกลุ่มงาน (By Process Group)"],
            index=0
        )

        if color_shading_mode == "แสดงสีตามโซน (By Zone)":
            zones_data = [
                {"Start Time": "00:00:00", "End Time": "00:00:04", "Zone Name": "XFER", "Color": "#F7DC6F"},
                {"Start Time": "00:00:05", "End Time": "00:01:31", "Zone Name": "Dryer Z#1", "Color": "#F7DC6F"},
                {"Start Time": "00:01:32", "End Time": "00:02:59", "Zone Name": "Dryer Z#2", "Color": "#F39C12"},
                {"Start Time": "00:03:00", "End Time": "00:04:31", "Zone Name": "Dryer Z#3", "Color": "#F39C12"},
                {"Start Time": "00:04:32", "End Time": "00:07:11", "Zone Name": "XFER#1", "Color": "#E67E22"},
                {"Start Time": "00:07:12", "End Time": "00:10:01", "Zone Name": "Z#1", "Color": "#FF0033"},       
                {"Start Time": "00:10:02", "End Time": "00:12:11", "Zone Name": "Z#2", "Color": "#E6002E"},       
                {"Start Time": "00:12:12", "End Time": "00:13:59", "Zone Name": "Z#3", "Color": "#CC0029"},       
                {"Start Time": "00:14:00", "End Time": "00:15:39", "Zone Name": "Z#4", "Color": "#B30024"},       
                {"Start Time": "00:15:40", "End Time": "00:17:18", "Zone Name": "Z#5", "Color": "#CC0029"},       
                {"Start Time": "00:17:19", "End Time": "00:18:48", "Zone Name": "Z#6", "Color": "#E6002E"},       
                {"Start Time": "00:18:49", "End Time": "00:20:05", "Zone Name": "Z#7", "Color": "#FF0033"},       
                {"Start Time": "00:20:06", "End Time": "00:22:06", "Zone Name": "WatCool#1", "Color": "#00B4D8"},
                {"Start Time": "00:22:07", "End Time": "00:23:35", "Zone Name": "WatCool#2", "Color": "#0096C7"},
                {"Start Time": "00:23:36", "End Time": "00:25:05", "Zone Name": "Exit curtain box", "Color": "#0077B6"},
                {"Start Time": "00:25:06", "End Time": "00:25:33", "Zone Name": "XFER#2", "Color": "#023E8A"},
                {"Start Time": "00:25:34", "End Time": "00:26:33", "Zone Name": "AirCool#1", "Color": "#48CAE4"},
                {"Start Time": "00:26:34", "End Time": "00:27:32", "Zone Name": "AirCool#2", "Color": "#90E0EF"},
                {"Start Time": "00:27:33", "End Time": "00:28:59", "Zone Name": "Exit", "Color": "#CAF0F8"}
            ]
            angle_setting = -90
        else:
            zones_data = [
                {"Start Time": "00:00:00", "End Time": "00:05:00", "Zone Name": "Dryer", "Color": "#F39C12"},      
                {"Start Time": "00:05:01", "End Time": "00:20:05", "Zone Name": "Brazing", "Color": "#FF0033"},    
                {"Start Time": "00:20:06", "End Time": "00:27:32", "Zone Name": "Cool", "Color": "#00B4D8"},       
                {"Start Time": "00:27:33", "End Time": "00:28:00", "Zone Name": "Exit", "Color": "#90E0EF"}
            ]
            angle_setting = 0

        # 📌 ตัดข้อมูลกราฟให้แสดงแค่วินาทีที่ 1739 (00:28:59) เพื่อให้ครอบคลุมโซนทั้งหมดพอดี
        max_view_sec = 1739
        df_chart = df[df["ElapsedSeconds"] <= max_view_sec].copy()
        if df_chart.empty:
            df_chart = df.copy()

        # 📋 แสดงผล Header Metadata
        col_h1, col_h2 = st.columns(2)
        with col_h1:
            st.markdown(f"""
                <div class="raw-header-box">
                    <div><span class="raw-header-key">#paqfile start date</span> = <span class="raw-header-val">{metadata.get('paqfile start date', '-')}</span></div>
                    <div><span class="raw-header-key">#paqfile start time</span> = <span class="raw-header-val">{metadata.get('paqfile start time', '-')}</span></div>
                    <div><span class="raw-header-key">#title</span> = <span class="raw-header-val">{metadata.get('title', '-')}</span></div>
                </div>
            """, unsafe_allow_html=True)
        with col_h2:
            st.markdown(f"""
                <div class="raw-header-box">
                    <div><span class="raw-header-key">#operator</span> = <span class="raw-header-val">{metadata.get('operator', '-')}</span></div>
                    <div><span class="raw-header-key">#product</span> = <span class="raw-header-val">{metadata.get('product', '-')}</span></div>
                </div>
            """, unsafe_allow_html=True)

        # สร้างกราฟ Plotly
        fig = make_subplots(specs=[[{"secondary_y": False}]])
        
        probe_colors = [
            "#FF0000", "#00FF00", "#0000FF", "#8B4513", 
            "#FF00FF", "#DAA520", "#800080", "#00FFFF"
        ]

        probe_cols = [c for c in df_chart.columns if c.startswith("Probe #")]
        for idx, col in enumerate(probe_cols[:8]):
            fig.add_trace(
                go.Scatter(
                    x=df_chart["Time (HH:MM:SS)"],
                    y=df_chart[col],
                    name=col,
                    mode="lines",
                    line=dict(color=probe_colors[idx % len(probe_colors)], width=2)
                )
            )

        fig.add_trace(
            go.Scatter(
                x=df_chart["Distance (m)"],
                y=[None] * len(df_chart),
                xaxis="x2",
                showlegend=False,
                hoverinfo="skip"
            )
        )

        # 📌 เรียกใช้ฟังก์ชัน safe_zones_for_plotly เพื่อปรับพิกัดโซนสีไม่ให้หลุดขอบกราฟ
        render_zones = safe_zones_for_plotly(zones_data, df_chart)

        # แสดงแถบสีพื้นหลังตามโหมดที่ผู้ใช้เลือกใน Sidebar
        for idx, z_item in enumerate(render_zones):
            start_t = z_item["Start Time"]
            end_t = z_item["End Time"]
            z_name = z_item["Zone Name"]
            color_hex = z_item["Color"]

            fill_opacity = 0.28 if color_shading_mode == "แสดงสีตามกลุ่มงาน (By Process Group)" else 0.22
            fill_rgba = hex_to_rgba(color_hex, fill_opacity)
            
            fig.add_vrect(
                x0=start_t,
                x1=end_t,
                fillcolor=fill_rgba,
                layer="below",
                line_width=0
            )

            font_sz = 12 if color_shading_mode == "แสดงสีตามกลุ่มงาน (By Process Group)" else 10
            
            fig.add_annotation(
                x=start_t,
                y=610,
                text=f"<b>{z_name}</b>",
                showarrow=False,
                xanchor="left",
                yanchor="bottom",
                font=dict(color="#FFFFFF", size=font_sz, family="Arial Bold"),
                textangle=angle_setting
            )

        # คำนวณช่วง Tick สำหรับแกน Time ให้เหมาะสม และป้องกันการซ้อนทับที่ปลายกราฟ
        step_tick = max(1, len(df_chart) // 16)
        tick_indices = list(range(0, len(df_chart), step_tick))
        if len(tick_indices) > 0 and (len(df_chart) - 1) - tick_indices[-1] < (step_tick * 0.5):
            tick_indices[-1] = len(df_chart) - 1
        elif len(df_chart) > 0:
            tick_indices.append(len(df_chart) - 1)
            
        # สร้างรายการ Tick สำหรับแกน Distance โดยเฉพาะ
        max_dist = df_chart["Distance (m)"].max() if not df_chart.empty else 50.0
        if max_dist <= 20:
            dist_dtick = 1.0
        elif max_dist <= 50:
            dist_dtick = 2.0
        else:
            dist_dtick = 4.0

        fig.update_layout(
            template="plotly_dark",
            plot_bgcolor="#161b22",
            paper_bgcolor="#0e1117",
            hovermode="x unified",
            showlegend=True,
            legend=dict(
                font=dict(color="#FFFFFF", size=11, family="Arial Bold"),
                bgcolor="rgba(27, 31, 36, 0.95)",
                bordercolor="#F0B90B",
                borderwidth=1.5,
                orientation="v",
                yanchor="top",
                y=0.88,
                xanchor="left",
                x=1.02
            ),
            yaxis=dict(
                title=dict(text="Temperature (°C)", font=dict(color="#FFFFFF", size=12)),
                tickfont=dict(color="#CCCCCC", size=10),
                showgrid=True,
                gridcolor="rgba(255,255,255,0.08)",
                zeroline=False,
                linecolor="#555555",
                domain=[0.22, 1.0],
                range=[0, 650]
            ),
            xaxis=dict(
                title=dict(text="Time (hh:mm:ss)", font=dict(color="#FFFFFF", size=11)),
                tickmode="array",
                tickvals=df_chart.loc[tick_indices, "Time (HH:MM:SS)"].tolist(),
                tickfont=dict(color="#CCCCCC", size=10),
                showgrid=True,
                gridcolor="rgba(255,255,255,0.08)",
                showline=True,
                linewidth=1,
                linecolor="#888888",
                anchor="free",
                position=0.12,
                tickangle=0
            ),
            xaxis2=dict(
                title=dict(text="Distance (m)", font=dict(color="#F0B90B", size=11)),
                overlaying="x",
                anchor="free",
                position=0.00,
                
                tickmode="linear",
                tick0=0,
                dtick=dist_dtick,
                tickformat=".2f",
                
                range=[0, max_dist],
                
                tickfont=dict(color="#F0B90B", size=10),
                showgrid=False,
                showline=True,
                linewidth=1,
                linecolor="#F0B90B",
                
                minor=dict(
                    tickmode="linear",
                    tick0=0,
                    dtick=dist_dtick / 2,
                    ticklen=4,
                    tickcolor="#F0B90B",
                    showgrid=False
                ),
                tickangle=0,
                ticks="outside",
                ticklen=6,
            ),
            height=660,
            margin=dict(l=60, r=240, t=50, b=120)
        )

        st.plotly_chart(fig, use_container_width=True)

        # ---------------------------------------------------------
        # 📊 ตารางสรุปค่า
        # ---------------------------------------------------------
        st.markdown("### 📊 ตารางสรุปผลการวิเคราะห์ (Data Table for Google Sheets Copy)")

        # ครอบคลุมโซน Dryer (00:05:00 หรือ 300 วินาที)
        dryer_max_sec = 300
        
        dryer_subset = df[(df["ElapsedSeconds"] >= 0) & (df["ElapsedSeconds"] <= dryer_max_sec)]
        brazing_ht_subset = df[(df["ElapsedSeconds"] >= 0)]
        brazing_max_subset = df[(df["ElapsedSeconds"] >= 300) & (df["ElapsedSeconds"] <= 1205)]

        probe_order = [1, 2, 3, 4, 5, 6, 7, 8]
        ordered_cols = []
        for p_num in probe_order:
            for c in probe_cols[:8]:
                if f"Probe #{p_num}" in c:
                    ordered_cols.append((p_num, c))
                    break

        title_meta = metadata.get("title", "")

        summary_rows = []
        for idx, (p_num, col_name) in enumerate(ordered_cols):
            # ตรวจสอบตำแหน่ง Probe อัตโนมัติจาก #title
            location = get_probe_location(p_num, title_meta)

            short_pb_name = f"PB#{p_num}"
            
            # Max Temp
            br_max = f"{brazing_max_subset[col_name].max():.1f}" if not brazing_max_subset.empty else "0.0"
            d_max = f"{dryer_subset[col_name].max():.1f}" if not dryer_subset.empty else "0.0"
            
            # Dwell Time
            br_dwell_600 = (brazing_ht_subset[col_name] >= 600.0).sum() if not brazing_ht_subset.empty else 0
            br_dwell_591 = (brazing_ht_subset[col_name] >= 591.0).sum() if not brazing_ht_subset.empty else 0
            br_dwell_577 = (brazing_ht_subset[col_name] >= 577.0).sum() if not brazing_ht_subset.empty else 0
            
            d_dwell_300 = (dryer_subset[col_name] >= 300.0).sum() if not dryer_subset.empty else 0

            summary_rows.append([
                location,
                short_pb_name,
                br_max,
                d_max,
                format_seconds_to_time(br_dwell_600),
                format_seconds_to_time(br_dwell_591),
                format_seconds_to_time(br_dwell_577),
                format_seconds_to_time(d_dwell_300)
            ])

        multi_cols = pd.MultiIndex.from_tuples([
            ("", "Location"),
            ("", "Probe"),
            ("Max Temp (°C)", "Brazing"),
            ("Max Temp (°C)", "Dryer"),
            ("Brazing Zone", "Dwell Time Above 600°C"),
            ("Brazing Zone", "Dwell Time Above 591°C"),
            ("Brazing Zone", "Dwell Time Above 577°C"),
            ("Dryer Zone", "Dwell Time Above 300°C")
        ])

        display_summary_df = pd.DataFrame(summary_rows, columns=multi_cols)

        st.dataframe(display_summary_df, use_container_width=True, hide_index=True)

        # คำอธิบายเกณฑ์มาตรฐาน (Process Standards Legend)
        st.markdown("""
            <div style="background-color: #161b22; border: 1px solid #30363d; border-radius: 6px; padding: 12px 18px; font-size: 13px; color: #CCCCCC; margin-top: 10px;">
                <b style="color: #F0B90B;">📌 เกณฑ์มาตรฐานอ้างอิง (Process Standards):</b><br>
                • <b>Brazing Zone:</b> Max Temperature: <b>595 - 608 °C</b> | Dwell Time Above 600°C: <b>≤ 08:00 min (≤480s)</b> | Above 591°C: <b>02:00 - 12:00 min (120s - 720s)</b> | Above 577°C: <b>04:00 - 14:00 min (240s - 840s)</b><br>
                • <b>Dryer Zone:</b> Max Temperature: <b>200 - 375 °C</b> | Dwell Time Above 300°C: <b>> 2:00 min (>120s)</b>
            </div>
        """, unsafe_allow_html=True)

        # ส่วนตรวจสอบและเลือกดาวน์โหลด Excel (.xlsx)
        with st.expander("📋 ตรวจสอบและเลือกดาวน์โหลดตารางข้อมูล Excel (.xlsx)"):
            st.dataframe(df)
            
            st.markdown("---")
            st.markdown("##### 📥 ตัวเลือกการดาวน์โหลดไฟล์ Excel")
            
            col_opt1, col_opt2 = st.columns([2, 1])
            with col_opt1:
                custom_filename = st.text_input(
                    "ตั้งชื่อไฟล์ดาวน์โหลด:", 
                    value="datapaq_nb3_btm_8probes_data.xlsx"
                )
                if not custom_filename.endswith('.xlsx'):
                    custom_filename += '.xlsx'
                    
            with col_opt2:
                st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                excel_bytes = to_excel_bytes(df, summary_dataframe=display_summary_df, fig_plotly=fig)
                st.download_button(
                    label="📊 ดาวน์โหลดไฟล์ Excel (พร้อมตารางและแนบรูปกราฟ)",
                    data=excel_bytes,
                    file_name=custom_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

else:
    st.info("👈 กรุณาเลือกอัปโหลดไฟล์ (.csv) ที่เมนูด้านซ้าย")
