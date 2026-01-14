# File: main.py
import streamlit as st
import pandas as pd
import sqlite3
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime, date, timedelta
import matplotlib.dates as mdates
import io
import time
import xlsxwriter
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

from db_module import DBManager  
import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd

# 1. Tạo kết nối (Streamlit sẽ tự đọc link từ Secrets)
conn = st.connection("gsheets", type=GSheetsConnection)

# 2. Đọc dữ liệu hiện có từ Sheet
existing_data = conn.read(ttl=0) # ttl=0 để luôn lấy dữ liệu mới nhất không qua cache

# 3. Giả sử bạn có một form nhập liệu
with st.form("iqc_form"):
    ma_hang = st.text_input("Mã hàng")
    ket_qua = st.selectbox("Kết quả", ["Đạt", "Không đạt"])
    submit = st.form_submit_button("Lưu dữ liệu")

    if submit:
        # Tạo một DataFrame mới từ dữ liệu vừa nhập
        new_row = pd.DataFrame([{
            "Thời gian": pd.Timestamp.now(),
            "Mã hàng": ma_hang,
            "Kết quả": ket_qua
        }])
        
        # Gộp dữ liệu cũ và mới
        updated_df = pd.concat([existing_data, new_row], ignore_index=True)
        
        # Ghi ngược lại Google Sheets
        conn.update(data=updated_df)
        st.success("Đã lưu dữ liệu vào Google Sheets thành công!")

# --- CẤU HÌNH ---
st.set_page_config(page_title="QLCL Phòng Xét Nghiệm", layout="wide", page_icon="🔬")
db = DBManager("lab_data.db")

# --- STYLE CSS TÙY CHỈNH ---
st.markdown("""
<style>
    .footer {position: fixed; left: 0; bottom: 0; width: 100%; background-color: #f1f1f1; color: #333; text-align: center; padding: 10px; font-size: 14px; z-index: 999;}
    .block-container {padding-bottom: 50px;}
</style>
""", unsafe_allow_html=True)

# --- ĐỊNH NGHĨA HÀM TRƯỚC ---

def manage_test_mapping():
    st.subheader("🔗 Mapping Tên xét nghiệm từ máy")
    df_tests = db.get_all_tests()
    
    if df_tests.empty:
        st.warning("Chưa có xét nghiệm nào.")
        return

    # Tạo từ điển Tên -> ID
    test_options = dict(zip(df_tests['name'], df_tests['id']))
    
    col1, col2 = st.columns(2)
    with col1:
        # Chọn theo tên (chuỗi), không bao giờ lo lỗi subscriptable
        selected_name = st.selectbox("Chọn xét nghiệm trong PM:", list(test_options.keys()))
        selected_id = test_options[selected_name]
        
    with col2:
        external_name = st.text_input("Tên trên máy (VD: GLU):")
    
    if st.button("Thêm liên kết"):
        db.add_mapping(selected_id, external_name)
        st.success(f"Đã map {external_name} thành công!")

def process_bulk_import(df):
    # (Giữ nguyên logic xử lý database của bạn ở đây)
    # Hàm này dùng để chạy vòng lặp insert dữ liệu
    conn = sqlite3.connect("lab_data.db")
    # ... logic như bạn đã viết ...
    return summary
def get_clean_stats_3sigma(df):
    if df.empty or len(df) < 2:
        return None
    
    values = pd.to_numeric(df['value'], errors='coerce').dropna()
    n_original = len(values)
    
    if n_original < 2:
        return None

    mean = values.mean()
    sd = values.std()
    
    if sd == 0:
        return {'n': n_original, 'mean': mean, 'sd': 0, 'cv': 0.0001, 'outliers': 0}

    # Bộ lọc Outlier 3SD
    clean_values = values[(values >= mean - 3*sd) & (values <= mean + 3*sd)]
    n_clean = len(clean_values)
    outliers_count = n_original - n_clean # Tính số lượng bị loại bỏ
    
    if n_clean < 2:
        return {
            'n': n_original, 
            'mean': mean, 
            'sd': sd, 
            'cv': (sd / mean) * 100 if mean != 0 else 0,
            'outliers': 0
        }

    return {
        'n': n_clean,
        'mean': clean_values.mean(),
        'sd': clean_values.std(),
        'cv': (clean_values.std() / clean_values.mean()) * 100 if clean_values.mean() != 0 else 0,
        'outliers': outliers_count # BẮT BUỘC PHẢI CÓ DÒNG NÀY
    }
def clean_outliers_3sigma(df, column='value', iterations=1):
    """
    Lọc bỏ các giá trị ngoại lai dựa trên quy tắc 3-SD.
    iterations: Số lần lặp lại việc lọc (thường dùng 1 hoặc 2).
    """
    df_clean = df.copy()
    outliers_detected = pd.DataFrame()

    for i in range(iterations):
        if len(df_clean) < 3:  # Không đủ dữ liệu để tính SD
            break
            
        mean = df_clean[column].mean()
        sd = df_clean[column].std()
        
        lower_bound = mean - 3 * sd
        upper_bound = mean + 3 * sd
        
        # Xác định Outliers
        is_outlier = (df_clean[column] < lower_bound) | (df_clean[column] > upper_bound)
        
        if not is_outlier.any():
            break
            
        # Lưu lại danh sách bị loại để báo cáo
        outliers_detected = pd.concat([outliers_detected, df_clean[is_outlier]])
        
        # Giữ lại dữ liệu sạch
        df_clean = df_clean[~is_outlier]
        
    return df_clean, outliers_detected

def get_stats_real_v2(df_input):
    """
    Hàm tính toán thống kê sau khi đã lọc Outliers.
    """
    if df_input.empty:
        return 0, 0, 0
    
    # Thực hiện lọc
    df_clean, df_outliers = clean_outliers_3sigma(df_input)
    
    m_lab = df_clean['value'].mean()
    sd_lab = df_clean['value'].std()
    cv_lab = (sd_lab / m_lab * 100) if m_lab > 0 else 0
    
    return m_lab, sd_lab, cv_lab, df_outliers
def export_mu_excel(test_name, mu_results, target_mau):
    """Xuất báo cáo MU ra file Excel"""
    output = io.BytesIO()
    report_list = []
    for lvl, res in mu_results.items():
        report_list.append({
            "Xét nghiệm": test_name,
            "Mức độ": f"Level {lvl}",
            "Số mẫu (n)": res['n_count'],
            "Trung bình": round(res['mean'], 4),
            "u_prec (Độ chụm %)": round(res['u_prec'], 2),
            "u_bias (Độ đúng %)": round(res['u_bias'], 2),
            "u_ref (Tham chiếu %)": round(res['u_ref'], 2),
            "Ue (KĐB mở rộng %)": round(res['ue'], 2),
            "Mục tiêu MAU (%)": round(target_mau, 2),
            "Đánh giá": "Đạt" if res['ue'] <= target_mau else "Không đạt"
        })
    
    df_report = pd.DataFrame(report_list)
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_report.to_excel(writer, index=False, sheet_name='Bao_Cao_MU')
    return output.getvalue()

def công_cụ_tạo_mẫu(df, filename):
    """Hàm chuyển đổi DataFrame thành dữ liệu Excel để tải về"""
    output = io.BytesIO()
    # Sử dụng engine xlsxwriter hoặc openpyxl
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

def plot_sigma_chart(sigma_plot_data, tea):
    # 1. Khởi tạo Figure nhỏ gọn
    fig, ax = plt.subplots(figsize=(5, 4), facecolor='white')
    
    # Thiết lập giới hạn trục dựa trên TEa
    max_cv = tea / 2
    max_bias = tea
    x_range = np.linspace(0, max_cv, 100)
    
    # 2. Định nghĩa màu sắc và nhãn
    sigma_levels = [
        (6, 'green', '6σ'),
        (5, 'blue', '5σ'),
        (4, 'purple', '4σ'),
        (3, 'orange', '3σ'),
        (2, 'red', '2σ')
    ]

    for s, color, label in sigma_levels:
        # Công thức: y = tea - s*x
        bias_line = tea - (s * x_range)
        bias_line = np.maximum(bias_line, 0)
        
        # Vẽ đường nét đứt
        ax.plot(x_range, bias_line, linestyle='--', color=color, linewidth=1.2, alpha=0.7)
        
        # --- CHỈNH SỬA NHÃN CHẠY THEO ĐƯỜNG ---
        # Chọn một điểm x đại diện (ví dụ 10% chiều rộng trục X) để đặt nhãn
        tx = max_cv * 0.2 
        ty = tea - (s * tx)
        
        if ty > 0:
            # Tính toán góc xoay (rotation) dựa trên độ dốc s
            # s càng lớn đường càng đứng, s càng nhỏ đường càng nằm ngang
            # Công thức xấp xỉ góc xoay để nhãn song song với đường
            angle = np.degrees(np.arctan2(-s * (max_cv/max_bias), 1)) 
            
            ax.text(tx, ty + (tea * 0.01), label, color=color, 
                    fontsize=9, fontweight='bold', 
                    rotation=angle, rotation_mode='anchor')

    # 3. Vẽ các điểm QC thực tế # Bảng màu chuẩn: Blue, Orange, Red
    colors_qc = ['#0000ff', '#ff7f0e', '#ff0000'] 
    
    for i, pt in enumerate(sigma_plot_data):
        label_text = pt.get('label', f'L{i+1}')
        color = colors_qc[i] if i < len(colors_qc) else '#7f7f7f'
        
        # Điểm dữ liệu hình tròn
        ax.scatter(pt['cv'], pt['bias'], s=90, color=color, marker='o', 
                   label=label_text, edgecolors='white', linewidth=1, zorder=10)
        
        # Đường dóng mờ
        ax.vlines(pt['cv'], 0, pt['bias'], linestyle=':', color=color, alpha=0.4)
        ax.hlines(pt['bias'], 0, pt['cv'], linestyle=':', color=color, alpha=0.4)

    # 4. Định dạng biểu đồ tối giản
    ax.set_title(f"Method Decision Chart (TEa = {tea}%)", fontsize=11, fontweight='bold', pad=10)
    ax.set_xlabel("Precision (CV %)", fontsize=9)
    ax.set_ylabel("Inaccuracy (Bias %)", fontsize=9)
    
    ax.set_xlim(0, max_cv)
    ax.set_ylim(0, max_bias)
    
    # Lưới xám nhạt
    ax.grid(True, linestyle='-', color='lightgray', alpha=0.4)
    
    # Legend
    ax.legend(loc='upper right', fontsize='8', frameon=True)

    # Loại bỏ khung viền Top/Right
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    plt.tight_layout()
    return fig
def get_stats_real(df_level):
    """
    Tính toán các thông số thống kê thực tế (Mean, SD, CV%) từ dữ liệu IQC.
    Hàm này được dùng cho cả tính MU (Độ không đảm bảo đo) và Six Sigma.
    """
    if df_level.empty or len(df_level) < 2:
        return 0.0, 0.0, 0.0
    
    # Tính toán các chỉ số cơ bản
    mean_val = df_level['value'].mean()
    sd_val = df_level['value'].std()
    
    # Tính CV%, tránh lỗi chia cho 0 nếu Mean = 0
    cv_val = (sd_val / mean_val * 100) if mean_val != 0 else 0.0
    
    return mean_val, sd_val, cv_val
def calculate_qgi(bias_pct, cv_pct):
    """
    Tính Quality Goal Index (QGI) để phân tích nguyên nhân khi chỉ số Sigma thấp.
    QGI giúp xác định lỗi do Độ đúng (Bias) hay Độ chụm (CV).
    """
    # Tránh lỗi chia cho 0
    if cv_pct == 0: 
        return 0.0, "Không xác định (CV=0)"
    
    # Công thức: QGI = Bias / (1.5 * CV)
    qgi = abs(bias_pct) / (1.5 * cv_pct)
    
    if qgi < 0.8: 
        reason = "Lỗi do ĐỘ CHỤM (Precision) - Ưu tiên cải thiện CV (bảo trì máy, thay kim, thuốc thử)"
    elif 0.8 <= qgi <= 1.2: 
        reason = "Lỗi do cả ĐỘ CHỤM & ĐỘ ĐÚNG - Cần xem xét toàn diện"
    else: 
        reason = "Lỗi do ĐỘ ĐÚNG (Accuracy) - Ưu tiên kiểm tra Bias (chuẩn lại máy, xem lại giá trị đích)"
        
    return qgi, reason
def upgrade_database_structure():
    import sqlite3
    conn = None 
    # SỬA TÊN FILE TẠI ĐÂY ĐỂ KHỚP VỚI CẤU HÌNH CỦA BẠN
    db_file = "lab_data.db" 
    try:
        conn = sqlite3.connect(db_file) 
        cursor = conn.cursor()
        
        # 1. Tự động tạo bảng nếu chưa tồn tại (đảm bảo không bị trống [])
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS iqc_results (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                lot_id INTEGER,
                date TIMESTAMP,
                level INTEGER,
                value REAL,
                note TEXT,
                action TEXT DEFAULT '' 
            )
        ''')
        conn.commit()
        
        # 2. Kiểm tra lại danh sách cột để bổ sung 'action' nếu thiếu
        cursor.execute("PRAGMA table_info(iqc_results)")
        columns = [column[1] for column in cursor.fetchall()]
        
        if 'action' not in columns:
            cursor.execute("ALTER TABLE iqc_results ADD COLUMN action TEXT DEFAULT ''")
            conn.commit()
            return True, f"Thành công: Đã kết nối '{db_file}' và cấu hình cột 'action'."
        else:
            return True, f"Hệ thống sẵn sàng: File '{db_file}' đã có đầy đủ cấu trúc."
            
    except Exception as e:
        return False, f"Lỗi SQL: {str(e)}"
    finally:
        if conn is not None:
            conn.close()

def show_qc_dashboard():
    st.header("📊 Bảng theo dõi chất lượng tổng thể")
    all_tests = db.get_all_tests()
    
    # Tạo lưới hiển thị (4 cột)
    cols = st.columns(4)
    for i, test in enumerate(all_tests):
        with cols[i % 4]:
            # Lấy kết quả mới nhất của XN này
            latest_status = db.get_latest_westgard_status(test['id']) 
            
            if latest_status == "OK":
                st.info(f"✅ **{test['name']}**\n\nTrạng thái: Đạt")
            else:
                st.error(f"❌ **{test['name']}**\n\nLỗi: {latest_status}")
def check_westgard_multi_level(df):
    """
    Kiểm tra lỗi 6x, 9x, 12x cho 3 mức nồng độ cùng 1 bên so với đường trung tâm.
    Quy tắc: n kết quả liên tiếp (gộp cả 3 levels) nằm cùng phía so với Mean.
    """
    # Sắp xếp toàn bộ dữ liệu theo thời gian
    df_sorted = df.sort_values('date').copy()
    if len(df_sorted) < 6:
        return []

    # Tính Z-score cho từng dòng để biết nằm bên nào của đường Mean (0)
    df_sorted['side'] = df_sorted.apply(
        lambda r: 1 if (r['value'] - r['target_mean']) > 0 else -1 if (r['value'] - r['target_mean']) < 0 else 0,
        axis=1
    )
    
    violation_logs = []
    sides = df_sorted['side'].tolist()
    dates = df_sorted['date'].tolist()
    levels = df_sorted['level'].tolist()

    for i in range(len(sides)):
        # Kiểm tra 12x (4 lượt chạy x 3 mức = 12 điểm liên tiếp)
        if i >= 11:
            window = sides[i-11:i+1]
            if all(x == 1 for x in window) or all(x == -1 for x in window):
                violation_logs.append(f"❌ Lỗi 12x: 4 lượt chạy (12 điểm) cùng bên tại {dates[i].strftime('%d/%m %H:%M')}")
                continue # Đã dính lỗi nặng nhất thì bỏ qua các lỗi nhỏ hơn tại điểm đó

        # Kiểm tra 9x (3 lượt chạy x 3 mức = 9 điểm liên tiếp)
        if i >= 8:
            window = sides[i-8:i+1]
            if all(x == 1 for x in window) or all(x == -1 for x in window):
                violation_logs.append(f"⚠️ Lỗi 9x: 3 lượt chạy (9 điểm) cùng bên tại {dates[i].strftime('%d/%m %H:%M')}")
                continue

        # Kiểm tra 6x (2 lượt chạy x 3 mức = 6 điểm liên tiếp)
        if i >= 5:
            window = sides[i-5:i+1]
            if all(x == 1 for x in window) or all(x == -1 for x in window):
                violation_logs.append(f"ℹ️ Lỗi 6x: 2 lượt chạy (6 điểm) cùng bên tại {dates[i].strftime('%d/%m %H:%M')}")

    return violation_logs
def get_westgard_violations(df, mean_map, sd_map):
    if df is None or df.empty:
        return df

    df = df.copy()
    if 'id' not in df.columns: df['id'] = range(len(df))
    df['date'] = pd.to_datetime(df['date'], format='mixed', dayfirst=True, errors='coerce')
    df_calc = df.dropna(subset=['date']).sort_values(by=['date', 'level']).copy()
    
    def calc_z(row):
        try:
            lvl = row['level']
            val = float(row['value'])
            
            # Lấy Mean/SD từ dict hoặc từ giá trị đơn lẻ một cách an toàn
            m = mean_map.get(lvl, 0) if isinstance(mean_map, dict) else mean_map
            s = sd_map.get(lvl, 0) if isinstance(sd_map, dict) else sd_map
            
            return (val - m) / s if s > 0 else 0
        except Exception:
            return 0

    # Tính toán Z-score an toàn, không còn lỗi dict > int
    df_calc['z_score'] = df_calc.apply(calc_z, axis=1)
    
    violation_map = {row_id: set() for row_id in df_calc['id']}

# --- 1. KIỂM TRA ACROSS-LEVEL (Cập nhật cho 3 mức) ---
    groups = [group for _, group in df_calc.groupby('date')]
    for i in range(len(groups)):
        df_day = groups[i]
        # Lấy dữ liệu 3 mức của ngày đó
        l1 = df_day[df_day['level'] == 1].head(1)
        l2 = df_day[df_day['level'] == 2].head(1)
        l3 = df_day[df_day['level'] == 3].head(1)
        
        levels_present = [l for l in [l1, l2, l3] if not l.empty]
        
        # Kiểm tra R-4s giữa bất kỳ cặp mức nào (1-2, 2-3, 1-3)
        if len(levels_present) >= 2:
            for a in range(len(levels_present)):
                for b in range(a + 1, len(levels_present)):
                    z_a = levels_present[a]['z_score'].iloc[0]
                    z_b = levels_present[b]['z_score'].iloc[0]
                    if (z_a >= 2 and z_b <= -2) or (z_a <= -2 and z_b >= 2):
                        violation_map[levels_present[a]['id'].iloc[0]].add("R-4s")
                        violation_map[levels_present[b]['id'].iloc[0]].add("R-4s")
        
        # 2-2s (Across): Cả 3 mức (hoặc 2/3 mức) cùng vi phạm > 2SD về 1 phía
        z_scores = [l['z_score'].iloc[0] for l in levels_present]
        if len(z_scores) >= 2:
            if all(z > 2 for z in z_scores) or all(z < -2 for z in z_scores):
                 for l in levels_present: violation_map[l['id'].iloc[0]].add("2-2s")

    # --- 1. KIỂM TRA ACROSS-LEVEL (So sánh giữa các mức) ---
    groups = [group for _, group in df_calc.groupby('date')]
    for i in range(len(groups)):
        df_day = groups[i]
        l1_curr = df_day[df_day['level'] == 1].head(1)
        l2_curr = df_day[df_day['level'] == 2].head(1)
        
        if not l1_curr.empty and not l2_curr.empty:
            z1, z2 = l1_curr['z_score'].iloc[0], l2_curr['z_score'].iloc[0]
            id1, id2 = l1_curr['id'].iloc[0], l2_curr['id'].iloc[0]

            # R-4s: 1 cái > +2SD và 1 cái < -2SD
            if (z1 >= 2 and z2 <= -2) or (z1 <= -2 and z2 >= 2):
                violation_map[id1].add("R-4s"); violation_map[id2].add("R-4s")

            # 2-2s (Across): Cả 2 mức cùng nằm 1 bên và rơi vào khoảng ±2SD đến ±3SD
            if (2 < z1 < 3 and 2 < z2 < 3) or (-3 < z1 < -2 and -3 < z2 < -2):
                violation_map[id1].add("2-2s") ; violation_map[id2].add("2-2s")

            # 4-1s (Across): 2 phiên liên tiếp của 2 mức cùng phía > 1SD
            if i >= 1:
                prev_g = groups[i-1]
                l1p, l2p = prev_g[prev_g['level']==1], prev_g[prev_g['level']==2]
                if not l1p.empty and not l2p.empty:
                    zs = [z1, z2, l1p['z_score'].iloc[0], l2p['z_score'].iloc[0]]
                    ids = [id1, id2, l1p['id'].iloc[0], l2p['id'].iloc[0]]
                    if all(v > 1 for v in zs) or all(v < -1 for v in zs):
                        for tid in ids: violation_map[tid].add("4-1s")

            # 10x (Across): 5 phiên liên tiếp của 2 mức cùng phía Mean
            if i >= 4:
                combined_z = []
                combined_ids = []
                for k in range(i-4, i+1):
                    combined_z.extend(groups[k]['z_score'].tolist())
                    combined_ids.extend(groups[k]['id'].tolist())
                if len(combined_z) >= 10 and (all(v > 0 for v in combined_z) or all(v < 0 for v in combined_z)):
                    for tid in combined_ids: violation_map[tid].add("10x")

    # --- 2. KIỂM TRA WITHIN-LEVEL (Chuỗi thời gian từng mức) ---
    for level, df_level in df_calc.groupby('level'):
        df_level = df_level.sort_values(by='date').reset_index(drop=True)
        z, ids = df_level['z_score'].tolist(), df_level['id'].tolist()
        for i in range(len(z)):
            cid = ids[i]
            if abs(z[i]) > 3: violation_map[cid].add("1-3s")
            if i >= 1 and ((2 < z[i] < 3 and 2 < z[i-1] < 3) or (-3 < z[i] < -2 and -3 < z[i-1] < -2)):
                violation_map[cid].add("2-2s")
            if i >= 3:
                sub4 = z[i-3:i+1]
                if all(v > 1 for v in sub4) or all(v < -1 for v in sub4): violation_map[cid].add("4-1s")
            if i >= 5:
                sub6 = z[i-5:i+1]
                if all(v > 1 for v in sub6) or all(v < -1 for v in sub6): violation_map[cid].add("Shift")
                if all(sub6[k] < sub6[k+1] for k in range(5)): violation_map[cid].add("Trend (+)")
                elif all(sub6[k] > sub6[k+1] for k in range(5)): violation_map[cid].add("Trend (-)")
            if i >= 9:
                sub10 = z[i-9:i+1]
                if all(v > 0 for v in sub10) or all(v < 0 for v in sub10): violation_map[cid].add("10x")
            if not violation_map[cid] and 2 < abs(z[i]) <= 3: violation_map[cid].add("1-2s")

    final_res = {k: ", ".join(sorted(list(v))) for k, v in violation_map.items()}
    df['Violation'] = df['id'].map(final_res).replace("", "ĐẠT").fillna("ĐẠT")
    return df

# 2. HÀM KIỂM TRA WESTGARD CHÍNH
def check_westgard_rules(df_all):
    """
    Hàm chính kiểm tra toàn bộ quy tắc Westgard (Within & Across).
    Input: DataFrame chứa toàn bộ dữ liệu IQC (cần có cột: id, date, level, z_score).
    Output: Tuple (final_rejections, final_warnings)
            Mỗi phần tử là list các tuple: (iqc_id, "Tên lỗi", "Mức độ")
    """
    
    if df_all.empty or 'z_score' not in df_all.columns:
        return ([], [])

    # Sắp xếp dữ liệu theo thời gian
    df_all = df_all.sort_values(by=['date', 'level']).reset_index(drop=True)

    rejection_details = {} # {id: "Rule name"}
    warning_details = {}   # {id: "Rule name"}

    # --- BƯỚC 1: KIỂM TRA ACROSS-LEVELS (Check từng ngày) ---
    for date, df_day in df_all.groupby('date'):
        cross_errors = check_cross_level_rules(df_day)
        # Cập nhật lỗi vào danh sách từ chối
        rejection_details.update(cross_errors)

    # --- BƯỚC 2: KIỂM TRA WITHIN-LEVEL (Check chuỗi thời gian của từng Level) ---
    
    # Tạo bản sao để xử lý
    df_temp = df_all.copy()
    
    # Lặp qua từng Level (L1, L2)
    for level, df_level in df_temp.groupby('level'):
        df_level = df_level.sort_values(by='date').reset_index(drop=True)
        z_values = df_level['z_score'].tolist()
        id_values = df_level['id'].tolist()
        n = len(z_values)
        
        for i in range(n):
            current_id = id_values[i]
            current_z = z_values[i]
            
     # Nếu điểm này đã bị lỗi Across-Level (R4s, v.v) thì bỏ qua
            if current_id in rejection_details:
                continue

    # === QUY TẮC TỪ CHỐI (REJECTION) ===
            
      # 1-3s: Một điểm nằm ngoài ±3SD
            if abs(current_z) > 3:
                rejection_details[current_id] = "1-3s: Điểm vượt quá 3SD (Lỗi Ngẫu nhiên)"
                continue
                
     # 2-2s (Within): Hai điểm liên tiếp cùng phía nằm ngoài ±2SD
            if i >= 1:
                prev_z = z_values[i-1]
                if ((current_z > 2 and prev_z > 2) or (current_z < -2 and prev_z < -2)):
                    rule = "2-2s(Within): 2 điểm liên tiếp vượt 2SD (Lỗi Hệ thống)"
                    rejection_details[current_id] = rule
                    rejection_details[id_values[i-1]] = rule # Đánh dấu cả điểm trước
                    continue
            
    # 4-1s (Within): Bốn điểm liên tiếp cùng phía ngoài ±1SD
            if i >= 3:
                last_4_z = z_values[i-3:i+1]
                if all(z > 1 for z in last_4_z) or all(z < -1 for z in last_4_z):
                    rule = "4-1s: 4 điểm liên tiếp vượt 1SD (Lỗi Hệ thống)"
                    for k in range(4): rejection_details[id_values[i-k]] = rule
                    continue
            
    # 10x (Shift): 10 điểm liên tiếp cùng phía Mean
            if i >= 9:
                last_10_z = z_values[i-9:i+1]
                if all(z > 0 for z in last_10_z) or all(z < 0 for z in last_10_z):
                    rule = "10x: 10 điểm liên tiếp cùng phía Mean (Shift)"
                    for k in range(10): rejection_details[id_values[i-k]] = rule
                    continue

            # === QUY TẮC CẢNH BÁO (WARNING) ===
            
     # 1-2s: Điểm nằm ngoài ±2SD (nhưng < 3SD)
            if current_id not in rejection_details:
                if abs(current_z) > 2 and abs(current_z) <= 3:
                    warning_details[current_id] = "1-2s: Cảnh báo (Điểm vượt 2SD)"
            
   # Trend: 6 điểm liên tiếp tăng hoặc giảm
            if i >= 5:
                last_6 = z_values[i-5:i+1]
          # Tăng dần
                if all(last_6[k] < last_6[k+1] for k in range(5)):
                    warning_details[current_id] = "Trend: 6 điểm tăng dần liên tiếp"
          # Giảm dần
                elif all(last_6[k] > last_6[k+1] for k in range(5)):
                    warning_details[current_id] = "Trend: 6 điểm giảm dần liên tiếp"


    # --- BƯỚC 3: TỔNG HỢP KẾT QUẢ ---

    final_rejections = []
    final_warnings = []
    
    # Duyệt lại df_all để giữ thứ tự thời gian khi trả về
    for index, row in df_all.iterrows():
        iqc_id = row['id']
        
        # Ưu tiên lỗi REJECTION trước
        if iqc_id in rejection_details:
            final_rejections.append((iqc_id, rejection_details[iqc_id], "REJECTION"))
            
        # Nếu không Rejection thì xem có Warning không
        elif iqc_id in warning_details:
            final_warnings.append((iqc_id, warning_details[iqc_id], "WARNING"))
            
    # Trả về kết quả
    return final_rejections, final_warnings



def plot_levey_jennings(df, title, show_legend=True):
    """
    Vẽ biểu đồ Levey-Jennings dựa trên Z-Score.
    Đã xử lý lỗi thiếu cột và định dạng ngày tháng hiển thị sai.
    """
    if df.empty: 
        return None
    
    # 1. Đảm bảo cột date là định dạng datetime để matplotlib xử lý đúng trục X
    df = df.copy()
    # Sử dụng dayfirst=True để tránh lỗi đảo ngược ngày/tháng
    df['date'] = pd.to_datetime(df['date'], dayfirst=True, errors='coerce')
    df = df.dropna(subset=['date'])
    # Sắp xếp toàn bộ dataframe theo ngày để tránh đường nối bị nhảy ngược
    df = df.sort_values('date')

    fig, ax = plt.subplots(figsize=(11, 6))
    
    # 2. Vẽ các vùng giới hạn SD (Duy trì các đường nằm ngang cố định tại Z = 0, 1, 2, 3)
    ax.axhline(0, color='green', lw=2, label='Mean (Target)')
    
    # Vẽ các đường SD với nhãn cụ thể
    sd_config = {
        1: {'color': 'gold', 'label': '±1SD'},
        2: {'color': 'red', 'label': '±2SD (Warning)'},
        3: {'color': 'black', 'label': '±3SD (Reject)'}
    }
    
    for sd, config in sd_config.items():
        ax.axhline(sd, color=config['color'], ls='--', alpha=0.6, lw=1)
        ax.axhline(-sd, color=config['color'], ls='--', alpha=0.6, lw=1)
        # Ghi chú nhãn ở mép phải biểu đồ (sử dụng ngày cuối cùng trong dữ liệu)
        last_date = df['date'].max()
        ax.text(last_date, sd, f" +{sd}SD", va='center', fontsize=8, color=config['color'])
        ax.text(last_date, -sd, f" -{sd}SD", va='center', fontsize=8, color=config['color'])
    
    colors = {1: 'blue', 2: 'orange', 3: 'red'}
        
    # 3. Tính Z-Score và Vẽ dữ liệu từng Level
    for lvl in [1, 2, 3]:
        d_lvl = df[df['level'] == lvl].copy()
        if d_lvl.empty:
            continue
            
        # Kiểm tra xem có đủ cột để tính toán không (Tránh KeyError)
        if 'target_mean' in d_lvl.columns and 'target_sd' in d_lvl.columns:
            # Tránh chia cho 0 nếu SD chưa được thiết lập
            d_lvl['z'] = d_lvl.apply(
                lambda r: (r['value'] - r['target_mean']) / r['target_sd'] if r['target_sd'] > 0 else 0, 
                axis=1
            )
        else:
            d_lvl['z'] = 0 
            
        # Vẽ đường nối và điểm dữ liệu
        ax.plot(d_lvl['date'], d_lvl['z'], color=colors[lvl], alpha=0.4, lw=1.5, zorder=2)
        ax.scatter(d_lvl['date'], d_lvl['z'], color=colors[lvl], s=40, 
                   label=f"Level {lvl}", edgecolors='white', zorder=4)
        
        # 4. Đánh dấu thay đổi Lot
        if 'lot_number' in d_lvl.columns and not d_lvl['lot_number'].isnull().all():
            changes = d_lvl.drop_duplicates(subset=['lot_number'], keep='first')
            for _, r in changes.iterrows():
                if r['date'] != df['date'].min():
                    ax.axvline(r['date'], color='gray', ls=':', alpha=0.4, zorder=1)
                    ax.text(r['date'], 3.8, f" Lot: {r['lot_number']}", 
                            rotation=90, fontsize=7, color='gray', va='top')
# 5. Kiểm tra Westgard và hiển thị thông báo
    violations = check_westgard_multi_level(df)
    
    # Hiển thị kết quả kiểm tra Westgard trực tiếp dưới biểu đồ bằng Streamlit
    if violations:
        with st.expander("🚨 CẢNH BÁO QUY TẮC WESTGARD (6x, 9x, 12x)", expanded=True):
            for v in violations[-5:]: # Hiển thị 5 lỗi gần nhất
                st.write(v)
    # 5. CẤU HÌNH ĐỊNH DẠNG NGÀY THÁNG (SỬA LỖI HIỂN THỊ)
    # Định dạng trục X hiển thị: Ngày/Tháng Giờ:Phút
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%d/%m'))
    
    # Thiết lập khoảng cách chia (tự động điều chỉnh để không quá dày)
    ax.xaxis.set_major_locator(mdates.AutoDateLocator())

    ax.set_ylim(-4.5, 4.5) 
    ax.set_ylabel("Z-Score (Độ lệch chuẩn)")
    ax.set_xlabel("Thời gian thực hiện")
    ax.set_title(title, fontweight='bold', pad=15)
    
    # Tự động xoay ngày tháng trên trục X và căn chỉnh
    fig.autofmt_xdate(rotation=30, ha='right')
    
    if show_legend:
        handles, labels = ax.get_legend_handles_labels()
        by_label = dict(zip(labels, handles))
        ax.legend(by_label.values(), by_label.keys(), loc='upper left', bbox_to_anchor=(1, 1))

    plt.tight_layout()
    return fig

def plot_cusum_chart(df_eqa):
    """
    Vẽ biểu đồ CUSUM với V-Mask (Góc 28°, d=10)
    Hàm trả về: (figure, is_violated)
    """
    if df_eqa.empty or len(df_eqa) < 2:
        return None, False

    # 1. Chuẩn bị dữ liệu
    # Đảm bảo cột date là datetime và dữ liệu được sắp xếp
    df_plot = df_eqa.copy().sort_values('date')
    dates = pd.to_datetime(df_plot['date'])
    cusum_values = df_plot['CUSUM'].values
    n_points = len(cusum_values)
    indices = np.arange(n_points)
    
    # 2. Thiết lập Figure tối giản nền trắng
    fig, ax = plt.subplots(figsize=(10, 5), facecolor='white')
    
    # 3. Tính toán V-Mask (Góc 28 độ, d=10)
    last_x = indices[-1]
    last_y = cusum_values[-1]
    theta_deg = 28 
    d = 10         
    k = np.tan(np.radians(theta_deg))
    
    vertex_x = last_x + d
    vertex_y = last_y
    
    # Vẽ đường mặt nạ
    x_mask = np.linspace(max(0, last_x - 30), vertex_x, 100)
    y_upper = vertex_y + k * (vertex_x - x_mask)
    y_lower = vertex_y - k * (vertex_x - x_mask)
    
    # 4. Vẽ CUSUM (Đường màu tím theo yêu cầu)
    ax.plot(indices, cusum_values, marker='o', linestyle='-', color='purple', 
            linewidth=2, label='CUSUM Line', zorder=3)
    
    # 5. Vẽ V-Mask (Nét đứt màu đỏ)
    ax.plot(x_mask, y_upper, color='red', linestyle='--', alpha=0.6, label='V-Mask Limit')
    ax.plot(x_mask, y_lower, color='red', linestyle='--', alpha=0.6)
    ax.plot(vertex_x, vertex_y, marker='x', color='black', label='Vertex')
    
    # 6. Kiểm tra vi phạm V-Mask
    is_violated = False
    for i in range(n_points):
        limit_upper = vertex_y + k * (vertex_x - i)
        limit_lower = vertex_y - k * (vertex_x - i)
        if cusum_values[i] > limit_upper or cusum_values[i] < limit_lower:
            is_violated = True
            ax.scatter(i, cusum_values[i], color='orange', s=100, edgecolors='black', zorder=5)

    # 7. Định dạng trục và lưới
    ax.axhline(0, color='black', linewidth=0.8)
    ax.set_title(f"Biểu đồ CUSUM & V-Mask (ISO 13528)", fontsize=12, fontweight='bold')
    ax.set_ylabel("CUSUM (SDI Tích lũy)")
    
    # Hiển thị ngày tháng trục X
    ax.set_xticks(indices)
    ax.set_xticklabels([d.strftime('%d/%m') for d in dates], rotation=45, fontsize=8)
    
    ax.grid(True, linestyle='-', color='lightgray', alpha=0.3)
    ax.legend(loc='upper left', fontsize='small')
    
    plt.tight_layout()
    return fig, is_violated

import numpy as np
import scipy.stats as stats
def handle_outliers_grubbs(matrix):
    """
    Tự động phát hiện và xử lý giá trị ngoại lệ theo chuẩn EP15-A3.
    Hệ số G tới hạn cho n=25 là 3.135.
    """
    flat_data = [item for sublist in matrix for item in sublist]
    n = len(flat_data)
    mean = np.mean(flat_data)
    sd = np.std(flat_data, ddof=1)
    
    g_critical = 3.135 # Giá trị tới hạn cho n=25, alpha=0.05
    
    outliers = []
    cleaned_matrix = []
    
    # Duyệt từng điểm dữ liệu
    for i, day in enumerate(matrix):
        new_day = []
        for val in day:
            g_score = abs(val - mean) / sd
            if g_score > g_critical:
                outliers.append({"day": i+1, "value": val, "g_score": g_score})
                # Thay thế giá trị ngoại lệ bằng trung bình của ngày đó (để không làm hỏng ANOVA)
                # Hoặc có thể dùng np.nan nếu hàm ANOVA của bạn xử lý được
                new_day.append(np.mean(day)) 
            else:
                new_day.append(val)
        cleaned_matrix.append(new_day)
        
    return cleaned_matrix, outliers
def calculate_clsi_ep15_a3_final(matrix, claim_sr, claim_sl, target_mean):
    # 1. Xử lý ngoại lệ trước khi tính toán
    cleaned_matrix, found_outliers = handle_outliers_grubbs(matrix)
    
    n_run = 5
    n_rep = 5
    
    # 2. ANOVA trên dữ liệu đã làm sạch
    flat_data = [item for sublist in cleaned_matrix for item in sublist]
    grand_mean = np.mean(flat_data)
    
    day_means = [np.mean(day) for day in cleaned_matrix]
    day_vars = [np.var(day, ddof=1) for day in cleaned_matrix]
    
    ms_within = np.mean(day_vars) 
    ms_between = np.var(day_means, ddof=1) * n_rep
    
    s_r = np.sqrt(ms_within)
    v_b = max(0, (ms_between - ms_within) / n_rep)
    s_l = np.sqrt(v_b + ms_within)
    
    # 3. Tính UVL và VI (giữ nguyên logic trước)
    uvl_l = claim_sl * 1.32 
    se_x_bar = np.sqrt((1/n_run) * (s_l**2 - (1 - 1/n_rep) * s_r**2))
    t_val = 2.776 
    vi_half = t_val * se_x_bar
    vi_range = (target_mean - vi_half, target_mean + vi_half)
    
    return {
        "grand_mean": grand_mean,
        "s_r": s_r, "s_l": s_l,
        "uvl_l": uvl_l,
        "vi_range": vi_range,
        "is_precision_pass": s_l <= uvl_l,
        "is_trueness_pass": vi_range[0] <= grand_mean <= vi_range[1],
        "outliers": found_outliers # Trả về danh sách ngoại lệ để hiển thị
    }
# --- CƠ SỞ DỮ LIỆU TRA CỨU TIÊU CHUẨN CLIA & BIOLOGICAL VARIATION ---
STANDARD_DB = {
    "Glucose": {"tea": 8.0, "cvi": 5.6, "cvg": 7.8, "unit": "mg/dL"},
    "Albumin": {"tea": 8.0, "cvi": 3.1, "cvg": 4.2, "unit": "g/dL"},
    "Creatinine": {"tea": 10.0, "cvi": 5.9, "cvg": 14.7, "unit": "mg/dL"},
    "ALT": {"tea": 15.0, "cvi": 19.4, "cvg": 27.6, "unit": "U/L"},
    "AST": {"tea": 15.0, "cvi": 12.3, "cvg": 18.2, "unit": "U/L"},
    "Cholesterol": {"tea": 10.0, "cvi": 6.0, "cvg": 15.2, "unit": "mg/dL"},
    "HbA1c": {"tea": 6.0, "cvi": 1.2, "cvg": 4.0, "unit": "%"},
    "Bilirubin Total": {"tea": 20.0, "cvi": 21.8, "cvg": 31.2, "unit": "mg/dL"}
}
def generate_excel_report_comprehensive(test_info, df_full_iqc, df_eqa, mu_data, sigma_data, img_lj, img_sigma, img_vmask, report_period, mau_limits):
    import xlsxwriter
    import pandas as pd
    import io
    from datetime import datetime
    import numpy as np

    # --- 0. TIỀN XỬ LÝ DỮ LIỆU (QUAN TRỌNG NHẤT) ---
    # Ép buộc tính toán lỗi Westgard ngay tại đây để có cột 'Violation'
    if df_full_iqc is not None and not df_full_iqc.empty:
        # Gọi hàm tính lỗi (Đảm bảo hàm get_westgard_violations đã có trong main)
        df_full_iqc = get_westgard_violations(df_full_iqc, mu_data, sigma_data)
        df_final = df_full_iqc.sort_values(['date', 'level'])
        
    else:
        df_final = pd.DataFrame()

    m_min, m_des, m_opt = mau_limits
    start_date, end_date = report_period
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True, 'nan_inf_to_errors': True})

    # --- 1. ĐỊNH DẠNG (FORMATS) ---
    fmt_head = wb.add_format({'bold': True, 'align': 'center', 'bg_color': '#DDEBF7', 'border': 1, 'valign': 'vcenter', 'text_wrap': True})
    fmt_cell = wb.add_format({'align': 'center', 'border': 1, 'valign': 'vcenter'})
    fmt_num = wb.add_format({'num_format': '0.00', 'align': 'center', 'border': 1})
    
    # Định dạng lỗi
    fmt_err = wb.add_format({'color': 'white', 'bg_color': '#FF0000', 'bold': True, 'align': 'center', 'border': 1}) # Đỏ
    fmt_warn = wb.add_format({'bg_color': '#FFFF00', 'color': 'black', 'bold': True, 'align': 'center', 'border': 1}) # Vàng
    fmt_pass = wb.add_format({'bold': True, 'align': 'center', 'border': 1, 'color': '#008000'}) # Chữ xanh cho Đạt
    
    fmt_note = wb.add_format({'italic': True, 'bold': True, 'color': '#C00000', 'border': 1, 'valign': 'vcenter', 'text_wrap': True})
    fmt_sig_label = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
    fmt_sig_sub = wb.add_format({'italic': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 10})

    # --- CÁC HÀM PHỤ TRỢ ---
    def add_admin_section(ws, title, last_col_letter):
        ws.merge_range(f'A1:{last_col_letter}1', title, fmt_head)
        ws.write('A3', "Đơn vị:", fmt_head)
        ws.merge_range('B3:D3', "PHÒNG KHÁM ĐA KHOA QUỐC TẾ YERSIN", fmt_cell)
        ws.write('E3', "Xét nghiệm:", fmt_head)
        ws.merge_range(f'F3:{last_col_letter}3', test_info.get('name', 'N/A'), fmt_cell)
        ws.write('A4', "Khoa:", fmt_head)
        ws.merge_range('B4:D4', "XÉT NGHIỆM", fmt_cell)
        ws.write('E4', "Tháng :", fmt_head)
        ws.merge_range(f'F4:{last_col_letter}4', datetime.now().strftime("%m/%Y"), fmt_cell)
        ws.write('A5', "Thời gian:", fmt_head)
        ws.merge_range('B5:D5', f"{start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}", fmt_cell)
        ws.write('E5', "Thiết bị:", fmt_head)
        ws.merge_range(f'F5:{last_col_letter}5', test_info.get('device', 'N/A'), fmt_cell)

    def add_signature_section(ws, row_start, last_col_index):
        sig_r = row_start + 3
        ws.merge_range(sig_r, 0, sig_r, 2, "NGƯỜI LẬP BÁO CÁO", fmt_sig_label)
        ws.merge_range(sig_r + 1, 0, sig_r + 1, 2, "(Ký và ghi rõ họ tên)", fmt_sig_sub)
        ws.merge_range(sig_r, last_col_index - 2, sig_r, last_col_index, "TRƯỜNG KHOA XÉT NGHIỆM", fmt_sig_label)
        ws.merge_range(sig_r + 1, last_col_index - 2, sig_r + 1, last_col_index, "(Ký và ghi rõ họ tên)", fmt_sig_sub)

# === SHEET 1: NỘI KIỂM & TỔNG HỢP ===
    ws1 = wb.add_worksheet("Nội Kiểm & Tổng Hợp")
    COL_LAST_W1 = 'G'
    ws1.set_column('A:E', 12)
    ws1.set_column('F:F', 20)
    ws1.set_column('G:G', 45) 
    
    title = f"BÁO CÁO QUẢN LÝ CHẤT LƯỢNG: {test_info.get('name', 'N/A').upper()}"
    add_admin_section(ws1, title, COL_LAST_W1)

    # 2. BẢNG SIX SIGMA 
    curr_row = 7
    ws1.merge_range(f'A{curr_row}:G{curr_row}', "I. SIX SIGMA & HIỆU NĂNG PHƯƠNG PHÁP", fmt_head)
    headers_sigma = ["Level", "Mean", "CV%", "Bias%", "Sigma", "QGI", "Ghi chú"]
    ws1.write_row(curr_row, 0, headers_sigma, fmt_head)
    curr_row += 1

    if sigma_data:
        for lvl in sorted(sigma_data.keys()):
            res = sigma_data[lvl]
            ws1.write(curr_row, 0, f"Level {lvl}", fmt_cell)
            ws1.write(curr_row, 1, res.get('mean', 0), fmt_num)
            ws1.write(curr_row, 2, res.get('cv', 0), fmt_num)
            ws1.write(curr_row, 3, res.get('bias', 0), fmt_num)
            ws1.write(curr_row, 4, res.get('sigma', 0), fmt_num)
            ws1.write(curr_row, 5, res.get('qgi', 0), fmt_num)
            ws1.write(curr_row, 6, "Đạt" if res.get('sigma', 0) >= 3 else "Cần cải thiện", fmt_cell)
            curr_row += 1

    # 3. CHI TIẾT DỮ LIỆU IQC
    curr_row += 1
    ws1.merge_range(curr_row, 0, curr_row, 6, "II. CHI TIẾT DỮ LIỆU NỘI KIỂM (IQC) & VI PHẠM WESTGARD", fmt_head)
    headers_iqc = ["Ngày", "Lot", "Level", "Kết quả", "Z-Score", "Đánh giá (Lỗi)", "Hành động khắc phục"]
    ws1.write_row(curr_row + 1, 0, headers_iqc, fmt_head)

    r = curr_row + 2 

    if not df_final.empty:
        # Sắp xếp theo thời gian để các lỗi chuỗi (Across-level) hiển thị logic
        df_export = df_final.sort_values(by=['date', 'level']).copy()
        
        for _, item in df_export.iterrows():
            # 1. Thông tin cơ bản
            dt_val = pd.to_datetime(item['date'])
            ws1.write(r, 0, dt_val.strftime('%d/%m/%Y %H:%M'), fmt_cell)
            ws1.write(r, 1, str(item.get('lot_number', 'N/A')), fmt_cell)
            ws1.write(r, 2, item.get('level', 'N/A'), fmt_cell)
            ws1.write(r, 3, item.get('value', 0), fmt_num)
            
            # 2. Tính toán Z-Score hiển thị (Hỗ trợ 3 Level từ mu_data/sigma_data dạng dict)
            lvl = item.get('level')
            # Lấy Mean và SD tương ứng với từng Level
            if isinstance(mu_data, dict):
                m_t = mu_data.get(lvl, 0)
            else:
                m_t = mu_data # Trường hợp fallback nếu không phải dict
                
            if isinstance(sigma_data, dict):
                s_t = sigma_data.get(lvl, 0)
                if isinstance(s_t, dict): s_t = s_t.get('sd', 0)
            else:
                s_t = sigma_data
            
            z = (item['value'] - m_t) / s_t if s_t > 0 else 0
            ws1.write(r, 4, z, fmt_num)
            
            # 3. ĐÁNH GIÁ LỖI (Bổ sung quy tắc mới 6X, 9X, 12X)
            note_content = str(item.get('note', '')).upper()
            vio_raw = str(item.get('Violation', item.get('violation', ''))).upper()
            
            error_label = "ĐẠT"
            f_style = fmt_pass
            
            # Danh sách quy tắc bao gồm cả quy tắc Across-level mới
            rules = ["1-3S", "2-2S", "R-4S", "4-1S", "10X", "12X", "9X", "6X", "1-2S", "SHIFT", "TREND"]
            found_rule = None
            for rule in rules:
                if rule in note_content or rule in vio_raw:
                    found_rule = rule
                    break
            
            if found_rule:
                error_label = found_rule
                # Phân loại màu: Vàng cho 1-2S, Đỏ cho các lỗi còn lại (bao gồm 6x, 9x, 12x)
                if found_rule == "1-2S":
                    f_style = fmt_warn
                else:
                    f_style = fmt_err
            
            ws1.write(r, 5, error_label, f_style)
                
            # 4. Ghi Hành động khắc phục (Lọc sạch từ khóa rác)
            note_raw = str(item.get('note', '')).strip()
            action_raw = str(item.get('action', '')).strip()
            
            blacklist = ["nan", "none", "", "đạt", "ok", "nhập tay", "import", "au640"]
            
            final_parts = []
            # Loại bỏ phần tên lỗi khỏi nội dung ghi chú
            clean_note = note_raw
            for rule in rules:
                # Xử lý xóa cả "Across-level" nếu có trong text
                clean_note = clean_note.replace(f"Cảnh báo {rule}", "").replace(f"Vi phạm {rule}", "").replace(rule, "")
            
            clean_note = clean_note.replace("ACROSS-LEVEL", "").strip(". ").strip()

            if clean_note.lower() not in blacklist and not any(word in clean_note.lower() for word in ["nhập tay", "import"]):
                final_parts.append(clean_note)
            if action_raw.lower() not in blacklist and not any(word in action_raw.lower() for word in ["nhập tay", "import"]):
                final_parts.append(action_raw)
            
            display_note = " | ".join(final_parts)
            ws1.write(r, 6, display_note if display_note else " ", fmt_note if display_note else fmt_cell)
            r += 1
            
        curr_row = r
    else:
        ws1.merge_range(r, 0, r, 6, "Không có dữ liệu", fmt_cell)
        curr_row = r + 1
    # Chèn biểu đồ LJ
    if img_lj is not None:
        try:
            ws1.insert_image('I12', 'lj.png', {'image_data': io.BytesIO(img_lj), 'x_scale': 0.8, 'y_scale': 0.8})
        except: pass
    
    add_signature_section(ws1, curr_row + 2, 6)
    


    # === SHEET 2: NGOẠI KIỂM (EQA) ===
    ws2 = wb.add_worksheet("Ngoại Kiểm (EQA)")
    ws2.set_column('A:H', 12)
    add_admin_section(ws2, "KẾT QUẢ NGOẠI KIỂM & VMASK CUSUM", 'H')
    ws2.write_row('A7', ["Ngày", "Mã Mẫu", "PXN", "Ref", "SD Nhóm", "SDi", "CUSUM", "Đánh giá"], fmt_head)

    r2 = 7 # Bắt đầu từ dòng tiêu đề đã ghi
    if df_eqa is not None and not df_eqa.empty:
        df_eqa_s = df_eqa.sort_values('date').copy()
        for _, row in df_eqa_s.iterrows():
            r2 += 1
            ws2.write(r2, 0, pd.to_datetime(row['date']).strftime('%d/%m/%Y'), fmt_cell)
            ws2.write(r2, 1, str(row.get('sample_id', '')), fmt_cell)
            ws2.write(r2, 2, row.get('lab_value', 0), fmt_num)
            ws2.write(r2, 3, row.get('ref_value', 0), fmt_num)
            ws2.write(r2, 4, row.get('sd_group', 1), fmt_num)
            
            sdi = (row['lab_value'] - row['ref_value']) / row['sd_group'] if row.get('sd_group', 0) > 0 else 0
            ws2.write(r2, 5, sdi, fmt_num)
            ws2.write(r2, 6, row.get('CUSUM', 0), fmt_num)
            ws2.write(r2, 7, "Đạt" if abs(sdi) <= 2 else "Cần xem xét", fmt_cell)
            
    if img_vmask is not None:
        ws2.insert_image('A23', 'vmask.png', {'image_data': io.BytesIO(img_vmask), 'x_scale': 0.8, 'y_scale': 0.8})
        r2 += 30
    
    add_signature_section(ws2, r2 + 2, 7)

# === SHEET 3: MU & SIX SIGMA (TỐI ƯU HIỂN THỊ) ===
    ws3 = wb.add_worksheet("MU & SixSigma")
    
    # Mở rộng cột để chứa nội dung đánh giá và số liệu
    ws3.set_column('A:A', 15) # Level
    ws3.set_column('B:E', 12) # Mean, CV, Bias, Sigma
    ws3.set_column('F:G', 15) # Ue (đơn vị), Ue (%)
    ws3.set_column('H:I', 20) # Đánh giá BV
    
    add_admin_section(ws3, f"BÁO CÁO ĐỘ KHÔNG ĐẢM BẢO ĐO (MU) & SIGMA", 'H')
    
    # Tiêu đề bảng: Thêm cột để hiển thị rõ các thành phần MU
    headers_mu_sigma = ['Level', 'Mean', 'CV%', 'Bias%', 'Sigma', 'Ue (Giá trị)', 'Ue (%)', 'Đánh giá MU']
    ws3.write_row('A8', headers_mu_sigma, fmt_head)
    
    r3 = 8
    # Duyệt qua danh sách Level (1 và 2)
    for lvl in sorted(sigma_data.keys()):
        res_s = sigma_data.get(lvl, {})
        # Đảm bảo lấy đúng dữ liệu MU đã tính toán từ Tab 4
        res_m = mu_data.get(lvl, {}) 
        
        mean_v = res_s.get('mean', 0)
        # Ưu tiên lấy Ue (%) từ kết quả MU, nếu không có mới lấy từ Sigma Data
        ue_pct = res_m.get('ue', res_s.get('cv', 0) * 2) 
        
        # 1. ĐỊNH DẠNG MÀU SIGMA (Giữ nguyên logic thông minh của bạn)
        sig_val = res_s.get('sigma', 0)
        if sig_val >= 6: sig_color = '#b3e6ff'   # World Class
        elif sig_val >= 3: sig_color = '#c6efce' # Đạt
        else: sig_color = '#ffc7ce'             # Kém
        fmt_sigma_dynamic = wb.add_format({'bg_color': sig_color, 'border': 1, 'align': 'center', 'num_format': '0.00', 'bold': True})

        # 2. ĐỊNH DẠNG MÀU ĐÁNH GIÁ MU (Theo mục tiêu Biological Variation)
        # Sử dụng các giá trị m_opt, m_des, m_min truyền vào từ mau_limits
        if ue_pct <= 0: 
            stt = "N/A"; mu_col = '#FFFFFF'
        elif ue_pct <= (m_opt or 0): 
            stt = "🌟 Tối ưu"; mu_col = '#b3e6ff'
        elif ue_pct <= (m_des or 0): 
            stt = "✅ Mong muốn"; mu_col = '#c6efce'
        elif ue_pct <= (m_min or 0): 
            stt = "⚠️ Tối thiểu"; mu_col = '#fff2cc'
        else: 
            stt = "❌ Không đạt"; mu_col = '#ffc7ce'
            
        fmt_mu_status = wb.add_format({'bg_color': mu_col, 'border': 1, 'align': 'center', 'bold': True})

        # 3. GHI DỮ LIỆU XUỐNG DÒNG
        ws3.write(r3, 0, f"Level {lvl}", fmt_cell)
        ws3.write(r3, 1, mean_v, fmt_num)
        ws3.write(r3, 2, res_s.get('cv', 0), fmt_num)
        ws3.write(r3, 3, res_s.get('bias', 0), fmt_num)
        ws3.write(r3, 4, sig_val, fmt_sigma_dynamic) 
        
        # Ue tuyệt đối = (Ue% / 100) * Mean
        ue_absolute = (ue_pct / 100) * mean_v if mean_v > 0 else 0
        ws3.write(r3, 5, ue_absolute, fmt_num)
        ws3.write(r3, 6, ue_pct, fmt_num)
        ws3.write(r3, 7, stt, fmt_mu_status)
        r3 += 1

    # BẢNG THAM CHIẾU MỤC TIÊU MAU (Cập nhật tiêu chuẩn BV)
    rt = r3 + 2
    ws3.merge_range(rt, 0, rt, 3, "MỤC TIÊU ĐỘ KHÔNG ĐẢM BẢO ĐO CHO PHÉP (MAU)", fmt_head)
    ws3.write_row(rt + 1, 0, ["Mức độ (BV)", "Hệ số", "Giới hạn (%)", "Trạng thái"], fmt_head)
    ws3.write(rt + 2, 0, "Tối ưu", fmt_cell);    ws3.write(rt + 2, 1, "0.25", fmt_cell); ws3.write(rt + 2, 2, m_opt, fmt_num); ws3.write(rt + 2, 3, "Rất tốt", fmt_cell)
    ws3.write(rt + 3, 0, "Mong muốn", fmt_cell); ws3.write(rt + 3, 1, "0.50", fmt_cell); ws3.write(rt + 3, 2, m_des, fmt_num); ws3.write(rt + 3, 3, "Đạt", fmt_cell)
    ws3.write(rt + 4, 0, "Tối thiểu", fmt_cell); ws3.write(rt + 4, 1, "0.75", fmt_cell); ws3.write(rt + 4, 2, m_min, fmt_num); ws3.write(rt + 4, 3, "Chấp nhận", fmt_cell)

    # Chèn ảnh Sigma/Performance Map nếu có
    if img_sigma is not None:
        ws3.insert_image(rt + 6, 0, 'sigma.png', {'image_data': io.BytesIO(img_sigma), 'x_scale': 0.7, 'y_scale': 0.7})
            
    add_signature_section(ws3, rt + 25, 7)

    wb.close()
    return output.getvalue()
def export_verification_excel(test_name, standard_info, input_matrix, results):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        ws = workbook.add_worksheet('Báo cáo EP15-A3')
        
        # --- ĐỊNH DẠNG (FORMATTING) ---
        fmt_title = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'})
        fmt_header_table = workbook.add_format({'bold': True, 'bg_color': '#4F81BD', 'font_color': 'white', 'border': 1, 'align': 'center'})
        fmt_bold = workbook.add_format({'bold': True, 'border': 1})
        fmt_cell = workbook.add_format({'border': 1, 'align': 'center'})
        fmt_pass = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'bold': True, 'border': 1, 'align': 'center'})
        fmt_fail = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'bold': True, 'border': 1, 'align': 'center'})
        fmt_note = workbook.add_format({'italic': True, 'text_wrap': True, 'valign': 'top', 'border': 1})
        fmt_sig_label = workbook.add_format({'bold': True, 'align': 'center'})
        fmt_sig_sub = workbook.add_format({'italic': True, 'align': 'center'})

        ws.set_column('A:A', 22)
        ws.set_column('B:F', 12)
        ws.set_column('G:G', 18)

        # --- 1. PHẦN HÀNH CHÍNH ---
        last_col = 'G'
        ws.merge_range(f'A1:{last_col}1', f"BÁO CÁO XÁC NHẬN GIÁ TRỊ SỬ DỤNG - {test_name.upper()}", fmt_title)
        ws.write('A3', "Đơn vị:", fmt_bold); ws.merge_range('B3:D3', "PHÒNG KHÁM ĐA KHOA QUỐC TẾ YERSIN", fmt_cell)
        ws.write('E3', "Xét nghiệm:", fmt_bold); ws.merge_range(f'F3:{last_col}3', test_name, fmt_cell)
        ws.write('A4', "Khoa:", fmt_bold); ws.merge_range('B4:D4', "XÉT NGHIỆM", fmt_cell)
        ws.write('E4', "Thiết bị:", fmt_bold); ws.merge_range(f'F4:{last_col}4', "Hệ thống tự động", fmt_cell)

        # --- 2. THÔNG SỐ MỤC TIÊU ---
        ws.write('A6', "I. THÔNG SỐ MỤC TIÊU (MỤC TIÊU CHẤT LƯỢNG)", fmt_bold)
        ws.write_row('A7', ['Thông số', 'TEa (%)', 'CVi (%)', 'CVg (%)', 'Sl NSX (%)', 'Giá trị đích'], fmt_header_table)
        ws.write_row('A8', [
            test_name, 
            standard_info.get('tea', 0), standard_info.get('cvi', 0), standard_info.get('cvg', 0), 
            results.get('claim_sl', 0), results.get('target_mean', 0)
        ], fmt_cell)

        # --- 3. DỮ LIỆU THỰC NGHIỆM 5x5 ---
        ws.write('A10', "II. DỮ LIỆU THỰC NGHIỆM (5 NGÀY x 5 LẦN)", fmt_bold)
        ws.write_row('A11', ['Ngày', 'Lần 1', 'Lần 2', 'Lần 3', 'Lần 4', 'Lần 5', 'TB Ngày'], fmt_header_table)
        row = 11
        for i, day_data in enumerate(input_matrix):
            ws.write(row, 0, f"Ngày {i+1}", fmt_cell)
            ws.write_row(row, 1, day_data, fmt_cell)
            ws.write(row, 6, sum(day_data)/len(day_data), fmt_cell)
            row += 1
        # Ẩn cột H (cột dữ liệu phục vụ biểu đồ)
        ws.set_column('H:H', None, None, {'hidden': True})

        # --- 4. TẠO BIỂU ĐỒ ANOVA ---
        chart = workbook.add_chart({'type': 'line'})
        
        # Series 1: Trung bình ngày
        chart.add_series({
            'name':       'TB Ngày',
            'categories': ['Báo cáo EP15-A3', 11, 0, 15, 0],
            'values':     ['Báo cáo EP15-A3', 11, 6, 15, 6],
            'marker':     {'type': 'circle', 'size': 8, 'border': {'color': 'blue'}, 'fill': {'color': 'blue'}},
            'line':       {'color': '#4F81BD', 'width': 2},
        })
        
        # Series 2: Grand Mean (Đường thẳng tham chiếu)
        chart.add_series({
            'name':       'Grand Mean',
            'values':     ['Báo cáo EP15-A3', 11, 7, 15, 7],
            'line':       {'color': 'red', 'width': 1.5, 'dash_type': 'dash'},
        })

        chart.set_title({'name': f'Biến thiên trung bình ngày - {test_name}'})
        chart.set_x_axis({'name': 'Thời gian (Ngày)'})
        chart.set_y_axis({'name': 'Kết quả', 'major_gridlines': {'visible': True}})
        chart.set_legend({'position': 'bottom'})
        chart.set_size({'width': 450, 'height': 300})

        # Chèn biểu đồ vào bên phải bảng dữ liệu
        ws.insert_chart('I2', chart)
        # --- 4. KẾT QUẢ PHÂN TÍCH (EP15-A3) ---
        res_row = row + 1
        ws.merge_range(res_row, 0, res_row, 3, "III. PHÂN TÍCH THỐNG KÊ THEO CLSI EP15-A3", fmt_bold )
        
        ws.write(res_row+1, 0, "Chỉ số", fmt_header_table); ws.write(res_row+1, 1, "Thực tế", fmt_header_table)
        ws.write(res_row+1, 2, "Giới hạn (UVL/VI)", fmt_header_table); ws.write(res_row+1, 3, "Kết luận", fmt_header_table)

        # Độ chụm
        prec_pass = results.get('is_precision_pass')
        ws.write(res_row+2, 0, "Độ chụm Lab (Sl)", fmt_cell)
        ws.write(res_row+2, 1, f"{results.get('s_l', 0):.4f}", fmt_cell)
        ws.write(res_row+2, 2, f"< {results.get('uvl_l', 0):.4f}", fmt_cell)
        ws.write(res_row+2, 3, "ĐẠT" if prec_pass else "K.ĐẠT", fmt_pass if prec_pass else fmt_fail)

        # Độ đúng
        tru_pass = results.get('is_trueness_pass')
        vi = results.get('vi_range', (0,0))
        ws.write(res_row+3, 0, "Độ đúng (Mean)", fmt_cell)
        ws.write(res_row+3, 1, f"{results.get('grand_mean', 0):.4f}", fmt_cell)
        ws.write(res_row+3, 2, f"{vi[0]:.2f} - {vi[1]:.2f}", fmt_cell)
        ws.write(res_row+3, 3, "ĐẠT" if tru_pass else "K.ĐẠT", fmt_pass if tru_pass else fmt_fail)

        # --- 5. GHI CHÚ NGOẠI LỆ (TRÌNH BÀY MỚI) ---
        note_row = res_row + 5
        ws.write(note_row, 0, "IV. GHI CHÚ KIỂM TRA NGOẠI LỆ (GRUBBS' TEST)", fmt_bold)
        
        outliers = results.get('outliers', [])
        if not outliers:
            note_text = "Dữ liệu đạt kiểm tra Grubbs (Mức ý nghĩa alpha=0.05). Không phát hiện giá trị ngoại lệ trong 25 mẫu thử."
        else:
            note_text = "Phát hiện giá trị ngoại lệ đã được xử lý:\n"
            for out in outliers:
                note_text += f"- Ngày {out['day']}: Giá trị {out['value']} (Chỉ số G={out['g_score']:.2f} > 3.135)\n"
            note_text += "=> Các giá trị này đã được thay thế bằng trung bình ngày để đảm bảo tính ổn định của ANOVA."

        ws.merge_range(note_row + 1, 0, note_row + 3, 6, note_text, fmt_note)

        # --- 6. CHỮ KÝ ---
        sig_r = note_row + 5
        ws.merge_range(sig_r, 0, sig_r, 2, "NGƯỜI LẬP BÁO CÁO", fmt_sig_label)
        ws.merge_range(sig_r + 1, 0, sig_r + 1, 2, "(Ký và ghi rõ họ tên)", fmt_sig_sub)
        ws.merge_range(sig_r, 4, sig_r, 6, "TRƯỞNG KHOA XÉT NGHIỆM", fmt_sig_label)
        ws.merge_range(sig_r + 1, 4, sig_r + 1, 6, "(Ký và ghi rõ họ tên)", fmt_sig_sub)

    return output.getvalue()

# --- SIDEBAR: CONTROL PANEL ---

st.sidebar.markdown("---")
st.title("🏥 Hệ Thống Quản Lý Chất Lượng Xét Nghiệm ")
st.sidebar.title("Phòng Khám Đa Khoa Quốc Tế Yersin")
st.sidebar.title("Phòng Xét Nghiệm ")
st.sidebar.title("🎛️ Control Panel")

# 1. Chọn Thiết bị
all_devices = db.get_all_devices()
device_filter = st.sidebar.selectbox("Lọc theo Thiết bị", ["-- Tất cả --"] + all_devices)

# 2. Chọn Test
tests_df = db.get_all_tests()
if device_filter != "-- Tất cả --": tests_df = tests_df[tests_df['device'] == device_filter]
tests_options = {r['name']: r for _, r in tests_df.iterrows()}

# Loại bỏ "➕ Thêm Test" khỏi selectbox và dời vào expander
selected_test_name = st.sidebar.selectbox("🔬 Chọn Xét Nghiệm", ["-- Chọn --"] + list(tests_options.keys()))
all_tests = db.get_all_tests()
if not all_tests.empty:
    # --- ĐÂY LÀ NƠI ĐẶT ĐOẠN CODE ĐÓ ---
    st.sidebar.header("Lựa chọn xét nghiệm")
    test_names = all_tests['name'].tolist()
    selected_name = st.sidebar.selectbox("Chọn Xét Nghiệm", test_names)
    
    # Lấy dữ liệu chi tiết của Test đang chọn (bao gồm cả cvi, cvg vừa sửa)
    current_test = all_tests[all_tests['name'] == selected_name].iloc[0].to_dict()
    # ----------------------------------
## --- QUẢN LÝ TEST (THÊM TEST MỚI) ---
with st.sidebar.expander("➕ Thêm Test Mới"):
    with st.form("new_test_form"):
        st.write("Nhập thông tin Test mới")
        n = st.text_input("Tên Test")
        u = st.text_input("Đơn vị")
        d = st.text_input("Thiết bị")
        tea = st.number_input("TEa%", value=10.0, format="%.2f")
        cvi = st.number_input("CVi", value=0.0)
        cvg = st.number_input("CVg", value=0.0)
        
        if st.form_submit_button("Lưu Test Mới"):
            if n and d:
                # SỬA TẠI ĐÂY: Truyền thêm cvi và cvg vào hàm
                db.add_test(n, u, tea, d, cvi, cvg) 
                
                st.success(f"Đã thêm Test '{n}' ({d}).")
                st.rerun()
            else:
                st.warning("Vui lòng nhập Tên Test và Thiết bị.")


if selected_test_name == "-- Chọn --":
    st.title("👋 Chào mừng đến với Phần mềm QLCL")
    st.info("Vui lòng chọn một xét nghiệm từ menu bên trái để bắt đầu.")
    st.stop()

current_test = tests_options[selected_test_name]


# --- THAO TÁC SỬA/XÓA TEST ĐÃ CHỌN ---
st.sidebar.markdown("---")
st.sidebar.subheader("✏️ Thao tác Test Đã Chọn")

# 1. Nút Sửa Test
with st.sidebar.expander(f"⚙️ Sửa Test: {current_test['name']}"):
    with st.form("edit_test_form"):
        st.write("Chỉnh sửa thông tin Test/Thiết bị")
        
        # Nhập liệu thông tin cũ
        n_e = st.text_input("Tên Test", value=current_test['name'])
        u_e = st.text_input("Đơn vị", value=current_test['unit'])
        d_e = st.text_input("Thiết bị", value=current_test['device'])
        tea_e = st.number_input("TEa%", value=float(current_test['tea']), format="%.2f")
        
        # --- THÊM MỚI CVi, CVg ---
        # Sử dụng float() để đảm bảo kiểu dữ liệu đồng nhất
        cvi_e = st.number_input("CVi (Biến thiên sinh học trong cá thể)", 
                                value=float(current_test.get('CVi', 0.0)), format="%.2f")
        cvg_e = st.number_input("CVg (Biến thiên sinh học giữa các cá thể)", 
                                value=float(current_test.get('CVg', 0.0)), format="%.2f")
        
        if st.form_submit_button("Lưu Thay Đổi"):
            # Gọi hàm cập nhật vào database
            # Lưu ý: truyền thêm cvi_e và cvg_e vào hàm
            success = db.update_test(
                current_test['id'], 
                n_e, u_e, d_e, tea_e, cvi_e, cvg_e
            )
            
            if success:
                st.success("Đã cập nhật CVi, CVg thành công!")
                st.rerun()
            else:
                st.error("Lỗi khi lưu dữ liệu.")

# 2. Nút Xóa Test
with st.sidebar.expander("🗑️ Xóa Test (NGUY HIỂM)"):
    st.warning(f"Thao tác này sẽ xóa **Test {current_test['name']}** và **TẤT CẢ** dữ liệu IQC/EQA liên quan (Lot, Kết quả).")
    delete_confirm = st.checkbox(f"Tôi xác nhận muốn xóa Test **{current_test['name']}**", key="delete_test_confirm")
    
    if delete_confirm and st.button(f"THỰC HIỆN XÓA TEST", type="primary"):
        db.delete_test(current_test['id'])
        st.success("Đã xóa Test và dữ liệu liên quan."); st.rerun()
# 3. QUẢN LÝ LOTS (CẬP NHẬT: Thêm Chỉnh sửa & Xóa)
st.sidebar.markdown("---")
st.sidebar.subheader("📦 Cấu hình Lot Đang Chạy")

# Lấy dữ liệu và phân loại
all_lots = db.get_lots_for_test(current_test['id'])
lots_l1 = all_lots[all_lots['level'] == 1]
lots_l2 = all_lots[all_lots['level'] == 2]
lots_l3 = all_lots[all_lots['level'] == 3]

# Tạo dict để selectbox
opts_l1 = {f"{r['lot_number']} (Hạn:{r['expiry_date']})": r.to_dict() for _, r in lots_l1.iterrows()}
opts_l2 = {f"{r['lot_number']} (Hạn:{r['expiry_date']})": r.to_dict() for _, r in lots_l2.iterrows()}
opts_l3 = {f"{r['lot_number']} (Hạn:{r['expiry_date']})": r.to_dict() for _, r in lots_l3.iterrows()}

# --- SELECTBOX CHỌN LOT ĐANG CHẠY ---
s_l1 = st.sidebar.selectbox("Lot Level 1:", ["-- Chọn L1 --"] + list(opts_l1.keys()))
s_l2 = st.sidebar.selectbox("Lot Level 2:", ["-- Chọn L2 --"] + list(opts_l2.keys()))
s_l3 = st.sidebar.selectbox("Lot Level 3:", ["-- Chọn L3 --"] + list(opts_l3.keys()))

cur_lot_l1 = opts_l1[s_l1] if s_l1 != "-- Chọn L1 --" else None
cur_lot_l2 = opts_l2[s_l2] if s_l2 != "-- Chọn L2 --" else None
cur_lot_l3 = opts_l3[s_l3] if s_l3 != "-- Chọn L3 --" else None

# --- KHU VỰC CHỈNH SỬA / XÓA LOT ---
with st.sidebar.expander("📝 Chỉnh sửa / Xóa Lot hiện có"):
    tab_edit_l1, tab_edit_l2, tab_edit_l3  = st.tabs(["L1", "L2", "L3"])
    
    # Xử lý cho Level 1
    with tab_edit_l1:
        if not lots_l1.empty:
            for _, r in lots_l1.iterrows():
                with st.form(f"edit_l1_{r['id']}"):
                    st.caption(f"Chỉnh sửa Lot: {r['lot_number']}")
                    e_num = st.text_input("Số Lot", value=r['lot_number'])
                    e_m = st.number_input("Mean", value=float(r['mean']), format="%.3f")
                    e_sd = st.number_input("SD", value=float(r['sd']), format="%.3f")
                    e_exp = st.date_input("Hạn dùng", value=pd.to_datetime(r['expiry_date']))
                    
                    c1, c2 = st.columns(2)
                    if c1.form_submit_button("💾 Lưu"):
                        db.update_lot(r['id'], e_num, e_m, e_sd, e_exp.strftime('%Y-%m-%d'))
                        st.success("Đã cập nhật!"); time.sleep(0.5); st.rerun()
                    
                    if c2.form_submit_button("🗑️ Xóa"):
                        db.delete_lot(r['id'])
                        st.warning("Đã xóa Lot!"); time.sleep(0.5); st.rerun()
        else:
            st.write("Chưa có Lot L1")

    # Xử lý cho Level 2
    with tab_edit_l2:
        if not lots_l2.empty:
            for _, r in lots_l2.iterrows():
                with st.form(f"edit_l2_{r['id']}"):
                    st.caption(f"Chỉnh sửa Lot: {r['lot_number']}")
                    e_num = st.text_input("Số Lot", value=r['lot_number'])
                    e_m = st.number_input("Mean", value=float(r['mean']), format="%.3f")
                    e_sd = st.number_input("SD", value=float(r['sd']), format="%.3f")
                    e_exp = st.date_input("Hạn dùng", value=pd.to_datetime(r['expiry_date']))
                    
                    c1, c2 = st.columns(2)
                    if c1.form_submit_button("💾 Lưu"):
                        db.update_lot(r['id'], e_num, e_m, e_sd, e_exp.strftime('%Y-%m-%d'))
                        st.success("Đã cập nhật!"); time.sleep(0.5); st.rerun()
                    
                    if c2.form_submit_button("🗑️ Xóa"):
                        db.delete_lot(r['id'])
                        st.warning("Đã xóa Lot!"); time.sleep(0.5); st.rerun()
        else:
            st.write("Chưa có Lot L2") 
    # Xử lý cho Level 3
    with tab_edit_l3:
        if not lots_l3.empty:
            for _, r in lots_l3.iterrows():
                with st.form(f"edit_l3_{r['id']}"):
                    st.caption(f"Chỉnh sửa Lot: {r['lot_number']}")
                    e_num = st.text_input("Số Lot", value=r['lot_number'])
                    e_m = st.number_input("Mean", value=float(r['mean']), format="%.3f")
                    e_sd = st.number_input("SD", value=float(r['sd']), format="%.3f")
                    e_exp = st.date_input("Hạn dùng", value=pd.to_datetime(r['expiry_date']))
                    
                    c1, c2 = st.columns(2)
                    if c1.form_submit_button("💾 Lưu"):
                        db.update_lot(r['id'], e_num, e_m, e_sd, e_exp.strftime('%Y-%m-%d'))
                        st.success("Đã cập nhật!"); time.sleep(0.5); st.rerun()
                    
                    if c2.form_submit_button("🗑️ Xóa"):
                        db.delete_lot(r['id'])
                        st.warning("Đã xóa Lot!"); time.sleep(0.5); st.rerun()
        else:
            st.write("Chưa có Lot L3")

# --- FORM THÊM LOT MỚI (GIỮ NGUYÊN) ---
with st.sidebar.expander("➕ Thêm Lot Mới (Tùy chọn)"):
    with st.form("add_lot_flex"):
        st.write("Nhập thông tin Lot mới")
        mt = st.text_input("Phương pháp/Máy", value=current_test['device'])
        
        c1, c2, c3 = st.columns(3)
        with c1: 
            st.caption("Level 1")
            ln1 = st.text_input("Lot L1"); m1 = st.number_input("Mean 1", format="%.3f", key="m1_new"); sd1 = st.number_input("SD 1", format="%.3f", key="sd1_new")
            ed1 = st.date_input("Hạn L1", key="ed1_new")
        with c2:
            st.caption("Level 2")
            ln2 = st.text_input("Lot L2"); m2 = st.number_input("Mean 2", format="%.3f", key="m2_new"); sd2 = st.number_input("SD 2", format="%.3f", key="sd2_new")
            ed2 = st.date_input("Hạn L2", key="ed2_new")
        with c3:
            st.caption("Level 3")
            ln3 = st.text_input("Lot L3"); m3 = st.number_input("Mean 3", format="%.3f"); sd3 = st.number_input("SD 3", format="%.3f")
            ed3 = st.date_input("Hạn L3")

        if st.form_submit_button("Lưu Lot Mới"):
            if ln1: db.add_lot(current_test['id'], ln1, 1, mt, ed1.strftime('%Y-%m-%d'), m1, sd1)
            if ln2: db.add_lot(current_test['id'], ln2, 2, mt, ed2.strftime('%Y-%m-%d'), m2, sd2)
            if ln3: db.add_lot(current_test['id'], ln3, 3, mt, ed3.strftime('%Y-%m-%d'), m3, sd3) 
            st.success("Đã lưu!"); time.sleep(0.5); st.rerun()

# --- PHẦN LIÊN HỆ & HỖ TRỢ (DÁN VÀO CUỐI SIDEBAR) ---
st.sidebar.markdown("---") # Đường kẻ phân cách
with st.sidebar.expander("📞 Thông tin Liên hệ & Hỗ trợ", expanded=False):
    st.markdown(f"""
    <div style="line-height: 1.6;">
        <h4 style="margin-bottom: 0;">QLCL Lab v1.0</h4>
        <p style="font-size: 0.9em; color: gray;">Phiên bản: 2025 </p>
        <hr style="margin: 10px 0;">
        <p><b>Nhà phát triển:</b> [ThS. Nguyễn Đình Thọ]</p>
        <p><b>Email:</b> <a href="mailto:support@lab.com">dinhtho32@gmail.com</a></p>
        <p><b>Hotline:</b> <a href="tel:08 7678 1818">08 7678 1818</a></p>
        <p style="font-style: italic; font-size: 0.8em; margin-top: 10px;">
            Vui lòng liên hệ để được hỗ trợ kỹ thuật, nâng cấp hoặc tùy chỉnh báo cáo ISO 15189.
        </p>
    </div>
    """, unsafe_allow_html=True)

# Nút gửi nhanh yêu cầu hỗ trợ qua Email (Tùy chọn)
if st.sidebar.button("📧 Gửi báo lỗi nhanh"):
    subject = f"Bao loi phan mem QLCL - Test: {current_test['name']}"
    body = "Mo ta loi: "
    st.sidebar.write(f"Nhấn để gửi: [Click tại đây](mailto:support@lab.com?subject={subject}&body={body})")
# --- PHẦN GIAO DIỆN CÀI ĐẶT HỆ THỐNG ---
st.sidebar.markdown("---")
st.sidebar.subheader("🛠 Quản trị hệ thống")

if st.sidebar.button("🔄 Cập nhật cấu trúc dữ liệu"):
    with st.spinner("Đang kiểm tra hệ thống..."):
        success, message = upgrade_database_structure()
        
        if success:
            st.sidebar.success(message)
            # Tự động load lại app để nhận diện cột mới ngay lập tức
            st.rerun() 
        else:
            st.sidebar.error(message)


# --- MAIN UI ---
st.title(f"📊 {current_test['name']} - {current_test['device']}")

tabs = st.tabs(["1. Nhập IQC", "2. Biểu đồ LJ", "3. Ngoại kiểm (EQA)", "4. Độ KĐB (MU)", "5. Six Sigma & Báo cáo", "6. Xác nhận giá trị sử dụng ", "7. Import dữ liệu", "8. Quản trị"])

# === TAB 1: NHẬP IQC & QUẢN LÝ MAPPING ===
with tabs[0]:
    c_in, c_dat = st.columns([1, 2])
    
# --- CỘT TRÁI: NHẬP LIỆU THỦ CÔNG ---
    with c_in:
        st.subheader("📝 Nhập Kết Quả Hàng Ngày")
        if not cur_lot_l1 and not cur_lot_l2:
            st.error("Vui lòng chọn ít nhất 1 Lot ở Sidebar để nhập liệu.")
        else:
            with st.form("iqc_entry", clear_on_submit=True):
                # Sử dụng ngày hiện tại làm mặc định
                d_in = st.date_input("Ngày chạy", datetime.now())
                note = st.text_input("Ghi chú")
                
                v1, v2, v3 = None, None, None
                if cur_lot_l1: 
                    st.markdown(f"**L1: {cur_lot_l1['lot_number']}** (Target: {cur_lot_l1['mean']})")
                    v1 = st.number_input("Kết quả L1", format="%.4f", key="val_l1", value=0.0)
                
                if cur_lot_l2:
                    st.markdown(f"**L2: {cur_lot_l2['lot_number']}** (Target: {cur_lot_l2['mean']})")
                    v2 = st.number_input("Kết quả L2", format="%.4f", key="val_l2", value=0.0)
                
                if cur_lot_l3:
                    st.markdown(f"**L3: {cur_lot_l3['lot_number']}** (Target: {cur_lot_l3['mean']})")
                    v3 = st.number_input("Kết quả L3", format="%.4f", key="val_l3", value=0.0)  

                if st.form_submit_button("💾 Lưu Kết Quả"):
                    saved = False
                    
                    # Chuyển ngày thành chuỗi để lưu đồng bộ
                    date_str = d_in.strftime('%Y-%m-%d')
                    
                    # Lưu Mức 1 (chỉ lưu nếu v1 > 0)
                    if cur_lot_l1 and v1 > 0: 
                        db.add_iqc_data(
                            lot_id=cur_lot_l1['id'], 
                            dt=date_str, 
                            level=1, 
                            value=v1, 
                            note=note if note else "Nhập tay"
                        )
                        saved = True
                        
                    # Lưu Mức 2 (chỉ lưu nếu v2 > 0)
                    if cur_lot_l2 and v2 > 0: 
                        db.add_iqc_data(
                            lot_id=cur_lot_l2['id'], 
                            dt=date_str, 
                            level=2, 
                            value=v2, 
                            note=note if note else "Nhập tay"
                        )
                        saved = True
                    # Lưu Mức 3 (chỉ lưu nếu v3 > 0)
                    if cur_lot_l3 and v3 > 0: 
                        db.add_iqc_data(
                            lot_id=cur_lot_l3['id'], 
                            dt=date_str, 
                            level=3, 
                            value=v3, 
                            note=note if note else "Nhập tay"
                        )
                        saved = True
                    
                    if saved:
                        st.success("✅ Đã lưu kết quả vào bảng iqc_results!")
                        st.rerun()
                    else:
                        st.warning("Vui lòng nhập kết quả trước khi nhấn lưu.")

# --- CỘT PHẢI: HIỂN THỊ LỊCH SỬ ---
with c_dat:
    st.subheader("📊 Lịch sử dữ liệu tổng hợp")
    
    for lvl, cur_lot in zip([1, 2, 3], [cur_lot_l1, cur_lot_l2, cur_lot_l3]):
        if cur_lot:
            # DI CHUYỂN CSS VÀO ĐÂY để biến 'lvl' có hiệu lực
            # Sửa CSS để nhận diện các nút có chứa tiền tố ID Lot
            st.markdown(f"""
                <style>
                div.stButton > button[key*="btn_save_lot"] {{
                    background-color: #28a745 !important;
                    color: white !important;
                }}
                div.stButton > button[key*="btn_del_lot"] {{
                    background-color: #dc3545 !important;
                    color: white !important;
                }}
                </style>
            """, unsafe_allow_html=True)

            st.markdown(f"**Kết quả Mức {lvl}** (Lot: `{cur_lot['lot_number']}`)")
            
            df_lvl = db.get_iqc_data_by_lot(cur_lot['id'])
            
            if not df_lvl.empty:
                df_lvl['date'] = pd.to_datetime(df_lvl['date'], errors='coerce')
                
                edited_df = st.data_editor(
                    df_lvl,
                    column_config={
                        "id": None, 
                        "date": st.column_config.DatetimeColumn("Thời gian chạy", format="DD/MM/YYYY HH:mm:ss"),
                        "value": st.column_config.NumberColumn("Kết quả", format="%.4f"),
                        "level": st.column_config.TextColumn("Mức", disabled=True),
                        "note": st.column_config.TextColumn("Ghi chú")
                    },
                    num_rows="dynamic",
                    key=f"editor_final_l{lvl}",
                    use_container_width=True
                )
                
                col_save, col_del = st.columns(2)
                
                with col_save:
                    if st.button(f"💾 Lưu chỉnh sửa {lvl}", key=f"btn_save_{lvl}", use_container_width=True):
                        state = st.session_state.get(f"editor_final_l{lvl}", {})
                        if state.get("edited_rows"):
                            for row_idx, changes in state["edited_rows"].items():
                                actual_id = int(df_lvl.iloc[int(row_idx)]['id'])
                                db.update_iqc_result(actual_id, changes)
                            st.success(f"✅ Đã cập nhật Mức {lvl}")
                            st.rerun()

                with col_del:
                    if st.button(f"🗑️ Lưu Xóa {lvl}", key=f"btn_del_{lvl}", use_container_width=True):
                        state = st.session_state.get(f"editor_final_l{lvl}", {})
                        deleted_indices = state.get("deleted_rows", [])
                        
                        if deleted_indices:
                            success_count = 0
                            for idx in deleted_indices:
                                try:
                                    actual_id = int(df_lvl.iloc[idx]['id'])
                                    if db.delete_iqc_result(actual_id):
                                        success_count += 1
                                except Exception as e:
                                    st.error(f"Lỗi truy xuất ID: {e}")
                            
                            if success_count > 0:
                                st.success(f"✅ Đã xóa {success_count} dòng Mức {lvl}")
                                st.rerun()
                        else:
                            st.warning("⚠️ Hãy chọn dòng (bấm đầu dòng), nhấn Delete trên bàn phím, rồi mới nhấn nút Xóa này.")
            else:
                st.info(f"Mức {lvl} chưa có dữ liệu.")


# === TAB 2: BIỂU ĐỒ LJ & NHẬT KÝ VI PHẠM ===
# === TAB 2: BIỂU ĐỒ LJ & NHẬT KÝ VI PHẠM ===
with tabs[1]:
    col_opt, col_chart = st.columns([1, 4])
    
    with col_opt:
        view_mode = st.radio("Chế độ xem:", ["Chỉ Lot đang chọn", "Toàn bộ lịch sử (Nối Lot)"])
        
        st.markdown("---")
        st.subheader("📅 Khoảng thời gian")

        time_options = ["1 Tuần", "1 Tháng", "2 Tháng", "3 Tháng", "Tùy chỉnh ngày"]
        selected_label = st.selectbox(
            "Xem dữ liệu trong:", 
            time_options, 
            index=0,
            key="chart_time_range_tab2"
        )

        # 1. Xử lý logic Ngày bắt đầu và Kết thúc an toàn
        # Sử dụng pd.Timestamp.now().floor('D') để lấy ngày hiện tại không kèm giờ phút giây lắt nhắt
        now = pd.Timestamp.now().floor('D')
        end_date = now.replace(hour=23, minute=59, second=59)
        
        if selected_label == "Tùy chỉnh ngày":
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                # Mặc định lùi lại 7 ngày từ hôm nay
                default_start = (now - pd.Timedelta(days=7)).date()
                custom_start = st.date_input("Từ ngày", value=default_start, format="DD/MM/YYYY")
            with col_d2:
                custom_end = st.date_input("Đến ngày", value=now.date(), format="DD/MM/YYYY")
            
            start_date = pd.Timestamp(custom_start).replace(hour=0, minute=0, second=0)
            end_date = pd.Timestamp(custom_end).replace(hour=23, minute=59, second=59)
            
            if start_date > end_date:
                st.error("⚠️ Ngày bắt đầu không được lớn hơn ngày kết thúc!")
        else:
            days_map = {"1 Tuần": 7, "1 Tháng": 30, "2 Tháng": 60, "3 Tháng": 90}
            start_date = (now - pd.Timedelta(days=days_map[selected_label])).replace(hour=0, minute=0, second=0)
        st.caption(f"📍 {start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}")


    with col_chart:
        # 2. Lấy dữ liệu từ DB
        if view_mode == "Chỉ Lot đang chọn":
            data_list = []
            for l in [cur_lot_l1, cur_lot_l2, cur_lot_l3]:
                if l:
                    df_tmp = db.get_iqc_data_by_lot(l['id'])
                    if df_tmp is not None and not df_tmp.empty:
                        data_list.append(df_tmp)
            df_plot = pd.concat(data_list) if data_list else pd.DataFrame()
        else:
            # Lấy dư dữ liệu một chút để đảm bảo không sót khi lọc
            months_needed = 4 if "3 Tháng" in selected_label else 2
            df_plot = db.get_iqc_data_continuous(current_test['id'], max_months=months_needed)

        if not df_plot.empty:
            # 3. CHUẨN HÓA NGÀY THÁNG (Bước quan trọng nhất)
            # Ép kiểu dữ liệu về datetime, ưu tiên hiểu ngày đứng trước (DD/MM/YYYY)
            df_plot['date'] = pd.to_datetime(df_plot['date'], dayfirst=True, errors='coerce')
            df_plot = df_plot.dropna(subset=['date'])
            
            # 4. LỌC DỮ LIỆU CHÍNH XÁC THEO TIMESTAMP
            mask = (df_plot['date'] >= start_date) & (df_plot['date'] <= end_date)
            df_plot = df_plot.loc[mask].sort_values('date')

            if not df_plot.empty:
                # Gán thông số Target cho 3 Level
                for lvl, lot in zip([1, 2, 3], [cur_lot_l1, cur_lot_l2, cur_lot_l3]):
                    if lot:
                        l_mask = df_plot['level'] == lvl
                        df_plot.loc[l_mask, 'target_mean'] = float(lot['mean'])
                        df_plot.loc[l_mask, 'target_sd'] = float(lot['sd'])
                        df_plot.loc[l_mask, 'lot_number'] = str(lot['lot_number'])

                # 5. VẼ BIỂU ĐỒ
                fig_lj = plot_levey_jennings(df_plot, f"Biểu đồ Levey-Jennings ({current_test['name']})")
                st.pyplot(fig_lj)
                
                # Lưu vào Session State
                st.session_state['fig_lj_report'] = fig_lj
            else:
                st.warning(f"Không tìm thấy dữ liệu trong khoảng từ {start_date.strftime('%d/%m/%Y')} đến {end_date.strftime('%d/%m/%Y')}. Hãy thử mở rộng khoảng thời gian hoặc kiểm tra lại định dạng ngày nhập liệu.")
        else:
            st.info("Chưa có dữ liệu nội kiểm trong hệ thống cho xét nghiệm này.")
# --- CẢNH BÁO WESTGARD NHANH ---
        st.markdown("#### ⚠️ Cảnh báo Westgard")
        violations = {}

        # KIỂM TRA AN TOÀN: Chỉ chạy nếu df_plot có dữ liệu và có cột 'level'
        if df_plot is not None and not df_plot.empty and 'level' in df_plot.columns:
            for lvl in [1, 2, 3]:
                lot = None
                if lvl == 1: lot = cur_lot_l1
                elif lvl == 2: lot = cur_lot_l2
                elif lvl == 3: lot = cur_lot_l3

                
                if lot:
                    sub = df_plot[df_plot['level'] == lvl].copy()
                    if not sub.empty:
                        # Sử dụng hàm get_westgard_violations đã tối ưu
                        analyzed = get_westgard_violations(sub, lot['mean'], lot['sd'])
                        
                        # Kiểm tra xem cột 'Violation' có tồn tại sau khi tính toán không
                        if 'Violation' in analyzed.columns:
                            current_v = analyzed['Violation'].iloc[-1]
                            violations[f"Mức {lvl}"] = current_v if (current_v and str(current_v).strip() != "") else "ĐẠT"
                        else:
                            violations[f"Mức {lvl}"] = "ĐẠT"
        
        # HIỂN THỊ KẾT QUẢ THEO MÀU SẮC
        if violations:
            for k, v in violations.items():
                status_upper = str(v).upper()
                
                # 1. Nếu ĐẠT hoặc không có lỗi: Hiện nền xanh (Success)
                if status_upper in ["ĐẠT", "PASS", "OK", "0", "NAN", "NONE"]:
                    st.success(f"**{k}**: ĐẠT")
                
                # 2. Nếu là Cảnh báo 1-2s: Hiện nền vàng (Warning)
                elif "1-2S" in status_upper:
                    st.warning(f"**{k}**: {v} (Cảnh báo - Theo dõi sát)")
                
                # 3. Nếu là Vi phạm quy tắc dừng (1-3s, 2-2s, R-4s...): Hiện nền đỏ (Error)
                else:
                    st.error(f"**{k}**: {v} (Vi phạm quy tắc dừng - Cần xử lý)")
        else:
            # Thông báo khi xét nghiệm mới tạo, chưa có dữ liệu để tính toán
            st.info("ℹ️ Hiện tại chưa có dữ liệu IQC để đánh giá Westgard cho xét nghiệm này.")

        st.divider()

        # --- NHẬT KÝ VI PHẠM (Sửa lỗi không lưu được khi nhập thủ công) ---
        for lvl_info in [{"id": 1, "lot": cur_lot_l1}, {"id": 2, "lot": cur_lot_l2}, {"id": 3, "lot": cur_lot_l3}]:
            lvl = lvl_info["id"]
            lot = lvl_info["lot"]
            if lot:
                unique_prefix = f"lot_{lot['id']}_lvl_{lvl}"
                df_raw = db.get_iqc_data_by_lot(lot['id'])
                
                if not df_raw.empty:
                    df_analyzed = get_westgard_violations(df_raw, lot['mean'], lot['sd'])
                    df_err = df_analyzed[~df_analyzed['Violation'].isin(["ĐẠT", "", "None", None])].copy()
                    
                    if not df_err.empty:
                        st.markdown(f"**📝 Nhật ký xử lý vi phạm Mức {lvl} ({lot['lot_number']})**")
                        
                        # Hiển thị bảng và nhận giá trị trả về ngay khi người dùng chỉnh sửa
                        edited_err = st.data_editor(
                            df_err[['id', 'date', 'value', 'level', 'Violation', 'note']].rename(columns={
                                'date': 'Thời điểm', 'value': 'Kết quả', 'level': 'Mức', 'Violation': 'Lỗi', 'note': 'Hành động khắc phục'
                            }),
                            column_config={
                                "id": None, "Mức": None,
                                "Hành động khắc phục": st.column_config.TextColumn(width="large")
                            },
                            disabled=["Thời điểm", "Kết quả", "Lỗi"],
                            key=f"editor_{unique_prefix}",
                            use_container_width=True,
                            hide_index=True
                        )
                        
                    if st.button(f"💾 Lưu xử lý Mức {lvl}", key=f"btn_save_{unique_prefix}"):
                        now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
                        success_count = 0
                        
                        SUGGESTIONS = {
                            "1-3s": "Vi phạm 1-3s. Kiểm tra bọt khí, kim hút, hóa chất. Calib hóa chất và chạy lại QC mới.",
                            "R-4s": "Vi phạm R-4s. Lỗi ngẫu nhiên. Kiểm tra độ đồng nhất và chạy lại.",
                            "2-2s": "Vi phạm 2-2s. Lỗi hệ thống. Kiểm tra hạn dùng hóa chất hoặc hiệu chuẩn lại.",
                            "4-1s": "Vi phạm 4-1s. Lỗi hệ thống nhỏ. Kiểm tra xu hướng trôi, xem xét hiệu chuẩn.",
                            "10x": "Vi phạm 10x. Lỗi hệ thống kéo dài. Kiểm tra bảo trì hoặc hiệu chuẩn lại.",
                            "Shift": "Lỗi hệ thống. Kiểm tra hóa chất/hiệu chuẩn.",
                            "Trend": "Lỗi hệ thống. Kiểm tra sự thoái hóa của hóa chất, bóng đèn.",
                            "1-2s": "Cảnh báo 1-2s. Theo dõi sát kết quả tiếp theo."
                        }

                        for _, row in edited_err.iterrows():
                            user_note = str(row['Hành động khắc phục']).strip()
                            v_type = str(row['Lỗi'])
                            
                            # Loại bỏ triệt để nội dung cũ
                            junk_words = ["nhập tay", "import", "au640", "none", "nan", ""]
                            is_junk = any(word in user_note.lower() for word in junk_words)
                            
                            if is_junk:
                                # Nếu là nội dung cũ hoặc trống -> Lấy gợi ý chuẩn
                                final_action = "Kiểm tra hệ thống theo quy trình chuẩn."
                                for k, msg in SUGGESTIONS.items():
                                    if k in v_type:
                                        final_action = msg
                                        break
                            else:
                                # Nếu người dùng đã gõ nội dung mới -> Giữ nguyên
                                final_action = user_note

                            # Thêm dấu thời gian
                            if " - [Xử lý lúc:" not in final_action:
                                final_action = f"{final_action} - [Xử lý lúc: {now_str}]"

                            # Gọi hàm đã sửa với thứ tự tham số mới: iqc_id, note, dt, level, value
                            if db.update_iqc_data(
                                iqc_id=int(row['id']),
                                note=final_action,
                                dt=row['Thời điểm'],
                                level=int(row['Mức']),
                                value=float(row['Kết quả'])
                            ):
                                success_count += 1

                        if success_count > 0:
                            st.success(f"✅ Đã lưu {success_count} dòng thành công!")
                            st.rerun()



# === TAB: NGOẠI KIỂM (EQA) & CUSUM ===
with tabs[2]:
    st.subheader("2. Ngoại Kiểm (EQA) & Biểu đồ CUSUM")

    # 1. LẤY DỮ LIỆU & TÍNH TOÁN
    df_eqa = db.get_eqa_data(current_test['id'])

    if not df_eqa.empty:
        df_eqa['date'] = pd.to_datetime(df_eqa['date']).dt.date
        df_eqa = df_eqa.sort_values(by='date').reset_index(drop=True)

        # Tính toán lại Z-Score và CUSUM (CUSUM cần tính trên df đã sắp xếp)
        df_eqa['Z-Score'] = (df_eqa['lab_value'] - df_eqa['ref_value']) / df_eqa['sd_group']
        df_eqa['%Bias'] = ((df_eqa['lab_value'] - df_eqa['ref_value']) / df_eqa['ref_value']) * 100
        df_eqa['CUSUM'] = df_eqa['Z-Score'].cumsum()
        
        # DataFrame hiển thị (sắp xếp mới nhất lên trên)
        df_display = df_eqa.sort_values(by='date', ascending=False).reset_index(drop=True)
    else:
        df_display = pd.DataFrame()

    # --- PHẦN 1: NHẬP LIỆU ---
    c1, c2 = st.columns([1, 2])
    
# --- PHẦN NHẬP KẾT QUẢ EQA (Bên cột c1) ---
    with c1:
        st.subheader("Nhập kết quả EQA")
        eqa_date = st.date_input("Ngày mẫu", value=datetime.now())
        eqa_pxn = st.number_input("Giá trị PXN", format="%.4f")
        eqa_ref = st.number_input("Giá trị Tham chiếu (Nhóm)", format="%.4f")
        eqa_sd = st.number_input("SD Nhóm (Group SD)", format="%.4f")
        eqa_code = st.text_input("Mã mẫu", value="Đợt 1")

        if st.button("Lưu EQA"):
            if eqa_sd > 0:
                # Tính toán SDI (Z-Score) trước khi lưu
                # Công thức: $sdi = \frac{lab\_value - ref\_value}{sd\_group}$
                sdi = (eqa_pxn - eqa_ref) / eqa_sd
                
                # Tạo dictionary dữ liệu
                data_to_save = {
                    'test_id': current_test['id'],
                    'date': eqa_date.strftime('%Y-%m-%d'),
                    'lab_value': eqa_pxn,
                    'ref_value': eqa_ref,
                    'sd_group': eqa_sd,
                    'sdi': sdi,
                    'program_name': eqa_code
                }
                
                if db.add_eqa(data_to_save):
                    st.success("✅ Đã lưu kết quả EQA!")
                    st.rerun() # Quan trọng để bảng bên phải cập nhật ngay
                else:
                    st.error("❌ Lỗi khi lưu vào cơ sở dữ liệu.")
            else:
                st.error("⚠️ SD Nhóm phải lớn hơn 0 để tính Z-Score.")

# --- PHẦN 2: BẢNG DỮ LIỆU CÓ CHỨC NĂNG CHỈNH SỬA & XÓA ---
    with c2:
        st.subheader("📊 Dữ liệu EQA")

 # --- PHẦN XỬ LÝ DỮ LIỆU HIỂN THỊ (Sau khi lấy df_eqa từ database) ---
        if not df_eqa.empty:
            df_display = df_eqa.copy()
            
            # 1. Ép kiểu dữ liệu số và xử lý None/NaN cho các cột tính toán
            # Điều này cực kỳ quan trọng để khắc phục lỗi 'None' trong hình của bạn
            numeric_cols = ['lab_value', 'ref_value', 'sd_group', 'sdi']
            for col in numeric_cols:
                if col in df_display.columns:
                    # Chuyển đổi sang số, các giá trị lỗi hoặc None sẽ thành NaN, sau đó điền 0
                    df_display[col] = pd.to_numeric(df_display[col], errors='coerce').fillna(0)
            
            # 2. Tính toán các giá trị phái sinh để hiển thị
            # SDI trong DB chính là Z-Score trên giao diện
            df_display['Z-Score'] = df_display['sdi']
            
            # Tính CUSUM dựa trên cột Z-Score vừa xử lý
            df_display['CUSUM'] = df_display['Z-Score'].cumsum()
            
            # 3. Chuẩn bị danh sách cột nguồn (Sử dụng program_name)
            source_cols = ['id', 'date', 'program_name', 'lab_value', 'ref_value', 'sd_group', 'Z-Score', 'CUSUM']
            actual_cols = [c for c in source_cols if c in df_display.columns]
            
            # Tạo bản sao cuối cùng để đưa vào Editor
            df_edit = df_display[actual_cols].copy()
            
            # Mapping tên cột Tiếng Việt
            column_mapping = {
                'id': 'ID',
                'date': 'Ngày',
                'program_name': 'Mã Mẫu',
                'lab_value': 'PXN',
                'ref_value': 'Ref',
                'sd_group': 'SD Nhóm',
                'Z-Score': 'Z-Score',
                'CUSUM': 'CUSUM'
            }
            
            new_names = [column_mapping[c] for c in actual_cols]
            df_edit.columns = new_names

            # Quan trọng: Đặt Index là ID trước khi chèn cột Xóa
            if 'ID' in df_edit.columns:
                df_edit = df_edit.set_index('ID')
            
            # --- GIẢI QUYẾT LỖI KEYERROR 'XÓA' ---
            # Phải chèn cột Xóa vào df_edit trước khi hiển thị trong data_editor
            if 'Xóa' not in df_edit.columns:
                df_edit.insert(0, 'Xóa', False)
            
            # 3. HIỂN THỊ data_editor
            edited_df = st.data_editor(
                df_edit,
                key="eqa_data_editor",
                column_config={
                    "PXN": st.column_config.NumberColumn(format="%.4f", required=True),
                    "Ref": st.column_config.NumberColumn(format="%.4f", required=True),
                    "SD Nhóm": st.column_config.NumberColumn(format="%.4f", required=True),
                    "Z-Score": st.column_config.NumberColumn(disabled=True, format="%.2f"),
                    "CUSUM": st.column_config.NumberColumn(disabled=True, format="%.2f"),
                    "Xóa": st.column_config.CheckboxColumn("Xóa", default=False)
                },
                hide_index=False,
                use_container_width=True,
            )

            # 4. XỬ LÝ HÀNH ĐỘNG (NÚT ÁP DỤNG)
            if st.button("Áp dụng thay đổi (Xóa/Sửa)"):
                # 1. Truy cập trực tiếp vào state của editor
                editor_state = st.session_state.get("eqa_data_editor", {})
                edits = editor_state.get("edited_rows", {})
                
                if not edits:
                    st.warning("⚠️ Hệ thống chưa ghi nhận thay đổi nào.")
                else:
                    deleted_count = 0
                    update_count = 0

                    for row_idx_str, changes in edits.items():
                        try:
                            # Lấy ID từ index của dòng dựa trên số thứ tự
                            row_num = int(row_idx_str)
                            # Ép kiểu ID về int để đảm bảo khớp với Database
                            actual_id = int(edited_df.index[row_num])
                            
                            # TRƯỜNG HỢP 1: XÓA
                            if changes.get('Xóa') == True:
                                if db.delete_eqa(actual_id):
                                    deleted_count += 1
                            
                            # TRƯỜNG HỢP 2: SỬA
                            else:
                                current_row = edited_df.loc[actual_id]
                                update_data = {}
                                
                                # Ánh xạ lại tên cột Database
                                if 'PXN' in changes: update_data['lab_value'] = changes['PXN']
                                if 'Ref' in changes: update_data['ref_value'] = changes['Ref']
                                if 'SD Nhóm' in changes: update_data['sd_group'] = changes['SD Nhóm']
                                if 'Mã Mẫu' in changes: update_data['program_name'] = changes['Mã Mẫu']
                                
                                # Tính toán lại SDI nếu có sửa số liệu
                                v_lab = update_data.get('lab_value', current_row['PXN'])
                                v_ref = update_data.get('ref_value', current_row['Ref'])
                                v_sd = update_data.get('sd_group', current_row['SD Nhóm'])
                                
                                if v_sd > 0:
                                    update_data['sdi'] = (v_lab - v_ref) / v_sd
                                
                                if update_data:
                                    if db.update_eqa(actual_id, update_data):
                                        update_count += 1
                        except Exception as e:
                            st.error(f"Lỗi tại dòng {row_idx_str}: {e}")

                    # THÔNG BÁO KẾT QUẢ
                    if deleted_count > 0 or update_count > 0:
                        st.success(f"✅ Thành công: Xóa {deleted_count} dòng, Cập nhật {update_count} dòng.")
                        # Xóa state cũ để tránh lặp lại hành động
                        st.rerun()

    # --- PHẦN 3: VẼ BIỂU ĐỒ CUSUM VỚI V-MASK ---
    st.markdown("---")
    
    # Sử dụng df_eqa đã sắp xếp và tính CUSUM ở bước 1
    if not df_eqa.empty and len(df_eqa) > 1:
        st.subheader(f"Biểu đồ CUSUM & V-Mask (Góc 28°, d=10)")
        
        dates = df_eqa['date']
        cusum_values = df_eqa['CUSUM'].values
        n_points = len(cusum_values)
        indices = np.arange(n_points)
        
        # --- TÍNH TOÁN V-MASK ---
        last_x = indices[-1]
        last_y = cusum_values[-1]
        theta_deg = 28
        d = 10
        k = np.tan(np.radians(theta_deg))
        vertex_x = last_x + d
        vertex_y = last_y
        
        x_mask = np.linspace(0, vertex_x, 100)
        y_upper = vertex_y + k * (vertex_x - x_mask)
        y_lower = vertex_y - k * (vertex_x - x_mask)
        
        # --- VẼ BIỂU ĐỒ ---
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.plot(indices, cusum_values, marker='o', linestyle='-', color='blue', label='CUSUM Line')
        
        mask_range_mask = x_mask >= 0 
        ax.plot(x_mask[mask_range_mask], y_upper[mask_range_mask], color='red', linestyle='--', alpha=0.7, label='V-Mask Upper')
        ax.plot(x_mask[mask_range_mask], y_lower[mask_range_mask], color='red', linestyle='--', alpha=0.7, label='V-Mask Lower')
        
        # Vẽ V-Mask 
        ax.plot(vertex_x, vertex_y, marker='x', color='black', markersize=10, label='Vertex (d=10)')
        ax.plot([last_x, vertex_x], [last_y, vertex_y], color='gray', linestyle=':', alpha=0.5)

        is_violated = False
        for i in range(n_points):
            limit_upper = vertex_y + k * (vertex_x - i)
            limit_lower = vertex_y - k * (vertex_x - i)
            
            if cusum_values[i] > limit_upper or cusum_values[i] < limit_lower:
                is_violated = True
                ax.scatter(i, cusum_values[i], color='orange', s=100, zorder=5)

        ax.axhline(0, color='black', linewidth=0.5)
        ax.set_title(f"Biểu đồ CUSUM (Mẫu cuối: {last_y:.2f})")
        ax.set_xlabel("Số thứ tự mẫu EQA")
        ax.set_ylabel("CUSUM (SDI tích lũy)")
        ax.legend()
        ax.grid(True, alpha=0.3)
        
        if n_points <= 10:
            ax.set_xticks(indices)
            ax.set_xticklabels([d.strftime('%d/%m') for d in dates], rotation=45)
        
        st.pyplot(fig)
        st.session_state['fig_vmask_report'] = fig

        if is_violated:
            st.error("⚠️ CẢNH BÁO: Đường CUSUM cắt V-Mask! Có dấu hiệu sai số hệ thống (Shift/Trend).")
        else:
            st.success("✅ Hệ thống ổn định (CUSUM nằm trong V-Mask).")
            
    elif not df_eqa.empty:
        st.warning("Cần ít nhất 2 điểm dữ liệu EQA để vẽ biểu đồ CUSUM.")

# === TAB 4: ĐỘ KĐB ĐO (MU) & QUẢN TRỊ CHẤT LƯỢNG ===
with tabs[3]:
    st.header("4. Độ Không Đảm Bảo Đo (MU) & Đánh giá Hiệu năng")

    if cur_lot_l1 is None and cur_lot_l2 is None:
        st.warning("⚠️ Vui lòng cấu hình Lot QC ở Sidebar để thực hiện tính toán MU.")
    else:
    # --- KIỂM TRA THỜI HẠN XEM XÉT (ISO 15189) ---
# --- KIỂM TRA THỜI HẠN XEM XÉT (ISO 15189) ---
        last_review = current_test.get('last_mu_review')

        if last_review is None or last_review == "":
            # Thụt lề 1 Tab (hoặc 4 dấu cách) cho các dòng bên trong IF
            st.warning("⚠️ Xét nghiệm này chưa có dữ liệu xem xét MU định kỳ.")
            last_review_dt = datetime.now().date() 
            last_review_display = "Chưa thiết lập"
        else:
            # Thụt lề 1 Tab cho các dòng bên trong ELSE
            try:
                last_review_dt = datetime.strptime(str(last_review), '%Y-%m-%d').date()
                last_review_display = last_review
            except ValueError:
                last_review_dt = datetime.now().date()
                last_review_display = "Định dạng sai"

        # Dòng này phải thẳng hàng với chữ IF/ELSE phía trên
        diff_days = (datetime.now().date() - last_review_dt).days

        if last_review is None:
            st.info("💡 Hãy thực hiện xem xét MU lần đầu cho xét nghiệm này.")
        elif diff_days > 180:
            st.error(f"🚨 Đã {diff_days} ngày chưa xem xét MU định kỳ (Yêu cầu: 6-12 tháng).")
        else:
            st.success(f"✅ Ngày xem xét gần nhất: {last_review_display} ({diff_days} ngày trước)")

        # --- CẤU HÌNH THÔNG SỐ ĐẦU VÀO ---
        with st.expander("⚙️ Cấu hình Mục tiêu MAU & Thành phần MU", expanded=True):
            c_cfg1, c_cfg2 = st.columns(2)
            
            with c_cfg1:
                st.subheader("1. Khoảng thời gian")
                col_t1, col_t2 = st.columns(2)
                d_start = col_t1.date_input("Từ ngày", datetime.now() - timedelta(days=90), key="mu_start")
                d_end = col_t2.date_input("Đến ngày", datetime.now(), key="mu_end")
                
                u_ref_pct = st.number_input("u_ref từ mẫu EQA (%)", value=1.5, step=0.1)
                clin_decision = st.number_input("Nồng độ chẩn đoán lâm sàng", value=0.0)

            with c_cfg2:
                st.subheader("2. Mục tiêu Biến thiên sinh học (BV)")
                cvi_in = st.number_input("CVi (Cá thể)", value=float(current_test.get('cvi', 0.0)), format="%.2f")
                cvg_in = st.number_input("CVg (Quần thể)", value=float(current_test.get('cvg', 0.0)), format="%.2f")
                
                if cvi_in > 0:
                    bv_combined = np.sqrt(cvi_in**2 + cvg_in**2)
                    mau_min = 0.75 * cvi_in + 1.65 * (0.375 * bv_combined)
                    mau_des = 0.5 * cvi_in + 1.65 * (0.25 * bv_combined)
                    mau_opt = 0.25 * cvi_in + 1.65 * (0.125 * bv_combined)
                    st.code(f"Tối ưu: {mau_opt:.2f}% | Mong muốn: {mau_des:.2f}% | Tối thiểu: {mau_min:.2f}%")
                    target_mau = mau_des 
                else:
                    target_mau = float(current_test.get('tea', 10.0))
                    st.warning(f"Sử dụng TEa cố định làm mục tiêu: {target_mau}%")

        # --- XỬ LÝ DỮ LIỆU ---
        df_iqc_raw = db.get_iqc_data_continuous(current_test['id'])
        df_eqa_hist = db.get_eqa_data(current_test['id'])

        # Tính Bias trung bình từ 3 kỳ EQA gần nhất
        bias_pct_val = 0.0
        if not df_eqa_hist.empty:
            recent_eqa = df_eqa_hist.tail(3).copy()
            recent_eqa['%Bias'] = abs((recent_eqa['lab_value'] - recent_eqa['ref_value'])/recent_eqa['ref_value'])*100
            bias_pct_val = recent_eqa['%Bias'].mean()

           # --- HIỂN THỊ KẾT QUẢ (CẬP NHẬT 3 LEVEL) ---
        st.markdown("---")
        # Chia thành 3 cột tương ứng với 3 mức QC
        c1, c2, c3 = st.columns(3)
        mu_results = {}
   
        # Danh sách các cột và các Lot đã chọn để lặp
        columns = [c1, c2, c3]
        current_lots = [cur_lot_l1, cur_lot_l2, cur_lot_l3]

        level_styles = {
            1: {"icon": "🔵", "color": "blue", "name": "Level 1"},
            2: {"icon": "🟠", "color": "orange", "name": "Level 2"},
            3: {"icon": "🔴", "color": "red", "name": "Level 3"}
        }

        for i, col in enumerate(columns, 1):
            style = level_styles[i]
            with col:
                # Hiển thị tiêu đề với màu sắc riêng biệt cho từng Level
                st.markdown(f"### {style['icon']} <span style='color:{style['color']}'>{style['name']}</span>", unsafe_allow_html=True)
                lot_info = current_lots[i-1]
   
                if lot_info:
                    # Lọc dữ liệu cho từng level
                    sub_df = df_plot[df_plot['level'] == i]
                
                if not df_iqc_raw.empty:
                    # Lọc theo Level và Ngày
                    df_temp = df_iqc_raw.copy()
                    df_temp['date'] = pd.to_datetime(df_temp['date']).dt.date
                    df_lvl = df_temp[(df_temp['level'] == i) & (df_temp['date'] >= d_start) & (df_temp['date'] <= d_end)]
                    
                    stats = get_clean_stats_3sigma(df_lvl)
                    
                    if stats:
                        u_prec = stats['cv']
                        # Công thức: uc = sqrt(u_prec^2 + u_bias^2 + u_ref^2)
                        uc = np.sqrt(u_prec**2 + bias_pct_val**2 + u_ref_pct**2)
                        ue = uc * 2 # Mở rộng k=2
                        
                        mu_results[i] = {
                            "ue": ue, "mean": stats['mean'], "u_prec": u_prec, 
                            "u_bias": bias_pct_val, "u_ref": u_ref_pct, "n_count": stats['n']
                        }

                        # Đánh giá màu sắc
                        if ue <= (mau_opt if cvi_in > 0 else target_mau): status, color = "🌟 TỐI ƯU", "green"
                        elif ue <= (mau_des if cvi_in > 0 else target_mau): status, color = "✅ MONG MUỐN", "blue"
                        else: status, color = "❌ KHÔNG ĐẠT", "red"

                        st.metric("Ue (Độ KĐB mở rộng)", f"{ue:.2f}%")
                        st.markdown(f"Hiệu năng: :{color}[**{status}**]")
                        outliers_val = stats.get('outliers', 0)
                        st.caption(f"Dữ liệu sạch: n={stats['n']}. Loại bỏ: {outliers_val} Outliers.")
                        
                        with st.expander("Chi tiết thành phần (%)"):
                            st.write(f"- Độ chụm ($u_{{prec}}$): {u_prec:.2f}%")
                            st.write(f"- Độ đúng ($u_{{bias}}$): {bias_pct_val:.2f}%")
                            st.write(f"- Tham chiếu ($u_{{ref}}$): {u_ref_pct:.2f}%")
                    else:
                        st.warning("Không có đủ dữ liệu sạch trong khoảng thời gian này.")
                else:
                    st.info("Chưa có dữ liệu nội kiểm.")
                # Trong Tab MU, tại vị trí dòng 1669 bạn gặp lỗi:
                stats = get_clean_stats_3sigma(df_lvl)
                # --- KIỂM TRA ĐIỀU KIỆN TRƯỚC KHI TRUY CẬP STATS ---
                # Sử dụng kiểm tra an toàn: stats không None, là dictionary và có n >= 2
                if stats and isinstance(stats, dict) and stats.get('n', 0) >= 2:
                    # 1. Trích xuất các giá trị an toàn
                    n_v = stats['n']
                    mean_v = stats.get('mean', 0)
                    sd_v = stats.get('sd', 0)
                    cv_v = stats.get('cv', 0)

                    # 2. Hiển thị kết quả thống kê
                    st.write(f"Số lượng mẫu (n): {n_v}")
                    
                    # Sử dụng cột để hiển thị các chỉ số cho đẹp (Tùy chọn)
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Mean", f"{mean_v:.2f}")
                    col2.metric("SD", f"{sd_v:.4f}")
                    col3.metric("CV (%)", f"{cv_v:.2f}%")

                    # 3. Tiếp tục các logic tính toán khác (như MU, Sigma...)
                    # Ví dụ: ue_pct = cv_v * 2
                    
                else:
                    # Trường hợp stats là None hoặc n < 2
                    if not stats:
                        st.info("ℹ️ Chưa có dữ liệu IQC cho lô này.")
                    else:
                        st.warning(f"⚠️ Chỉ có {stats.get('n', 0)} kết quả sạch. Cần tối thiểu 2 kết quả để tính toán thống kê.")
       # --- 5. DIỄN GIẢI LÂM SÀNG & XÁC NHẬN ---
        st.markdown("---")
        st.subheader("📝 Diễn giải kết quả & Xác nhận")
        col_rep1, col_rep2 = st.columns([1, 2])
                        
        with col_rep1:
            val_input = st.number_input("Nhập kết quả BN để tính khoảng sai số:", value=clin_decision if clin_decision > 0 else 0.0)
                            
        with col_rep2:
            if val_input > 0 and mu_results:
                avg_ue = np.mean([v['ue'] for v in mu_results.values()])
                abs_error = (avg_ue / 100) * val_input
                st.info(f"""
                **Kết luận cho Bác sĩ:**
                * Kết quả xét nghiệm: **{val_input}**
                * Khoảng giá trị có thể có của bệnh nhân (Tin cậy 95%): **{val_input - abs_error:.3f}** đến **{val_input + abs_error:.3f}**
                * Ý nghĩa: Sai số tối đa do phương pháp đo là ±{avg_ue:.2f}%.
                """)

        st.divider()
# --- XỬ LÝ XÁC NHẬN XEM XÉT MU ---
        if st.button("✅ Xác nhận Xem xét MU định kỳ hôm nay"):
            try:
                # 1. Lấy ngày hiện tại
                today_str = datetime.now().date().strftime('%Y-%m-%d')
                
                # 2. Gọi hàm cập nhật vào Database (Đảm bảo bạn đã thêm hàm này vào db_module)
                # Giả sử hàm trả về True nếu thành công
                success = db.update_mu_review(current_test['id'], today_str)
                
                if success:
                    st.success(f"Đã ghi nhận ngày xem xét MU: {today_str}. Hệ thống sẽ nhắc nhở sau 6 tháng.")
                    # 3. Ép Streamlit xóa cache để cập nhật lại giao diện (tùy chọn)
                    st.rerun() 
                else:
                    st.error("Không thể cập nhật cơ sở dữ liệu. Vui lòng kiểm tra lại.")
            except Exception as e:
                st.error(f"Lỗi hệ thống: {str(e)}")




# TÍNH SIX-SIGMA

with tabs[4]:
    st.header("5. Six Sigma, QGI & Báo Cáo tổng hợp")

    # 1. BỘ LỌC THỜI GIAN
    with st.expander("📅 Chọn khoảng thời gian báo cáo", expanded=True):
        c_d1, c_d2 = st.columns(2)
        
        # Tính toán ngày bắt đầu: Ngày hiện tại trừ đi 90 ngày (~3 tháng)
        default_start_date = datetime.now() - timedelta(days=90)
        
        # Thiết lập bộ lọc thời gian
        start_d = c_d1.date_input(
            "Từ ngày", 
            default_start_date, # Mặc định lùi 3 tháng
            key="rep_start"
        )
        end_d = c_d2.date_input(
            "Đến ngày", 
            datetime.now(), 
            key="rep_end"
        )
    # 2. LẤY DỮ LIỆU
    df_full_history = db.get_iqc_data_continuous(current_test['id'])
    df_raw = db.get_iqc_data_continuous(current_test['id'])
    df_eqa = db.get_eqa_data(current_test['id'])
    tea = float(current_test.get('tea', 10.0))
    if df_full_history is not None:
        st.write(f"🔍 Tìm thấy tổng {len(df_full_history)} kết quả cho Sigma.")   
    # 3. TÍNH BIAS (Sử dụng trung bình 3 kỳ gần nhất để khớp với Tab MU)
    bias_pct = 0.0
    if not df_eqa.empty:
        recent_eqa = df_eqa.tail(3).copy()
        recent_eqa['pct_bias'] = abs((recent_eqa['lab_value'] - recent_eqa['ref_value']) / recent_eqa['ref_value']) * 100
        bias_pct = recent_eqa['pct_bias'].mean()


# 4. XỬ LÝ DỮ LIỆU NỘI KIỂM & TÍNH SIGMA
    sigma_results = {}
    summary_data = []
    sigma_plot_data = []

    if not df_raw.empty:
        # --- BƯỚC 1: ĐỒNG BỘ HÓA DỮ LIỆU ---
        # Chuyển đổi cột date sang datetime (xử lý cả dạng chuỗi từ nhập tay và timestamp từ excel)
        df_raw['date_dt'] = pd.to_datetime(df_raw['date'], errors='coerce')
        
        # Đảm bảo cột giá trị (value) là số thực để không bị lỗi khi tính Mean/SD
        df_raw['value'] = pd.to_numeric(df_raw['value'], errors='coerce')
        
        # Loại bỏ các dòng bị lỗi dữ liệu nghiêm trọng (không có ngày hoặc không có kết quả)
        df_raw = df_raw.dropna(subset=['date_dt', 'value'])
        
        # Lấy khoảng ngày thực tế có trong DB để gợi ý cho người dùng nếu không thấy dữ liệu
        min_date = df_raw['date_dt'].min().date()
        max_date = df_raw['date_dt'].max().date()
        
        # --- BƯỚC 2: BỘ LỌC THEO THỜI GIAN ---
        df_raw['date_only'] = df_raw['date_dt'].dt.date
        df_filtered = df_raw[(df_raw['date_only'] >= start_d) & (df_raw['date_only'] <= end_d)].copy()

        if df_filtered.empty:
            st.warning(f"⚠️ Không tìm thấy dữ liệu trong khoảng từ {start_d} đến {end_d}.")
            st.info(f"💡 Dữ liệu hiện có sẵn từ ngày **{min_date}** đến **{max_date}**. Vui lòng điều chỉnh lại bộ lọc ngày ở trên.")
        else:
            st.markdown(f"### 🎯 Hiệu năng Six Sigma (Bias sử dụng: {bias_pct:.2f}%)")
            c1, c2, c3 = st.columns(3)
            cols = [c1, c2, c3]

            for lvl in [1, 2, 3]:
                # Lọc theo level (chuyển sang string để so sánh khớp tuyệt đối)
                df_lvl = df_filtered[df_filtered['level'].astype(str) == str(lvl)]
                
                # --- BƯỚC 3: TÍNH TOÁN STATS (Sử dụng hàm get_clean_stats_3sigma đã cải tiến) ---
                stats = get_clean_stats_3sigma(df_lvl)
                current_col = cols[lvl-1]

                if stats and stats['n'] >= 2:
                    cv = stats['cv']
                    # Công thức Sigma: (TEa - Bias) / CV
                    sigma = (tea - bias_pct) / cv if cv > 0 else 0
                    
                    # Tính QGI (Quality Goal Index)
                    qgi = bias_pct / (1.5 * cv) if cv > 0 else 0
                    if qgi < 0.8: qgi_reason = "Lỗi do Độ chụm (CV)"
                    elif 0.8 <= qgi <= 1.2: qgi_reason = "Lỗi do cả Bias và CV"
                    else: qgi_reason = "Lỗi do Độ đúng (Bias)"

                    # Lưu kết quả vào biến tạm
                    sigma_results[lvl] = stats
                    sigma_results[lvl].update({'sigma': sigma, 'qgi': qgi, 'bias': bias_pct})
                    
                    sigma_plot_data.append({'label': f"L{lvl}", 'bias': bias_pct, 'cv': cv})
                    summary_data.append({
                        "Mức độ": f"Level {lvl}",
                        "N (Sạch)": stats['n'],
                        "CV%": cv,
                        "Bias%": bias_pct,
                        "Sigma": sigma,
                        "QGI": qgi,
                        "Đánh giá": "✅ Đạt" if sigma >= 3 else "❌ Không đạt"
                    })

            # Hiển thị UI trực quan vào đúng cột
                    with current_col:
                        with st.container(border=True):
                            st.write(f"**LEVEL {lvl}** (n={stats['n']})")
                            if sigma >= 6: 
                                st.success(f"Sigma: {sigma:.2f}")
                                st.caption("🏆 World Class")
                            elif sigma >= 4: 
                                st.info(f"Sigma: {sigma:.2f}")
                                st.caption("✨ Excellent")
                            elif sigma >= 3: 
                                st.warning(f"Sigma: {sigma:.2f}")
                                st.caption("⚠️ Marginal")
                            else: 
                                st.error(f"Sigma: {sigma:.2f}")
                                st.caption("🚨 Poor")
                            
                            st.divider()
                            st.caption(f"**QGI:** {qgi:.2f}")
                            st.caption(f"🔍 {qgi_reason}")
                else:
                    with current_col:
                        st.info(f"**Level {lvl}**")
                        st.caption("Không đủ dữ liệu sạch (n < 2) để tính toán.")

            # 5. BẢNG TỔNG HỢP
            if summary_data:
                st.markdown("---")
                st.subheader("📋 Bảng tổng hợp hiệu năng")
                df_sum = pd.DataFrame(summary_data)
                
                def color_sigma(val):
                    if val >= 6: return 'background-color: #b3e6ff'
                    elif val >= 4: return 'background-color: #c6efce'
                    elif val >= 3: return 'background-color: #ffeb9c'
                    return 'background-color: #ffc7ce'

                st.dataframe(
                    df_sum.style.map(color_sigma, subset=['Sigma'])
                    .format({'CV%': "{:.2f}", 'Bias%': "{:.2f}", 'Sigma': "{:.2f}", 'QGI': "{:.2f}"}),
                    use_container_width=True, hide_index=True
                )

            # 6. BIỂU ĐỒ DECISION CHART
            st.markdown("---")
            st.subheader("📈 Biểu đồ Method Decision Chart")
            # Đảm bảo hàm plot_sigma_chart đã được định nghĩa
            fig_sigma = plot_sigma_chart(sigma_plot_data, tea)
            if fig_sigma:
                st.pyplot(fig_sigma)
                st.session_state['fig_sigma_report'] = fig_sigma
                # plt.close(fig_sigma)


# === TAB: IMPORT DỮ LIỆU ===📂


# --- KHỞI TẠO DỮ LIỆU MẪU CHO FILE IMPORT ---
# Mẫu 1: IQC hàng ngày
mau_iqc = pd.DataFrame({
    'Tên xét nghiệm': ['Glucose', 'Urea', 'Creatinine'],
    'Ngày xét nghiệm': ['2026-01-12 08:30:00', '2026-01-12 08:35:00', '2026-01-12 08:40:00'],
    'Kết quả': [5.5, 7.2, 85.0],
    'Mức': [1, 1, 2],
    'Ghi chú': ['Nhập máy', 'Nhập máy', 'Chạy lại']
})

# Mẫu 2: EQA (Ngoại kiểm)
mau_eqa = pd.DataFrame({
    'Tên xét nghiệm': ['Glucose', 'AST', 'ALT'],
    'Ngày mẫu': ['2026-01-10', '2026-01-10', '2026-01-10'],
    'Giá trị PXN': [5.6, 35.0, 40.0],
    'Giá trị mục tiêu': [5.4, 38.0, 42.0],
    'SD Nhóm (Group SD)': [0.15, 1.2, 1.5],
    'Mã mẫu': ['Đợt 1', 'Đợt 1', 'Đợt 1']
})

# Mẫu 3: NSX (Lot & Target) - Dựa trên cấu trúc db.add_lot của bạn
mau_nsx = pd.DataFrame({
    'test_name': ['Glucose', 'Glucose', 'Glucose'],
    'lot_number': ['L1-2026', 'L2-2026', 'L3-2026'],
    'level': [1, 2, 3],
    'expiry_date': ['2027-12-31', '2027-12-31', '2027-12-31'],
    'mean': [5.5, 10.2, 15.8],
    'sd': [0.15, 0.35, 0.55],
    'device': ['AU640', 'AU640', 'AU640']
})

# XÁC NHẬN GIÁ TRỊ SỬ DỤNG THEO CLSI EP 15 A3: đã ổn
with tabs[5]:
    st.header("🔬 Xác nhận giá trị sử dụng (CLSI EP15-A3 Standard)")
    
    # Khởi tạo dữ liệu tra cứu và biến chọn xét nghiệm
    if not STANDARD_DB:
        st.error("Cơ sở dữ liệu tiêu chuẩn (STANDARD_DB) chưa được khai báo.")
    else:
        test_selected = st.selectbox("Chọn xét nghiệm xác nhận", options=list(STANDARD_DB.keys()), key="ep15_test_sel")
        ref = STANDARD_DB[test_selected]
        
        col1, col2, col3 = st.columns(3)
        with col1:
            v_target = st.number_input("Giá trị đích (Target Mean)", value=100.0)
            tea_val = st.number_input("TEa cho phép (%)", value=ref['tea'])
            cvi_val = st.write(f"**CVi:** {ref['cvi']}%")
            cvg_val = st.write(f"**CVg:** {ref['cvg']}%")

        with col2:
            v_claim_sr = st.number_input("SD công bố (NSX)", value=2.0)
            
        with col3:
            v_claim_sl = st.number_input("CV% công bố (NSX)", value=3.0)


        # Ma trận nhập liệu 5x5
        st.subheader("Ma trận dữ liệu thực nghiệm (5 Ngày x 5 Lần)")
        input_matrix = []
        rows = st.columns(5)
        for i in range(5):
            with rows[i]:
                raw_input = st.text_area(f"Ngày {i+1}", value="100, 101, 99, 100, 102", key=f"raw_d{i}")
                input_matrix.append([float(x.strip()) for x in raw_input.split(",") if x.strip()])

        if st.button("🚀 Chạy phân tích CLSI EP15-A3", key="btn_run_clsi"):
            results = calculate_clsi_ep15_a3_final(input_matrix, v_claim_sr, v_claim_sl, v_target)
            # Hiển thị cảnh báo ngoại lệ nếu có
            if results['outliers']:
                for out in results['outliers']:
                    st.warning(f"⚠️ Phát hiện giá trị ngoại lệ tại Ngày {out['day']}: **{out['value']}** (G={out['g_score']:.2f}). Giá trị này đã được xử lý để không làm sai lệch ANOVA.")
            else:
                st.info("✅ Không phát hiện giá trị ngoại lệ (Grubbs' Test Pass)")

            # 2. BỔ SUNG CÁC KEY CÒN THIẾU VÀO ĐIỂM KẾT QUẢ ĐỂ XUẤT EXCEL
            results['claim_sr'] = v_claim_sr    # Lưu Sr công bố
            results['claim_sl'] = v_claim_sl    # Lưu Sl công bố (thường gọi là claim_cv)
            results['claim_cv'] = v_claim_sl    # Gán tạm để khớp với hàm Excel cũ của bạn
            results['target_mean'] = v_target   # Lưu giá trị đích
            
            # Tính TE% theo CLSI (Bias% + 1.65 * CV_lab%)
            bias_pct = abs((results['grand_mean'] - v_target) / v_target) * 100
            cv_l_pct = (results['s_l'] / results['grand_mean']) * 100
            results['te_calc'] = bias_pct + 1.65 * cv_l_pct
            # Hiển thị kết quả Độ chụm
            st.markdown("### 📊 Kết quả Độ chụm")
            c1, c2, c3 = st.columns(3)
            c1.metric("Sl thực tế", f"{results['s_l']:.3f}")
            c2.metric("Giới hạn UVL", f"{results['uvl_l']:.3f}")
            c3.write("Kết luận: " + ("✅ ĐẠT" if results['is_precision_pass'] else "❌ KHÔNG ĐẠT"))
            
            # Hiển thị kết quả Độ đúng
            st.markdown("### 🎯 Kết quả Độ đúng")
            t1, t2 = st.columns(2)
            t1.metric("Mean thực tế", f"{results['grand_mean']:.3f}")
            t2.write(f"**Khoảng xác nhận (VI):** {results['vi_range'][0]:.3f} - {results['vi_range'][1]:.3f}")
            
            if results['is_trueness_pass']:
                st.success(f"Độ đúng đạt yêu cầu: Mean nằm trong khoảng VI")
            else:
                st.error(f"Độ đúng KHÔNG đạt: Mean nằm ngoài khoảng VI")

            # Xuất Excel báo cáo
            # Đảm bảo hàm export_verification_excel sử dụng đúng key results['s_l'] và results['grand_mean']
            report_data = export_verification_excel(test_selected, ref, input_matrix, results)
            st.download_button("📥 Tải báo cáo lưu trữ", data=report_data, file_name=f"EP15A3_{test_selected}.xlsx", key="dl_clsi")


# IMPORT EXCEL
with tabs[6]: 
    import_sub1, import_sub2 = st.tabs(["📥 Xử lý Import Dữ liệu", "🔗 Cấu hình Mapping"])

    # --- SUB-TAB 1: XỬ LÝ IMPORT ---
    with import_sub1:
        # --- PHẦN 1: IMPORT IQC ---
        st.markdown("### 🧬 1. Nhập kết quả Nội kiểm (IQC)")
        st.download_button(
            label="📥 Tải file mẫu IQC (.xlsx)",
            data=công_cụ_tạo_mẫu(mau_iqc, "Mau_IQC.xlsx"),
            file_name="Mau_Import_IQC.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )        
        with st.container(border=True):
            uploaded_file = st.file_uploader("Chọn file Excel kết quả từ máy xét nghiệm (AU640, Abbott, Roche...)", type=["xlsx", "xls"], key="iqc_main_uploader")

            if uploaded_file:
                df_preview = pd.read_excel(uploaded_file)
                with st.expander("🔍 Xem trước dữ liệu vừa tải lên", expanded=False):
                    st.dataframe(df_preview.head(10), use_container_width=True)

                if 'Tên xét nghiệm' in df_preview.columns:
                    excel_names = df_preview['Tên xét nghiệm'].unique().tolist()
                    unmapped_list = db.get_unmapped_tests(excel_names)

                    if unmapped_list:
                        st.warning(f"⚠️ **Phát hiện {len(unmapped_list)} xét nghiệm chưa được ánh xạ (Mapping)**")
                        st.info("Các mã lạ: " + ", ".join([f"`{name}`" for name in unmapped_list]))
                        
                        col_msg, col_btn = st.columns([3, 1])
                        with col_msg:
                            st.error("Vui lòng sang tab **'Cấu hình Mapping'** để thiết lập trước khi Import.")
                        with col_btn:
                            st.button("🚀 Xác nhận Import", disabled=True, use_container_width=True, key="btn_iqc_disabled")
                    else:
                        st.success("✅ Dữ liệu hợp lệ. Tất cả xét nghiệm đã được ánh xạ.")
                        if st.button("🚀 Xác nhận Import IQC vào Database", type="primary", use_container_width=True, key="btn_iqc_confirm"):
                            with st.spinner("Đang lưu dữ liệu..."):
                                if 'Ngày xét nghiệm' in df_preview.columns:
                                    df_preview['Ngày xét nghiệm'] = pd.to_datetime(df_preview['Ngày xét nghiệm']).dt.strftime('%Y-%m-%d %H:%M:%S')
                                success_count, logs = db.import_iqc_from_dataframe(df_preview)
                                if success_count > 0:
                                    st.toast(f"Đã Import {success_count} kết quả!", icon="✅")
                                    st.success(f"✅ Thành công: {success_count} kết quả. Dữ liệu đã sẵn sàng tại Tab Six Sigma.")
                                    time.sleep(1)
                                if logs:
                                    with st.expander("📝 Chi tiết log xử lý"):
                                        for log in logs: st.write(log)
                                st.rerun()
                else:
                    st.error("❌ File không đúng định dạng: Thiếu cột **'Tên xét nghiệm'**")

        st.markdown("---")

        # --- PHẦN 2: IMPORT EQA ---
        st.markdown("### 🧪 2. Nhập kết quả Ngoại kiểm (EQA)")
        st.download_button(
            label="📥 Tải file mẫu EQA (.xlsx)",
            data=công_cụ_tạo_mẫu(mau_eqa, "Mau_EQA.xlsx"),
            file_name="Mau_Import_EQA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        with st.expander("📥 Click để mở trình Import EQA", expanded=False):
            st.info("Yêu cầu file có các cột: `Tên xét nghiệm`, `Giá trị PXN`, `Giá trị mục tiêu`, `Ngày nhận kết quả`")
            eqa_file = st.file_uploader("Chọn file Excel EQA", type=["xlsx", "xls"], key="eqa_uploader")

            if eqa_file:
                df_eqa_preview = pd.read_excel(eqa_file)
                st.dataframe(df_eqa_preview.head())

                if st.button("🚀 Xác nhận Import EQA", key="btn_eqa_confirm"):
                    with st.spinner("Đang xử lý..."):
                        count, logs = db.import_eqa_from_dataframe(df_eqa_preview)
                        if count > 0:
                            st.success(f"✅ Đã thêm {count} kết quả EQA.")
                            if logs:
                                with st.expander("Xem chi tiết"):
                                    for log in logs: st.write(f"- {log}")
                            time.sleep(1)
                            st.rerun()

        st.markdown("---")

        # --- PHẦN 3: IMPORT GIÁ TRỊ NHÀ SẢN XUẤT (NSX) ---
        st.markdown("### 📋 3. Nhập giá trị Target từ Nhà sản xuất (Lot, Mean, SD)")
        st.download_button(
            label="📥 Tải file mẫu NSX (.xlsx)",
            data=công_cụ_tạo_mẫu(mau_nsx, "Mau_NSX.xlsx"),
            file_name="Mau_Import_NSX.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        with st.expander("📥 Click để mở trình Import Lot & Target", expanded=False):
            st.info("Cấu trúc file mẫu: `test_name`, `lot_number`, `level`, `expiry_date`, `mean`, `sd`")
            nsx_file = st.file_uploader("Chọn file Excel/CSV chứa giá trị NSX", type=["xlsx", "csv"], key="nsx_target_uploader")
            
            if nsx_file:
                try:
                    df_nsx = pd.read_csv(nsx_file) if nsx_file.name.endswith('.csv') else pd.read_excel(nsx_file)
                    st.dataframe(df_nsx.head(), use_container_width=True)
                    
                    if st.button("🚀 Xác nhận Import giá trị NSX", type="primary"):
                        success_count = 0
                        with st.spinner("Đang cập nhật Lot..."):
                            for _, row in df_nsx.iterrows():
                                # Tìm test_id từ test_name (ánh xạ tên xét nghiệm)
                                test_info = db.get_test_by_name(row['test_name'])
                                if test_info:
                                    db.add_lot(
                                        test_id=test_info['id'],
                                        lot_number=str(row['lot_number']),
                                        level=int(row['level']),
                                        method="Import NSX",
                                        expiry_date=str(row['expiry_date']),
                                        mean=float(row['mean']),
                                        sd=float(row['sd'])
                                    )
                                    success_count += 1
                        
                        st.success(f"✅ Đã cập nhật thành công {success_count} thông số Lot vào hệ thống!")
                        time.sleep(1)
                        st.rerun()
                        
                except Exception as e:
                    st.error(f"Lỗi khi xử lý file NSX: {e}")

### Bạn có muốn tôi thiết kế một nút "Tải File Excel Mẫu" chứa đúng các tiêu đề cột này để nhân viên chỉ cần điền dữ liệu không? Điều này sẽ giúp tránh lỗi sai tên cột khi Import.

        # --- PHẦN 3: XUẤT BÁO CÁO ---
        st.markdown("### 📄 3. Xuất Báo Cáo")
        with st.container(border=True):
            st.write("Khởi tạo báo cáo tổng hợp bao gồm biểu đồ LJ, Sigma Chart và V-Mask dựa trên dữ liệu hiện tại.")
            if st.button("📥 Khởi tạo file Báo Cáo Tổng Hợp (Excel)", key="btn_export_all", type="secondary"):
                
                if df_filtered is None or df_filtered.empty:
                    st.error("❌ Không có dữ liệu IQC (Vui lòng chọn Test và Khoảng ngày ở Sidebar)")
                else:
                    with st.spinner("🚀 Đang vẽ biểu đồ và khởi tạo file..."):
                        try:
                            import io
                            import matplotlib.pyplot as plt

                            # --- 1. CHUẨN BỊ DỮ LIỆU ---
                            df_prep = df_filtered.copy()

                            # Lấy Mean/SD từ current_test hoặc từ dữ liệu thực tế để tính Westgard
                            # Giả sử current_test chứa thông tin cài đặt của Test đó
                            mean_map = {1: current_test.get('mean_l1', 0), 2: current_test.get('mean_l2', 0), 3: current_test.get('mean_l3', 0)}
                            sd_map = {1: current_test.get('sd_l1', 0), 2: current_test.get('sd_l2', 0), 3: current_test.get('sd_l3', 0)}

                            # --- 2. QUAN TRỌNG: TÍNH LẠI WESTGARD TRƯỚC KHI XUẤT ---
                            # Gọi hàm này để đảm bảo cột 'Violation' có dữ liệu
                            # (Hàm get_westgard_violations tôi đã gửi ở những phản hồi đầu tiên)
                            df_prep = get_westgard_violations(df_prep, mean_map, sd_map)

                            # --- 3. XỬ LÝ EQA & BIỂU ĐỒ (Giữ nguyên logic của bạn) ---
                            for lvl in [1, 2]:
                                mask = df_prep['level'] == lvl
                                if mask.any():
                                    # Nếu thiếu target_mean trong DB, lấy trung bình thực tế
                                    if 'target_mean' not in df_prep.columns or df_prep.loc[mask, 'target_mean'].isnull().all():
                                        df_prep.loc[mask, 'target_mean'] = mean_map.get(lvl) if mean_map.get(lvl) else df_prep.loc[mask, 'value'].mean()
                                    if 'target_sd' not in df_prep.columns or df_prep.loc[mask, 'target_sd'].isnull().all():
                                        df_prep.loc[mask, 'target_sd'] = sd_map.get(lvl) if sd_map.get(lvl) else df_prep.loc[mask, 'value'].std()

                            # --- LOGIC XỬ LÝ DỮ LIỆU & VẼ BIỂU ĐỒ (Giữ nguyên nội dung của bạn) ---
                            df_prep = df_filtered.copy()
                            for lvl in [1, 2]:
                                mask = df_prep['level'] == lvl
                                if mask.any():
                                    if 'target_mean' not in df_prep.columns or df_prep.loc[mask, 'target_mean'].isnull().all():
                                        df_prep.loc[mask, 'target_mean'] = df_prep.loc[mask, 'value'].mean()
                                    if 'target_sd' not in df_prep.columns or df_prep.loc[mask, 'target_sd'].isnull().all():
                                        actual_sd = df_prep.loc[mask, 'value'].std()
                                        df_prep.loc[mask, 'target_sd'] = actual_sd if (actual_sd and actual_sd > 0) else 1.0

                            df_eqa_prep = df_eqa.copy() if (df_eqa is not None and not df_eqa.empty) else pd.DataFrame()
                            if not df_eqa_prep.empty:
                                if 'sdi' not in df_eqa_prep.columns:
                                    m_e = df_eqa_prep['target'].mean() if 'target' in df_eqa_prep.columns else df_eqa_prep['value'].mean()
                                    s_e = df_eqa_prep['sd_target'].mean() if 'sd_target' in df_eqa_prep.columns else 1.0
                                    df_eqa_prep['sdi'] = (df_eqa_prep['value'] - m_e) / s_e
                                if 'CUSUM' not in df_eqa_prep.columns:
                                    df_eqa_prep = df_eqa_prep.sort_values('date')
                                    df_eqa_prep['CUSUM'] = df_eqa_prep['sdi'].cumsum()

                            def fig_to_bytes_internal(fig_obj):
                                if fig_obj is None: return None
                                buf = io.BytesIO()
                                fig_obj.savefig(buf, format='png', bbox_inches='tight', dpi=100)
                                plt.close(fig_obj)
                                return buf.getvalue()

                            img_lj = fig_to_bytes_internal(plot_levey_jennings(df_prep, f"Biểu đồ LJ: {current_test['name']}"))
                            img_sigma = fig_to_bytes_internal(plot_sigma_chart(sigma_plot_data, tea))
                            img_vmask = None
                            if not df_eqa_prep.empty:
                                fig_vmask_raw, _ = plot_cusum_chart(df_eqa_prep)
                                img_vmask = fig_to_bytes_internal(fig_vmask_raw)

        # --- 4. GỌI HÀM TẠO EXCEL ---
                            excel_data = generate_excel_report_comprehensive(
                                test_info=current_test, 
                                df_full_iqc=df_prep,  # Lúc này df_prep đã CÓ cột 'Violation'
                                df_eqa=df_eqa_prep,
                                mu_data=st.session_state.get('mu_results', {}), 
                                sigma_data=sigma_results,
                                img_lj=img_lj, 
                                img_sigma=img_sigma, 
                                img_vmask=img_vmask,
                                report_period=(start_d, end_d), 
                                mau_limits=(mau_min, mau_des, mau_opt)
                            )

                            st.download_button(
                                label="📂 Tải file Báo cáo ngay",
                                data=excel_data,
                                file_name=f"Bao_cao_QLCL_{current_test['name']}_{start_d.strftime('%Y%m%d')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )
                        except Exception as e:
                            st.error(f"❌ Lỗi: {str(e)}")
    # --- SUB-TAB 2: QUẢN LÝ MAPPING ---
    with import_sub2:
        st.markdown("### 🔗 Thiết lập mapping Xét nghiệm")
        
        # Thêm mới Mapping
        with st.container(border=True):
            st.write("**➕ Thêm Mapping mới**")
            all_tests_df = db.get_all_tests()
            test_dict = {row['name']: row['id'] for _, row in all_tests_df.iterrows()}
            
            c1, c2, c3 = st.columns([2, 2, 1])
            with c1:
                sel_internal = st.selectbox("Xét nghiệm hệ thống:", list(test_dict.keys()), key="map_sel_int")
            with c2:
                suggested_name = unmapped_list[0] if 'unmapped_list' in locals() and unmapped_list else ""
                new_ext = st.text_input("Tên trên Excel:", value=suggested_name, key="map_ext_input")
            with c3:
                st.write(" ") # Tạo khoảng cách
                if st.button("Lưu Mapping", use_container_width=True, type="primary"):
                    if new_ext:
                        db.add_mapping(test_dict[sel_internal], new_ext)
                        st.success("Đã lưu!"); time.sleep(0.5); st.rerun()

        st.markdown("#### 📋 Danh sách mapping")
        df_map = db.get_all_mappings()
        if not df_map.empty:
            edited_map_df = st.data_editor(
                df_map[['id', 'internal_name', 'external_name']],
                column_config={
                    "id": None,
                    "internal_name": st.column_config.TextColumn("Xét nghiệm hệ thống", disabled=True),
                    "external_name": st.column_config.TextColumn("Tên trên Excel (Sửa tại đây)", required=True),
                },
                num_rows="dynamic",
                use_container_width=True,
                key="mapping_table_editor"
            )

            if st.button("💾 Lưu tất cả thay đổi trên bảng Mapping", ):
                # Logic xử lý cập nhật (Xóa/Sửa) - Giữ nguyên của bạn
                current_ids = set(df_map['id'])
                edited_ids = set(edited_map_df['id'])
                for d_id in (current_ids - edited_ids): db.delete_mapping(d_id)
                for _, row in edited_map_df.iterrows():
                    old_data = df_map[df_map['id'] == row['id']].iloc[0]
                    if row['external_name'] != old_data['external_name']:
                        db.update_mapping(row['id'], row['external_name'])
                st.success("Đã cập nhật!"); st.rerun()
        else:
            st.info("Chưa có dữ liệu mapping.")

# === TAB 6: QUẢN TRỊ (ADMIN) ===

# Lấy mật khẩu quản trị hiện tại từ DB (Mặc định là 'admin123' nếu chưa thiết lập)
ADMIN_PASSWORD_KEY = "admin_password"
current_admin_pwd = db.get_setting(ADMIN_PASSWORD_KEY, "admin123")


# === TAB 6: QUẢN TRỊ (ADMIN) ===
with tabs[7]:
    st.header("🔐 Khu vực Quản trị")
    st.sidebar.markdown("---")
    if st.sidebar.button("⚙️ Nâng cấp Database"):
        success, msg = upgrade_database_structure()
        if success:
            st.sidebar.success(msg)
            st.rerun()
        else:
            st.sidebar.error(msg)
    # 1. PHẦN CÀI ĐẶT MẬT KHẨU QUẢN TRỊ
    with st.expander("🔑 Cài đặt Mật khẩu Quản trị", expanded=False):
        st.info(f"Mật khẩu hiện tại (để đăng nhập bên dưới): ***{len(current_admin_pwd)} ký tự***")
        with st.form("set_admin_pwd_form"):
            new_pwd = st.text_input("Mật khẩu Mới", type="password")
            confirm_pwd = st.text_input("Xác nhận Mật khẩu Mới", type="password")

            if st.form_submit_button("Lưu Mật khẩu Mới"):
                if new_pwd != confirm_pwd:
                    st.error("Mật khẩu xác nhận không khớp.")
                elif len(new_pwd) < 6:
                    st.error("Mật khẩu phải có ít nhất 6 ký tự.")
                else:
                    # Lưu mật khẩu mới vào DB
                    db.set_setting(ADMIN_PASSWORD_KEY, new_pwd)
                    st.success("Đã cập nhật mật khẩu quản trị! Vui lòng đăng nhập lại.")
                    st.rerun()

    st.markdown("---")

    # 2. PHẦN ĐĂNG NHẬP VÀ MỞ KHÓA
    pwd = st.text_input("Nhập mật khẩu quản trị", type="password", key="admin_login_pwd")

    if pwd == current_admin_pwd: # So sánh với mật khẩu đã lưu trong DB
        st.success("Đã mở khóa!")
        
        st.markdown("---")

        # 4. VÙNG NGUY HIỂM (GIỮ NGUYÊN LOGIC)
        with st.expander("⚠️ Vùng nguy hiểm: Reset Dữ liệu"):
            st.warning("Hành động này sẽ xóa dữ liệu! Hãy cẩn thận.")
            if st.button(f"Xóa TOÀN BỘ dữ liệu IQC của Test: {current_test['name']}"):
                # Logic xóa
                lots = db.get_lots_for_test(current_test['id'])
                for _, l in lots.iterrows(): 
                    # Giả định db.delete_lot đã tồn tại và hoạt động đúng
                    db.delete_lot(l['id']) 
                st.success("Đã xóa sạch dữ liệu IQC!")
                st.rerun()
                
        # 5. BACKUP DATABASE (GIỮ NGUYÊN LOGIC)
        with st.expander("📋 Backup Database"):
            # Thêm timestamp vào tên file
            with open("lab_data.db", "rb") as f:
                st.download_button("Tải file Backup (.db)", f, f"lab_data_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db")
            
    elif pwd:
        st.error("Sai mật khẩu.")
        # Giao diện nút bấm trên Sidebar
