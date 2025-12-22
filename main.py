# File: main.py
import streamlit as st
import pandas as pd
import sqlite3
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime, date, timedelta
import io
import xlsxwriter
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

from db_module import DBManager  

# Nhập hàm từ file license_check.py
# from license_check import verify_license, get_hwid 

# Giả lập hàm kiểm tra ngay trong main nếu không muốn tách file
#if not verify_license():
#    st.error("🚫 PHẦN MỀM CHƯA ĐƯỢC KÍCH HOẠT")
#    st.info(f"Vui lòng gửi mã máy sau cho quản trị viên: **{get_hwid()}**")
#    st.stop() # Dừng toàn bộ ứng dụng nếu không có bản quyền

# --- TIẾP TỤC CODE PHẦN MỀM CỦA BẠN ---
#st.success("Bản quyền hợp lệ. Chào mừng bạn!")
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

# Sửa trong hàm manage_test_mapping của main.py
def manage_test_mapping():
    st.subheader("🔗 Mapping Tên xét nghiệm từ máy")
    
    # Lấy dữ liệu và đảm bảo nó là danh sách các hàng
    df_tests = db.get_all_tests() 
    
    # Nếu db.get_all_tests() trả về DataFrame, hãy dùng .to_dict('records')
    if isinstance(df_tests, pd.DataFrame):
        all_tests = df_tests.to_dict('records')
    else:
        all_tests = df_tests # Giả sử đã là list rồi
        
    if not all_tests:
        st.warning("Chưa có xét nghiệm nào trong hệ thống.")
        return

    col1, col2 = st.columns(2)
    with col1:
        # Bây giờ x sẽ là một Dictionary, có thể truy cập x['name']
        selected_test = st.selectbox(
            "Chọn xét nghiệm trong PM:", 
            all_tests, 
            format_func=lambda x: x['name']
        )

def process_bulk_import(df):
    # (Giữ nguyên logic xử lý database của bạn ở đây)
    # Hàm này dùng để chạy vòng lặp insert dữ liệu
    conn = sqlite3.connect("lab_data.db")
    # ... logic như bạn đã viết ...
    return summary
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

    # 3. Vẽ các điểm QC thực tế (L1: Xanh dương, L2: Cam)
    colors_qc = ['#1f77b4', '#ff7f0e'] 
    
    for i, pt in enumerate(sigma_plot_data):
        label_text = pt.get('label', f'L{i+1}')
        color = colors_qc[i % 2]
        
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
def get_westgard_violations(df, mean, sd):
    """
    Hàm tổng hợp kiểm tra toàn bộ quy tắc Westgard (Within & Across level).
    Input: df (DataFrame), mean (giá trị trung bình), sd (độ lệch chuẩn).
    """
    if df.empty:
        df['Violation'] = ""
        return df

    # 1. TỰ ĐỘNG NHẬN DIỆN TÊN CỘT DỮ LIỆU
    potential_cols = ['lab_value', 'value', 'result']
    actual_col = next((c for c in potential_cols if c in df.columns), None)
    
    if not actual_col:
        df['Violation'] = ""
        return df

    # Tạo bản sao và tính toán Z-Score
    df_calc = df.sort_values(by=['date', 'level']).copy()
    df_calc['z_score'] = (df_calc[actual_col] - mean) / sd
    
    # Dictionary để lưu lỗi theo ID: {id: [danh sách lỗi]}
    violation_map = {row_id: [] for row_id in df_calc['id']}

    # 2. KIỂM TRA ACROSS-LEVELS (So sánh giữa các Level trong cùng 1 ngày)
    for date, df_day in df_calc.groupby('date'):
        if len(df_day) >= 2:
            row_l1 = df_day[df_day['level'] == 1]
            row_l2 = df_day[df_day['level'] == 2]
            
            if not row_l1.empty and not row_l2.empty:
                z_l1 = row_l1['z_score'].iloc[0]
                z_l2 = row_l2['z_score'].iloc[0]
                id_l1 = row_l1['id'].iloc[0]
                id_l2 = row_l2['id'].iloc[0]

                # R-4s (Across): Chênh lệch >= 4SD giữa 2 Level (1 cái >+2, 1 cái <-2)
                if abs(z_l1 - z_l2) >= 4 and (z_l1 * z_l2 < 0):
                    msg = "R-4s (Across)"
                    violation_map[id_l1].append(msg)
                    violation_map[id_l2].append(msg)
                
                # 2-2s (Across): Cả 2 level vượt 2SD cùng phía
                elif (z_l1 > 2 and z_l2 > 2) or (z_l1 < -2 and z_l2 < -2):
                    msg = "2-2s (Across)"
                    violation_map[id_l1].append(msg)
                    violation_map[id_l2].append(msg)

    # 3. KIỂM TRA WITHIN-LEVEL (Chuỗi thời gian cho từng Level)
    for level, df_level in df_calc.groupby('level'):
        df_level = df_level.sort_values(by='date').reset_index(drop=True)
        z = df_level['z_score'].tolist()
        ids = df_level['id'].tolist()
        n = len(z)

        for i in range(n):
            curr_id = ids[i]
            
            # --- QUY TẮC TỪ CHỐI (REJECTION) ---
            # 1-3s
            if abs(z[i]) > 3:
                violation_map[curr_id].append("1-3s")

            if i >= 1:
                # 2-2s (Within)
                if (z[i] > 2 and z[i-1] > 2) or (z[i] < -2 and z[i-1] < -2):
                    violation_map[curr_id].append("2-2s")
                
                # R-4s (Within): Hiệu số Z giữa 2 điểm liên tiếp vượt quá 4
                if abs(z[i] - z[i-1]) > 4:
                    violation_map[curr_id].append("R-4s")

            # 4-1s (4 điểm liên tiếp vượt 1SD cùng phía)
            if i >= 3:
                sub_z = z[i-3:i+1]
                if all(val > 1 for val in sub_z) or all(val < -1 for val in sub_z):
                    violation_map[curr_id].append("4-1s")

            # 10x (10 điểm liên tiếp cùng phía so với Mean)
            if i >= 9:
                sub_z = z[i-9:i+1]
                if all(val > 0 for val in sub_z) or all(val < 0 for val in sub_z):
                    violation_map[curr_id].append("10x")

            # --- QUY TẮC CẢNH BÁO (WARNING) ---
            # 1-2s: Nếu chưa dính lỗi từ chối nào mà vượt 2SD
            if not violation_map[curr_id] and abs(z[i]) > 2:
                violation_map[curr_id].append("1-2s")

            # Trend: 6 điểm liên tiếp tăng hoặc giảm
            if i >= 5:
                sub_6 = z[i-5:i+1]
                if all(sub_6[k] < sub_6[k+1] for k in range(5)):
                    violation_map[curr_id].append("Trend (Tăng)")
                elif all(sub_6[k] > sub_6[k+1] for k in range(5)):
                    violation_map[curr_id].append("Trend (Giảm)")

    # 4. ÁNH XẠ KẾT QUẢ LẠI DATAFRAME GỐC
    # Chuyển list lỗi thành chuỗi cách nhau bởi dấu phẩy, loại bỏ trùng lặp
    final_violations = []
    for row_id in df['id']:
        errors = sorted(list(set(violation_map.get(row_id, []))))
        final_violations.append(", ".join(errors))
    
    df['Violation'] = final_violations
    return df

    # --- LOGIC KIỂM TRA WESTGARD NÂNG CAO (Thay thế cho evaluate_westgard_series) ---
# --- CÁC HÀM KIỂM TRA QUY TẮC WESTGARD ---

def evaluate_westgard_series(df_sub):
    """
    Kiểm tra các quy tắc Westgard cho một chuỗi kết quả QC (thường là 20-30 điểm gần nhất).
    df_sub: DataFrame chứa cột 'value', 'target_mean', 'target_sd', 'z_score'
    """
    if df_sub.empty:
        return []

    violations = []
    # Chuyển dữ liệu sang list để duyệt cho nhanh
    values = df_sub['value'].tolist()
    z = df_sub['z_score'].tolist()
    n = len(values)

    for i in range(n):
        # 1. Quy tắc 1-3s (Lỗi ngẫu nhiên hoặc hệ thống nghiêm trọng)
        if abs(z[i]) > 3:
            violations.append(f"Điểm {i+1}: Vi phạm 1-3s (Z={z[i]:.2f})")

        if i > 0:
            # 2. Quy tắc 2-2s (Lỗi hệ thống)
            # Hai điểm liên tiếp cùng nằm ngoài +2s hoặc cùng ngoài -2s
            if (z[i] > 2 and z[i-1] > 2) or (z[i] < -2 and z[i-1] < -2):
                violations.append(f"Điểm {i} & {i+1}: Vi phạm 2-2s")
            
            # 3. Quy tắc R-4s (Lỗi ngẫu nhiên)
            # Hiệu số Z giữa 2 điểm liên tiếp vượt quá 4
            if abs(z[i] - z[i-1]) > 4:
                violations.append(f"Điểm {i} & {i+1}: Vi phạm R-4s")

        if i > 3:
            # 4. Quy tắc 4-1s (Lỗi hệ thống)
            # Bốn điểm liên tiếp cùng nằm về một phía và vượt quá 1s
            sub_z = z[i-3:i+1]
            if all(val > 1 for val in sub_z) or all(val < -1 for val in sub_z):
                violations.append(f"Cụm điểm {i-2} đến {i+1}: Vi phạm 4-1s")

        if i > 9:
            # 5. Quy tắc 10-x (Lỗi hệ thống)
            # Mười điểm liên tiếp nằm về một phía của trị số trung bình
            sub_z = z[i-9:i+1]
            if all(val > 0 for val in sub_z) or all(val < 0 for val in sub_z):
                violations.append(f"Cụm điểm {i-8} đến {i+1}: Vi phạm 10-x")

    # Loại bỏ các thông báo trùng lặp và trả về
    return list(set(violations))
# 1. HÀM HỖ TRỢ KIỂM TRA R-4S & 2-2s ACROSS GIỮA L1 VÀ L2
def check_cross_level_rules(df_day):
    """
    Kiểm tra các quy tắc liên quan đến so sánh giữa các Level trong CÙNG 1 NGÀY.
    Input: df_day (DataFrame chứa dữ liệu của 1 ngày cụ thể).
    Output: Dictionary các lỗi {iqc_id: "Tên lỗi"}
    """
    errors = {}
    
    # Cần tối thiểu 2 level để so sánh
    if len(df_day) < 2 or 'z_score' not in df_day.columns:
        return errors
        
    try:
        # Lấy dữ liệu của L1 và L2
        row_l1 = df_day[df_day['level'] == 1]
        row_l2 = df_day[df_day['level'] == 2]
        
        if row_l1.empty or row_l2.empty:
            return errors # Thiếu 1 trong 2 level
            
        z_l1 = row_l1['z_score'].iloc[0]
        z_l2 = row_l2['z_score'].iloc[0]
        id_l1 = row_l1['id'].iloc[0]
        id_l2 = row_l2['id'].iloc[0]
        
    except (IndexError, KeyError):
        return errors
        
    # --- A. Kiểm tra R-4s (Rejection) ---
    # Điều kiện: Chênh lệch >= 4SD VÀ nằm về 2 phía khác nhau (1 cái > +2, 1 cái < -2)
    delta_z = abs(z_l1 - z_l2)
    if delta_z >= 4:
        condition1 = (z_l1 >= 2 and z_l2 <= -2)
        condition2 = (z_l2 >= 2 and z_l1 <= -2)
        
        if condition1 or condition2:
            rule = "R-4s: Chênh lệch > 4SD (Lỗi Ngẫu nhiên)"
            errors[id_l1] = rule
            errors[id_l2] = rule
            return errors # Nếu dính R-4s thì return luôn, không check 2-2s nữa

    # --- B. Kiểm tra 2-2s Across Levels (Rejection) ---
    # Điều kiện: Cả L1 và L2 đều vượt quá 2SD CÙNG PHÍA
    if (z_l1 > 2 and z_l2 > 2) or (z_l1 < -2 and z_l2 < -2):
        rule = "2-2s(Across): L1 & L2 vượt 2SD cùng phía (Lỗi Hệ thống)"
        errors[id_l1] = rule
        errors[id_l2] = rule
            
    return errors

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

# --- 2. HÀM VẼ BIỂU ĐỒ ---
def plot_levey_jennings(df, title, show_legend=True):
    if df.empty: return None
    
    fig, ax = plt.subplots(figsize=(10, 5))
    
    # Vẽ các vùng SD
    ax.axhline(0, color='green', lw=1, label='Mean')
    for sd in [1, 2, 3]:
        # Dùng màu rõ ràng hơn cho 2SD (red) và 3SD (black)
        color_sd = 'gold' if sd==1 else ('red' if sd==2 else 'black')
        ax.axhline(sd, color=color_sd, ls='--', alpha=0.5)
        ax.axhline(-sd, color=color_sd, ls='--', alpha=0.5)

    colors = {1: 'blue', 2: 'orange'}
    
    # Tính Z-Score và Vẽ
    for lvl in [1, 2]:
        d_lvl = df[df['level'] == lvl].copy()
        if not d_lvl.empty:
            # Tính Z-Score dựa trên Target Mean/SD của TỪNG LOT
            d_lvl['z'] = (d_lvl['value'] - d_lvl['target_mean']) / d_lvl['target_sd']
            
            # Vẽ đường nối
            ax.plot(d_lvl['date'], d_lvl['z'], color=colors[lvl], alpha=0.5, lw=1)
            ax.scatter(d_lvl['date'], d_lvl['z'], color=colors[lvl], s=30, label=f"Level {lvl}", zorder=3)
            
            # Đánh dấu thay đổi Lot
            changes = d_lvl.drop_duplicates(subset=['lot_number'], keep='first')
            for _, r in changes.iterrows():
                if r['date'] != df['date'].min():
                    ax.axvline(r['date'], color='gray', ls=':', alpha=0.5)
                    # Ghi số Lot ở trên cùng
                    ax.text(r['date'], 3.2, r['lot_number'], rotation=90, fontsize=8, ha='right', va='center')

    ax.set_ylim(-4, 4)
    ax.set_ylabel("Z-Score")
    ax.set_title(title)
    if show_legend: ax.legend(loc='upper right')
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

# --- 2. HÀM VẼ BIỂU ĐỒ ---
def plot_levey_jennings(df, title, show_legend=True):
    if df.empty: return None

    fig, ax = plt.subplots(figsize=(10, 5))
    
    # Vẽ các vùng SD
    ax.axhline(0, color='green', lw=1, label='Mean')
    for sd in [1, 2, 3]:
        # Dùng màu rõ ràng hơn cho 2SD (red) và 3SD (black)
        color_sd = 'gold' if sd==1 else ('red' if sd==2 else 'black')
        ax.axhline(sd, color=color_sd, ls='--', alpha=0.5)
        ax.axhline(-sd, color=color_sd, ls='--', alpha=0.5)

    colors = {1: 'blue', 2: 'orange'}
    
    # Tính Z-Score và Vẽ
    for lvl in [1, 2]:
        d_lvl = df[df['level'] == lvl].copy()
        if not d_lvl.empty:
            # Tính Z-Score dựa trên Target Mean/SD của TỪNG LOT
            d_lvl['z'] = (d_lvl['value'] - d_lvl['target_mean']) / d_lvl['target_sd']
            
            # Vẽ đường nối
            ax.plot(d_lvl['date'], d_lvl['z'], color=colors[lvl], alpha=0.5, lw=1)
            ax.scatter(d_lvl['date'], d_lvl['z'], color=colors[lvl], s=30, label=f"Level {lvl}", zorder=3)
            
            # Đánh dấu thay đổi Lot
            changes = d_lvl.drop_duplicates(subset=['lot_number'], keep='first')
            for _, r in changes.iterrows():
                if r['date'] != df['date'].min():
                    ax.axvline(r['date'], color='gray', ls=':', alpha=0.5)
                    # Ghi số Lot ở trên cùng
                    ax.text(r['date'], 3.2, r['lot_number'], rotation=90, fontsize=8, ha='right', va='center')

    ax.set_ylim(-4, 4)
    ax.set_ylabel("Z-Score")
    ax.set_title(title)
    if show_legend: ax.legend(loc='upper right')
    plt.tight_layout()
    return fig
    fig = plot_lj_chart(test_info['name'], iqc_data, st.session_state.get('last_update'))

# --- 3. XUẤT BÁO CÁO EXCEL CHUYÊN NGHIỆP (Đã cập nhật Westgard) ---
# Cập nhật tham số đầu vào (thêm mau_limits)
def generate_excel_report_comprehensive(test_info, df_full_iqc, df_eqa, mu_data, sigma_data, img_lj, img_sigma, report_period, mau_limits):
    m_min, m_des, m_opt = mau_limits

    start_date, end_date = report_period
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True, 'nan_inf_to_errors': True})

    # Format
    fmt_head = wb.add_format({'bold': True, 'align': 'center', 'bg_color': '#DDEBF7', 'border': 1, 'valign': 'vcenter', 'text_wrap': True})
    fmt_cell = wb.add_format({'align': 'center', 'border': 1, 'valign': 'vcenter'})
    fmt_num = wb.add_format({'num_format': '0.0000', 'align': 'center', 'border': 1})
    fmt_err = wb.add_format({'color': 'white', 'bg_color': 'red', 'bold': True, 'align': 'center', 'border': 1}) # Lỗi từ chối
    fmt_warn = wb.add_format({'color': 'black', 'bg_color': 'yellow', 'bold': True, 'align': 'center', 'border': 1}) # Lỗi cảnh báo
    fmt_bold = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})

# === SHEET 1: TỔNG HỢP & IQC ===
    ws1 = wb.add_worksheet("Nội Kiểm & Tổng Hợp")
    # Mở rộng độ rộng cột (Cột G cho hành động khắc phục cần rộng hơn)
    ws1.set_column('A:A', 12); ws1.set_column('B:E', 10); ws1.set_column('F:F', 20); ws1.set_column('G:G', 35)
    
    # 1. TIÊU ĐỀ CHÍNH VÀ THÔNG TIN HÀNH CHÍNH
    ws1.merge_range('A1:G1', f"BÁO CÁO QUẢN LÝ CHẤT LƯỢNG: {test_info['name'].upper()}", fmt_head)
    
    ws1.write('A3', "Đơn vị:", fmt_head)
    ws1.merge_range('B3:D3', "PHÒNG KHÁM ĐA KHOA QUỐC TẾ YERSIN", fmt_cell)
    ws1.write('E3', "Xét nghiệm:", fmt_head)
    ws1.merge_range('F3:G3', test_info['name'], fmt_cell)
    
    ws1.write('A4', "Khoa:", fmt_head)
    ws1.merge_range('B4:D4', "XÉT NGHIỆM", fmt_cell)
    ws1.write('E4', "Tháng :", fmt_head)
    ws1.merge_range('F4:G4', datetime.now().strftime("%m/%Y"), fmt_cell)
    
    ws1.write('A5', "Thời gian:", fmt_head)
    ws1.merge_range('B5:D5', f"{start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}", fmt_cell)
    ws1.write('E5', "Thiết bị:", fmt_head)
    ws1.merge_range('F5:G5', test_info.get('device', 'N/A'), fmt_cell)
    
    # === 2. SIX SIGMA TABLE ===
    ws1.merge_range('A7:G7', "SIX SIGMA & HIỆU NĂNG PHƯƠNG PHÁP", fmt_head)
    ws1.write_row('A8', ["Level", "Mean", "CV%", "Bias%", "Sigma", "QGI", "Ghi chú"], fmt_head)
    
    r = 8
    if sigma_data:
        for lvl, res in sigma_data.items():
            ws1.write(r, 0, f"Level {lvl}", fmt_cell)
            ws1.write(r, 1, res.get('mean', 0), fmt_num)
            ws1.write(r, 2, res.get('cv', 0), fmt_num)
            ws1.write(r, 3, res.get('bias', 0), fmt_num)
            ws1.write(r, 4, res.get('sigma', 0), fmt_num)
            qgi_val = res.get('qgi', 0)
            ws1.write(r, 5, round(qgi_val, 2) if isinstance(qgi_val, (int, float)) else qgi_val, fmt_num)
            ws1.write(r, 6, "", fmt_cell) # Cột ghi chú trống
            r += 1
    
# === 3. CHI TIẾT DỮ LIỆU IQC & WESTGARD ===
    r_start_iqc = r + 2
    ws1.merge_range(f'A{r_start_iqc}:G{r_start_iqc}', "CHI TIẾT DỮ LIỆU NỘI KIỂM (IQC) & VI PHẠM WESTGARD", fmt_head)
    
    headers_iqc = ["Ngày", "Lot", "Level", "Kết quả", "Z-Score", "Đánh giá (Lỗi)", "Hành động khắc phục"]
    ws1.write_row(r_start_iqc, 0, headers_iqc, fmt_head)
    
    row = r_start_iqc + 1
    df_calc = df_full_iqc.copy()
    
    if df_calc.empty:
         ws1.merge_range(row, 0, row, 6, "Không có dữ liệu Nội kiểm.", fmt_cell)
    else:
        # Lọc dữ liệu theo thời gian báo cáo trước
        df_calc['date_only'] = df_calc['date'].dt.date
        df_filtered = df_calc[(df_calc['date_only'] >= start_date) & (df_calc['date_only'] <= end_date)].copy()
        
        if df_filtered.empty:
            ws1.merge_range(row, 0, row, 6, "Không có dữ liệu trong khoảng thời gian báo cáo.", fmt_cell)
        else:
            # --- BƯỚC QUAN TRỌNG: PHÂN TÍCH WESTGARD ĐỂ HIỆN LỖI ---
            processed_data = []
            # Phân tích riêng cho từng Level để đảm bảo các quy tắc chuỗi (như 2-2s, 4-1s) chính xác
            for lvl in df_filtered['level'].unique():
                df_lvl = df_filtered[df_filtered['level'] == lvl].sort_values('date').copy()
                
                # Lấy Mean/SD từ dòng đầu tiên của Level đó (vì cùng 1 Lot trong báo cáo)
                m_val = df_lvl['target_mean'].iloc[0]
                s_val = df_lvl['target_sd'].iloc[0]
                
                # Gọi hàm phân tích (phải trùng tên với hàm dùng cho biểu đồ LJ)
                df_lvl_analyzed = get_westgard_violations(df_lvl, m_val, s_val)
                processed_data.append(df_lvl_analyzed)
            
            # Gộp lại và sắp xếp theo thời gian
            df_print = pd.concat(processed_data).sort_values(['date', 'level'])

            # === VÒNG LẶP GHI DỮ LIỆU ĐÃ PHÂN TÍCH ===
            for _, item in df_print.iterrows():
                report_date = pd.to_datetime(item['date'])
                ws1.write(row, 0, report_date.strftime('%d/%m/%Y'), fmt_cell)
                ws1.write(row, 1, item['lot_number'], fmt_cell)
                ws1.write(row, 2, item['level'], fmt_cell)
                ws1.write(row, 3, item['value'], fmt_num)
                
                # Tính lại Z-Score để in
                z = (item['value'] - item['target_mean']) / item['target_sd']
                ws1.write(row, 4, z, fmt_num)
                
                # Đánh giá lỗi (Cột 5)
                violation = item.get('Violation', "")
                if violation and violation != "":
                    error_label = violation
                    # Định dạng màu: Đỏ cho lỗi vi phạm dừng, Vàng cho lỗi cảnh báo (1-2s)
                    if any(rule in violation for rule in ["1-3s", "2-2s", "R-4s", "4-1s"]):
                        f_style = fmt_err
                    else:
                        f_style = fmt_warn # Lỗi 1-2s sẽ vào đây
                else:
                    error_label = "ĐẠT"
                    f_style = fmt_cell
                
                ws1.write(row, 5, error_label, f_style)
                
                # Hành động khắc phục (Cột 6) - Lấy từ cột 'note' như đã thống nhất
                # Chú ý: dùng .get('note') vì bạn nhập liệu vào cột note trên giao diện
                action_text = item.get('note', '') 
                ws1.write(row, 6, action_text, fmt_cell)
                
                row += 1
  
    # --- 4. CHÈN BIỂU ĐỒ (Giữ nguyên vị trí cột H để không đè dữ liệu) ---
    if img_lj:
        ws1.insert_image('H2', 'lj.png', {'image_data': img_lj, 'x_scale': 0.8, 'y_scale': 0.8})
        
    # --- 5. CHỮ KÝ ---
    sig_r = row + 4
    ws1.merge_range(sig_r, 1, sig_r, 4, "TRƯỞNG KHOA XÉT NGHIỆM", fmt_bold)
    ws1.merge_range(sig_r + 1, 1, sig_r + 1, 4, "(Ký và ghi rõ họ tên)", fmt_bold)
    # Đặt vùng in tự động cho Sheet 1
    ws1.print_area(0, 0, row + 2, 6)
    ws1.set_paper(9) # Giấy A4
    
    # === SHEET 2: NGOẠI KIỂM (EQA) ===
    ws2 = wb.add_worksheet("Ngoại Kiểm (EQA)")
    ws2.set_column('A:G', 15)

    # --- 1. Xử lý thời gian báo cáo (Tích hợp từ Đoạn 1) ---
    t_start = report_period[0].strftime('%d/%m/%Y') if hasattr(report_period[0], 'strftime') else str(report_period[0])
    t_end = report_period[1].strftime('%d/%m/%Y') if hasattr(report_period[1], 'strftime') else str(report_period[1])

    # --- 2. Tiêu đề chính và Thông tin hành chính (Tích hợp từ Đoạn 1) ---
    ws2.merge_range('A1:G1', "KẾT QUẢ NGOẠI KIỂM & CUSUM CỘNG DỒN", fmt_head)

    ws2.write('A3', "Đơn vị:", fmt_head)
    ws2.merge_range('B3:D3', "PHÒNG KHÁM ĐA KHOA QUỐC TẾ YERSIN", fmt_cell)
    ws2.write('E3', "Xét nghiệm:", fmt_head)
    ws2.merge_range('F3:G3', test_info['name'], fmt_cell)

    ws2.write('A4', "Khoa:", fmt_head)
    ws2.merge_range('B4:D4', "XÉT NGHIỆM", fmt_cell)
    ws2.write('E4', "Tháng :", fmt_head)
    ws2.merge_range('F4:G4', datetime.now().strftime("%m/%Y"), fmt_cell)

    ws2.write('A5', "Thời gian:", fmt_head)
    ws2.merge_range('B5:D5', f"{t_start} - {t_end}", fmt_cell)
    ws2.write('E5', "Thiết bị:", fmt_head)
    ws2.merge_range('F5:G5', test_info.get('device', 'N/A'), fmt_cell)

    # --- 3. Tiêu đề bảng dữ liệu (Bắt đầu từ dòng 7 để không đè lên thông tin hành chính) ---
    ws2.write_row('A7', ["Ngày", "Mã Mẫu", "PXN", "Ref", "SD Nhóm", "SDi (Z)", "CUSUM"], fmt_head)

    r2 = 7 # Bắt đầu ghi dữ liệu từ dòng 8 (index 7)
    if not df_eqa.empty:
        df_eqa_sort = df_eqa.sort_values('date').copy()
        
        # Tính toán Z-Score và CUSUM
        df_eqa_sort['Z-Score'] = (df_eqa_sort['lab_value'] - df_eqa_sort['ref_value']) / df_eqa_sort['sd_group']
        df_eqa_sort['CUSUM'] = df_eqa_sort['Z-Score'].cumsum()
        
        for _, row_eqa in df_eqa_sort.iterrows():
            ws2.write(r2, 0, pd.to_datetime(row_eqa['date']).strftime('%d/%m/%Y'), fmt_cell)
            ws2.write(r2, 1, row_eqa['sample_id'], fmt_cell)
            ws2.write(r2, 2, row_eqa['lab_value'], fmt_num)
            ws2.write(r2, 3, row_eqa['ref_value'], fmt_num)
            ws2.write(r2, 4, row_eqa['sd_group'], fmt_num)
            ws2.write(r2, 5, row_eqa['Z-Score'], fmt_num)
            ws2.write(r2, 6, row_eqa['CUSUM'], fmt_num)
            r2 += 1
            
        # Chèn biểu đồ CUSUM phía dưới bảng dữ liệu
        fig_cusum, violated = plot_cusum_chart(df_eqa_sort)
        if fig_cusum is not None:
            img_data = io.BytesIO()
            fig_cusum.savefig(img_data, format='png', bbox_inches='tight')
            img_data.seek(0)
            # Chèn cách bảng dữ liệu 2 dòng
            ws2.insert_image(f'A{r2 + 2}', 'cusum_chart.png', {'image_data': img_data})
    else:
         ws2.merge_range('A8:G8', "Không có dữ liệu Ngoại kiểm.", fmt_cell)


# === SHEET 3: MU & SIX SIGMA ===
    ws3 = wb.add_worksheet("MU & SixSigma")
    ws3.set_column('A:A', 15)
    ws3.set_column('B:I', 15)
    
    # 1. Tiêu đề chính (Dòng 1)
    ws3.merge_range('A1:H1', f"BÁO CÁO ĐỘ KHÔNG ĐẢM BẢO ĐO (MU): {test_info['name'].upper()}", fmt_head)
    
    # 2. Xử lý thời gian báo cáo
    t_start = report_period[0].strftime('%d/%m/%Y') if hasattr(report_period[0], 'strftime') else str(report_period[0])
    t_end = report_period[1].strftime('%d/%m/%Y') if hasattr(report_period[1], 'strftime') else str(report_period[1])

    # 3. Thông tin hành chính (Dòng 3 - 5)
    ws3.write('A3', "Đơn vị:", fmt_head)
    ws3.merge_range('B3:D3', "PHÒNG KHÁM ĐA KHOA QUỐC TẾ YERSIN", fmt_cell)
    ws3.write('E3', "Xét nghiệm:", fmt_head)
    ws3.merge_range('F3:H3', test_info['name'], fmt_cell)
    
    ws3.write('A4', "Khoa:", fmt_head)
    ws3.merge_range('B4:D4', "XÉT NGHIỆM", fmt_cell)
    ws3.write('E4', "Tháng :", fmt_head)
    ws3.merge_range('F4:H4', datetime.now().strftime("%m/%Y"), fmt_cell)
    
    ws3.write('A5', "Thời gian:", fmt_head)
    ws3.merge_range('B5:D5', f"{t_start} - {t_end}", fmt_cell)
    ws3.write('E5', "Thiết bị:", fmt_head)
    ws3.merge_range('F5:H5', test_info.get('device', 'N/A'), fmt_cell)

    # 4. Bảng Kết quả thực tế (Dòng 7 - 10)
    ws3.merge_range('A7:H7', "KẾT QUẢ THỰC TẾ & ĐÁNH GIÁ HIỆU NĂNG", fmt_head)
    ws3.write('A8', 'Level', fmt_head)
    ws3.write_row('B8', ['Mean', 'CV%', 'Bias%', 'Sigma', 'Ue (k=2)', 'Ue (%)', 'Đánh giá BV'], fmt_head)
    
    m_min, m_des, m_opt = mau_limits
    
    r3 = 8 # Dòng index bắt đầu ghi Level 1 (Dòng 9 trong Excel)
    for lvl in [1, 2]:
        res_sigma = sigma_data.get(lvl, {}) if sigma_data else {}
        res_mu = mu_data.get(lvl, {}) if mu_data else {}
        
        mean_val = res_sigma.get('mean', 0)
        ue_abs = res_mu.get('ue', 0)
        ue_pct = (ue_abs / mean_val) * 100 if mean_val > 0 else 0
        
        if ue_pct <= 0: status = "N/A"
        elif ue_pct <= m_opt: status = "Tối ưu"
        elif ue_pct <= m_des: status = "Mong muốn"
        elif ue_pct <= m_min: status = "Tối thiểu"
        else: status = "Không đạt"

        ws3.write(r3, 0, f"Level {lvl}", fmt_cell)
        ws3.write(r3, 1, mean_val, fmt_num)
        ws3.write(r3, 2, res_sigma.get('cv', 0), fmt_num)
        ws3.write(r3, 3, res_sigma.get('bias', 0), fmt_num)
        ws3.write(r3, 4, res_sigma.get('sigma', 0), fmt_num)
        ws3.write(r3, 5, ue_abs, fmt_num)
        ws3.write(r3, 6, ue_pct, fmt_num)
        ws3.write(r3, 7, status, fmt_cell)
        r3 += 1
    if img_sigma:
        ws3.insert_image('F12', 'sigma.png', {'image_data': img_sigma, 'x_scale': 0.8, 'y_scale': 0.8})

    # 5. Bảng Mục tiêu Đánh giá (Nằm dưới bảng thực tế)
    target_row = r3 + 2
    ws3.merge_range(target_row, 0, target_row, 3, "MỤC TIÊU ĐỘ KHÔNG ĐẢM BẢO ĐO CHO PHÉP (MAU)", fmt_head)
    ws3.write_row(target_row + 1, 0, ["Mức độ", "Hệ số", "Giới hạn (%)", "Trạng thái"], fmt_head)
    ws3.write_row(target_row + 2, 0, ["Tối ưu", "0.25", m_opt, "Rất tốt"], fmt_cell)
    ws3.write_row(target_row + 3, 0, ["Mong muốn", "0.50", m_des, "Đạt"], fmt_cell)
    ws3.write_row(target_row + 4, 0, ["Tối thiểu", "0.75", m_min, "Chấp nhận"], fmt_cell)
    # 6. PHẦN CHỮ KÝ (Nằm dưới bảng mục tiêu hoặc dưới ảnh nếu có)
    # Tính toán dòng bắt đầu cho chữ ký (cách bảng mục tiêu khoảng 2 dòng hoặc sau ảnh)
    sig_row = target_row + 17 
    
    # Định dạng chữ ký (Căn giữa, in đậm)
    fmt_sig_label = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
    fmt_sig_sub = wb.add_format({'italic': True, 'align': 'center', 'valign': 'vcenter'})

    # Cột Người lập biểu (Cột A -> C)
    ws3.merge_range(sig_row, 0, sig_row, 2, "NGƯỜI LẬP BÁO CÁO", fmt_sig_label)
    ws3.merge_range(sig_row + 1, 0, sig_row + 1, 2, "(Ký và ghi rõ họ tên)", fmt_sig_sub)

    # Cột Trưởng khoa (Cột E -> H)
    ws3.merge_range(sig_row, 4, sig_row, 7, "TRƯỞNG KHOA XÉT NGHIỆM", fmt_sig_label)
    ws3.merge_range(sig_row + 1, 4, sig_row + 1, 7, "(Ký và ghi rõ họ tên)", fmt_sig_sub)
    wb.close()
    return output.getvalue()

# --- SIDEBAR: CONTROL PANEL ---

st.sidebar.markdown("---")
st.title("🔬 Hệ Thống QLCL Xét Nghiệm ")
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
# 3. QUẢN LÝ LOTS (NÂNG CẤP: Tách biệt chọn L1 và L2)
st.sidebar.markdown("---")
st.sidebar.subheader("📦 Cấu hình Lot Đang Chạy")

all_lots = db.get_lots_for_test(current_test['id'])
lots_l1 = all_lots[all_lots['level'] == 1]
lots_l2 = all_lots[all_lots['level'] == 2]

# Tạo dict để selectbox
opts_l1 = {f"{r['lot_number']} (Exp:{r['expiry_date']})": r.to_dict() for _, r in lots_l1.iterrows()}
opts_l2 = {f"{r['lot_number']} (Exp:{r['expiry_date']})": r.to_dict() for _, r in lots_l2.iterrows()}

# Selectbox riêng biệt
s_l1 = st.sidebar.selectbox("Lot Level 1:", ["-- Chọn L1 --"] + list(opts_l1.keys()))
s_l2 = st.sidebar.selectbox("Lot Level 2:", ["-- Chọn L2 --"] + list(opts_l2.keys()))

cur_lot_l1 = opts_l1[s_l1] if s_l1 != "-- Chọn L1 --" else None
cur_lot_l2 = opts_l2[s_l2] if s_l2 != "-- Chọn L2 --" else None

# Form thêm Lot mới (Linh hoạt: cho phép thêm lẻ)
with st.sidebar.expander("➕ Thêm Lot Mới (Tùy chọn)"):
    with st.form("add_lot_flex"):
        st.write("Thêm Lot mới (Nhập cái nào lưu cái đó)")
        mt = st.text_input("Phương pháp/Máy", value=current_test['device'])
        
        c1, c2 = st.columns(2)
        with c1: 
            st.caption("Level 1")
            ln1 = st.text_input("Lot L1"); m1 = st.number_input("Mean 1", format="%.3f"); sd1 = st.number_input("SD 1", format="%.3f")
            ed1 = st.date_input("Hạn L1")
        with c2:
            st.caption("Level 2")
            ln2 = st.text_input("Lot L2"); m2 = st.number_input("Mean 2", format="%.3f"); sd2 = st.number_input("SD 2", format="%.3f")
            ed2 = st.date_input("Hạn L2")
            
        if st.form_submit_button("Lưu Lot"):
            if ln1: db.add_lot(current_test['id'], ln1, 1, mt, ed1, m1, sd1)
            if ln2: db.add_lot(current_test['id'], ln2, 2, mt, ed2, m2, sd2)
            st.success("Đã lưu!"); st.rerun()
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

tabs = st.tabs(["1. Nhập IQC", "2. Biểu đồ LJ", "3. Ngoại kiểm (EQA)", "4. Độ KĐB (MU)", "5. Six Sigma & Báo cáo", "6. Quản trị", "6. Import dữ liệu"])

# === TAB 1: NHẬP IQC ===
with tabs[0]:
    c_in, c_dat = st.columns([1, 2])
    with c_in:
        st.subheader("Nhập Kết Quả Hàng Ngày")
        if not cur_lot_l1 and not cur_lot_l2:
            st.error("Vui lòng chọn ít nhất 1 Lot ở Sidebar để nhập liệu.")
        else:
            with st.form("iqc_entry"):
                d_in = st.date_input("Ngày chạy", datetime.now())
                note = st.text_input("Ghi chú")
                
                v1, v2 = None, None
                if cur_lot_l1: 
                    st.markdown(f"**L1: {cur_lot_l1['lot_number']}** (Target: {cur_lot_l1['mean']})")
                    v1 = st.number_input("Kết quả L1", format="%.4f")
                
                if cur_lot_l2:
                    st.markdown(f"**L2: {cur_lot_l2['lot_number']}** (Target: {cur_lot_l2['mean']})")
                    v2 = st.number_input("Kết quả L2", format="%.4f")
                
                if st.form_submit_button("Lưu Kết Quả"):
                    if cur_lot_l1 and v1: db.add_iqc(cur_lot_l1['id'], d_in, 1, v1, note)
                    if cur_lot_l2 and v2: db.add_iqc(cur_lot_l2['id'], d_in, 2, v2, note)
                    st.success("Đã lưu!"); st.rerun()

with c_dat:
        st.subheader("Lịch sử dữ liệu (Lot hiện tại) & Chỉnh sửa")
        
        # --- CẬP NHẬT: Dùng data_editor cho cả 2 Level ---

        if cur_lot_l1:
            st.caption(f"Dữ liệu L1 ({cur_lot_l1['lot_number']})")
            df_l1 = db.get_iqc_data_by_lot(cur_lot_l1['id'])
            
            edited_df_l1 = st.data_editor(
                df_l1[['id', 'date', 'value', 'note']].sort_values('date', ascending=False),
                column_config={
                    "date": st.column_config.DatetimeColumn("Ngày", format="YYYY-MM-DD", required=True),
                    "value": st.column_config.NumberColumn("Kết quả", format="%.4f", required=True),
                    "note": st.column_config.TextColumn("Ghi chú"),
                    "id": st.column_config.NumberColumn("ID", disabled=True),
                },
                num_rows="dynamic",
                key="editor_l1",
                use_container_width=True
            )
            
            # Xử lý các thay đổi (Chỉnh sửa/Xóa)
            if st.button("Lưu thay đổi L1", key="save_l1_btn"):
                # 1. Tìm các hàng bị xóa
                deleted_rows_l1 = df_l1[~df_l1['id'].isin(edited_df_l1['id'])]
                for iqc_id in deleted_rows_l1['id']:
                    db.delete_iqc_data(iqc_id)
                
                # 2. Tìm và cập nhật các hàng được chỉnh sửa
                for _, row in edited_df_l1.iterrows():
                    original_row = df_l1[df_l1['id'] == row['id']].iloc[0]
                    # Chỉ update nếu có thay đổi
                    if (row['date'] != original_row['date'] or 
                        row['value'] != original_row['value'] or 
                        row['note'] != original_row['note']):
                        
                        db.update_iqc_data(row['id'], row['date'], 1, row['value'], row['note'])
                
                st.success("Đã cập nhật dữ liệu L1!")
                st.rerun()


        if cur_lot_l2:
            st.caption(f"Dữ liệu L2 ({cur_lot_l2['lot_number']})")
            df_l2 = db.get_iqc_data_by_lot(cur_lot_l2['id'])

            edited_df_l2 = st.data_editor(
                df_l2[['id', 'date', 'value', 'note']].sort_values('date', ascending=False),
                column_config={
                    "date": st.column_config.DatetimeColumn("Ngày", format="YYYY-MM-DD", required=True),
                    "value": st.column_config.NumberColumn("Kết quả", format="%.4f", required=True),
                    "note": st.column_config.TextColumn("Ghi chú"),
                    "id": st.column_config.NumberColumn("ID", disabled=True),
                },
                num_rows="dynamic",
                key="editor_l2",
                use_container_width=True
            )
            
            if st.button("Lưu thay đổi L2", key="save_l2_btn"):
                # 1. Tìm các hàng bị xóa
                deleted_rows_l2 = df_l2[~df_l2['id'].isin(edited_df_l2['id'])]
                for iqc_id in deleted_rows_l2['id']:
                    db.delete_iqc_data(iqc_id)
                
                # 2. Tìm và cập nhật các hàng được chỉnh sửa
                for _, row in edited_df_l2.iterrows():
                    original_row = df_l2[df_l2['id'] == row['id']].iloc[0]
                    # Chỉ update nếu có thay đổi
                    if (row['date'] != original_row['date'] or 
                        row['value'] != original_row['value'] or 
                        row['note'] != original_row['note']):
                        
                        db.update_iqc_data(row['id'], row['date'], 2, row['value'], row['note'])
                st.session_state['last_update'] = datetime.now()
                st.success("Đã cập nhật dữ liệu L2!")
                st.rerun()

# === TAB 2: BIỂU ĐỒ LJ & NHẬT KÝ VI PHẠM (Tự động chèn Timestamp) ===
with tabs[1]:
    import sqlite3
    from datetime import datetime

    col_opt, col_chart = st.columns([1, 4])
    with col_opt:
        view_mode = st.radio("Chế độ xem:", ["Toàn bộ lịch sử (Nối Lot)", "Chỉ Lot đang chọn"])
    
    # Lấy dữ liệu IQC liên tục
    df_all = db.get_iqc_data_continuous(current_test['id'], max_months=12)
    
    if not df_all.empty:
        # 1. Lọc dữ liệu theo chế độ xem
        if view_mode == "Chỉ Lot đang chọn":
            active_ids = []
            if cur_lot_l1: active_ids.append(cur_lot_l1['id'])
            if cur_lot_l2: active_ids.append(cur_lot_l2['id'])
            df_plot = df_all[df_all['lot_id'].isin(active_ids)].copy()
        else:
            df_plot = df_all.copy()

        # 2. Vẽ biểu đồ LJ
        st.pyplot(plot_levey_jennings(df_plot, f"Biểu đồ Levey-Jennings ({view_mode})"))
        
        # 3. Cảnh báo Westgard nhanh
        st.markdown("#### ⚠️ Cảnh báo Westgard (Dữ liệu hiển thị)")
        violations_summary = {}
        for lvl in [1, 2]:
            sub = df_plot[df_plot['level'] == lvl].copy()
            if not sub.empty:
                v = evaluate_westgard_series(sub)
                if v: violations_summary[f"Level {lvl}"] = list(set(v))

        if violations_summary:
            for k, v in violations_summary.items(): 
                st.error(f"**{k}**: {', '.join(v)}")
        else:
            st.success("Không phát hiện vi phạm quy tắc dừng (Rejection Rules).")

        st.divider()

        # 4. VÒNG LẶP XỬ LÝ NHẬT KÝ VI PHẠM CHO CẢ 2 LEVEL
        levels_config = [
            {"id": 1, "name": "Level 1", "lot": cur_lot_l1},
            {"id": 2, "name": "Level 2", "lot": cur_lot_l2}
        ]

        for lvl in levels_config:
            l_id = lvl["id"]
            l_name = lvl["name"]
            l_lot = lvl["lot"]
            
            if l_lot:
                df_lvl = db.get_iqc_data_by_lot(l_lot['id'])
                if not df_lvl.empty:
                    # Phân tích Westgard chi tiết
                    df_analyzed = get_westgard_violations(df_lvl, l_lot['mean'], l_lot['sd'])
                    df_err_only = df_analyzed[df_analyzed['Violation'] != ""].copy()
                    
                    st.markdown(f"#### 📝 Nhật ký Vi phạm & Xử lý ({l_name})")
                    
                    if not df_err_only.empty:
                        # Chuẩn bị dữ liệu cho Editor
                        df_editor = df_err_only[['id', 'date', 'value', 'Violation', 'note']].copy()
                        df_editor['date'] = pd.to_datetime(df_editor['date']).dt.strftime('%d/%m/%Y %H:%M')
                        df_editor['id'] = df_editor['id'].astype(str)

                        edited_df = st.data_editor(
                            df_editor.rename(columns={
                                'date': 'Ngày giờ lỗi',
                                'value': 'Kết quả',
                                'Violation': 'Lỗi Westgard',
                                'note': 'Hành động khắc phục (Note nội dung xử lý tại đây)'
                            }),
                            column_config={
                                "id": None,
                                "Hành động khắc phục (Note nội dung xử lý tại đây)": st.column_config.TextColumn(width="large")
                            },
                            disabled=["Ngày giờ lỗi", "Kết quả", "Lỗi Westgard"], 
                            key=f"editor_lvl_{l_id}",
                            hide_index=True,
                            use_container_width=True
                        )

                        # Nút lưu có tự động điền thời gian xử lý
                        if st.button(f"💾 Lưu & Đóng dấu thời gian xử lý {l_name}", key=f"btn_save_{l_id}"):
                            try:
                                conn = sqlite3.connect("lab_data.db")
                                cursor = conn.cursor()
                                now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
                                
                                for _, row in edited_df.iterrows():
                                    raw_note = row['Hành động khắc phục (Note nội dung xử lý tại đây)']
                                    if raw_note:
                                        # Nếu nội dung đã có dấu thời gian thì không chèn thêm, tránh trùng lặp
                                        if " - [Xử lý lúc:" in raw_note:
                                            final_note = raw_note
                                        else:
                                            final_note = f"{raw_note} - [Xử lý lúc: {now_str}]"
                                        
                                        cursor.execute(
                                            "UPDATE iqc_results SET note = ? WHERE id = ?", 
                                            (final_note, row['id'])
                                        )
                                conn.commit()
                                conn.close()
                                st.success(f"✅ Đã lưu hành động cho {l_name} lúc {now_str}!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Lỗi: {e}")
                    else:
                        st.info(f"✅ {l_name}: Không có vi phạm cần xử lý.")
            st.write("") 
    else:
        st.info("Chưa có dữ liệu nội kiểm cho xét nghiệm này.")
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
    
    with c1:
        st.subheader("Nhập kết quả EQA")
        with st.form("eqa_in"):
            ed = st.date_input("Ngày mẫu", key="eqa_date")
            el = st.number_input("Giá trị PXN", format="%.4f", key="eqa_lab_value")
            er = st.number_input("Giá trị Tham chiếu (Nhóm)", format="%.4f", key="eqa_ref_value")
            es = st.number_input("SD Nhóm (Group SD)", value=1.0, format="%.4f", key="eqa_sd_group")
            en = st.text_input("Mã mẫu", key="eqa_sample_id")
            
            if st.form_submit_button("Lưu EQA"):
                if es > 0:
                    db.add_eqa(current_test['id'], ed, el, er, es, en)
                    st.success("Đã lưu kết quả EQA!")
                    st.rerun()
                else:
                    st.error("SD Nhóm phải lớn hơn 0")

    # --- PHẦN 2: BẢNG DỮ LIỆU CÓ CHỨC NĂNG CHỈNH SỬA & XÓA ---
    with c2:
        st.subheader("Dữ liệu EQA (Chỉnh sửa trực tiếp)")

        if not df_display.empty:
            
            # 2. CHUẨN BỊ DATAFRAME CHO EDITOR
            df_edit = df_display[['id', 'date', 'sample_id', 'lab_value', 'ref_value', 'sd_group', 'Z-Score', 'CUSUM']].copy()
            df_edit.columns = ['ID', 'Ngày', 'Mã Mẫu', 'PXN', 'Ref', 'SD Nhóm', 'Z-Score', 'CUSUM']
            df_edit = df_edit.set_index('ID')
            df_edit.insert(0, 'Xóa', False) # Thêm cột xóa vào vị trí đầu tiên
            
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
                    "Xóa": st.column_config.CheckboxColumn(default=False)
                },
                hide_index=False,
                use_container_width=True,
            )

            # --- PHẦN XỬ LÝ HÀNH ĐỘNG CẬP NHẬT/XÓA ---

        # 4. XỬ LÝ HÀNH ĐỘNG (NÚT ÁP DỤNG)
        if st.button("Xóa Dữ liệu"):
            
            # 1. Lấy dữ liệu ID đã bị đánh dấu xóa
            deleted_ids = edited_df[edited_df['Xóa'] == True].index.tolist()
            
            # 2. Lấy dữ liệu đã chỉnh sửa
            updates = st.session_state.get("eqa_data_editor", {}).get("edited_rows", {})
            update_count = 0
            
            # 3. Thực hiện CẬP NHẬT (trước khi xóa)
            for row_index_str, changes in updates.items():
                try:
                    # Lấy ID thực tế từ index của edited_df
                    eqa_id = edited_df.index[int(row_index_str)]
                except IndexError:
                    continue # Bỏ qua nếu lỗi index

                update_data = {}
                # Ánh xạ lại tên cột: Tên hiển thị -> Tên DB
                if 'PXN' in changes: update_data['lab_value'] = changes['PXN']
                if 'Ref' in changes: update_data['ref_value'] = changes['Ref']
                if 'SD Nhóm' in changes: update_data['sd_group'] = changes['SD Nhóm']
                if 'Mã Mẫu' in changes: update_data['sample_id'] = changes['Mã Mẫu']
                if 'Ngày' in changes: update_data['date'] = changes['Ngày']
                
                if update_data:
                    # Chỉ cập nhật nếu bản ghi đó KHÔNG bị đánh dấu xóa
                    if eqa_id not in deleted_ids:
                        if db.update_eqa(eqa_id, update_data):
                            update_count += 1
            
            # 4. Thực hiện XÓA
            deleted_count = 0
            if deleted_ids:
                for eqa_id in deleted_ids:
                    if db.delete_eqa(eqa_id):
                        deleted_count += 1
            
            # 5. Báo cáo kết quả và tải lại
            if deleted_count > 0 or update_count > 0:
                st.success(f"✅ Đã cập nhật {update_count} bản ghi và xóa {deleted_count} bản ghi.")
                st.rerun()
            else:
                st.info("Không có thay đổi nào cần áp dụng.")

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
        
        if is_violated:
            st.error("⚠️ CẢNH BÁO: Đường CUSUM cắt V-Mask! Có dấu hiệu sai số hệ thống (Shift/Trend).")
        else:
            st.success("✅ Hệ thống ổn định (CUSUM nằm trong V-Mask).")
            
    elif not df_eqa.empty:
        st.warning("Cần ít nhất 2 điểm dữ liệu EQA để vẽ biểu đồ CUSUM.")

# === TAB 4: ĐỘ KĐB ĐO (MU) & BIOLOGICAL VARIATION ===
with tabs[3]: # Đảm bảo tabs[3] tương ứng với vị trí MU trong st.tabs([...])
    st.header("4. Độ Không Đảm Bảo Đo & Đánh giá Hiệu năng")
    
    # --- KIỂM TRA LOT TRƯỚC KHI TÍNH TOÁN ---
    # Hệ thống mới tách biệt L1 và L2, kiểm tra xem có ít nhất 1 Lot được cấu hình không
    if cur_lot_l1 is None and cur_lot_l2 is None:
        st.warning("⚠️ Vui lòng cấu hình Lot (Lô) QC ở Sidebar để lấy thông số SD nhà sản xuất.")
    else:
        # 1. BỘ LỌC THỜI GIAN & INPUT CƠ BẢN
        with st.expander("⚙️ Cài đặt Thông số & Thời gian tính toán", expanded=True):
            col_time, col_bv = st.columns(2)
            with col_time:
                st.subheader("1. Khoảng thời gian tính toán")
                today = datetime.now().date()
                start_default = today.replace(day=1) # Mặc định ngày 1 tháng này
                d_start = st.date_input("Từ ngày", start_default, key="mu_start")
                d_end = st.date_input("Đến ngày", today, key="mu_end")

            with col_bv:
                st.subheader("2. Thông số Biến thiên Sinh học")
                
                # Lấy dữ liệu đã có từ database
                test_id = current_test.get('id', 'default')
                db_cvi = float(current_test.get('cvi', 0.0))
                db_cvg = float(current_test.get('cvg', 0.0))
                
                cvi_in = st.number_input("CVi (Intra-individual)", value=db_cvi, format="%.2f", key=f"mu_cvi_{test_id}")
                cvg_in = st.number_input("CVg (Inter-individual)", value=db_cvg, format="%.2f", key=f"mu_cvg_{test_id}")
                
                if cvi_in > 0:
                    # Tính toán 3 mức MAU (%)
                    # 1. Tối thiểu (Hệ số 0.75)
                    mau_min = 0.75 * cvi_in + 1.65 * (0.375 * np.sqrt(cvi_in**2 + cvg_in**2))
                    # 2. Mong muốn (Hệ số 0.5) - Đây là mức phổ biến nhất
                    mau_des = 0.5 * cvi_in + 1.65 * (0.25 * np.sqrt(cvi_in**2 + cvg_in**2))
                    # 3. Tối ưu (Hệ số 0.25)
                    mau_opt = 0.25 * cvi_in + 1.65 * (0.125 * np.sqrt(cvi_in**2 + cvg_in**2))

                    # Hiển thị bảng so sánh nhanh
                    st.success("🎯 Giới hạn MU cho phép (MAU):")
                    cols = st.columns(3)
                    cols[0].metric("Tối thiểu", f"{mau_min:.2f}%")
                    cols[1].metric("Mong muốn", f"{mau_des:.2f}%")
                    cols[2].metric("Tối ưu", f"{mau_opt:.2f}%")
                    
                    # Gán giá trị mục tiêu để so sánh ở phần kết quả phía dưới (thường dùng mức Mong muốn)
                    tea_limit = mau_des 
                else:
                    tea_limit = float(current_test.get('tea', 10.0))
                    st.warning(f"Chưa có CVi. Sử dụng TEa cài đặt ({tea_limit}%) làm MAU.")
                
                # --- LOGIC TÍNH TEa (MAU) ---
                if cvi_in > 0:
                    # Công thức tính TEa mong muốn dựa trên Biological Variation
                    # TEa = 0.5 * CVi + 1.65 * (0.25 * sqrt(CVi² + CVg²))
                    tea_des = 0.5 * cvi_in + 1.65 * (0.25 * np.sqrt(cvi_in**2 + cvg_in**2))
                    st.success(f"✅ Đang dùng CVi/CVg từ cài đặt để tính MAU.")
                    st.info(f"**TEa Mong muốn (BV): {tea_des:.2f}%**")
                else:
                    # Fallback về TEa cố định nếu không có dữ liệu BV
                    tea_des = float(current_test.get('tea', 0.0))
                    st.warning(f"⚠️ Chưa có dữ liệu CVi. Sử dụng TEa mặc định ({tea_des}%) làm giới hạn MAU.")
        # 2. TÍNH TOÁN
        st.markdown("---")
        st.subheader("3. Kết quả Tính toán")

        # --- LẤY DỮ LIỆU IQC LIÊN TỤC ---
        try:
            # Sử dụng hàm get_iqc_data_continuous đã cập nhật trong db_module
            df_all_mu = db.get_iqc_data_continuous(current_test['id'])
            if not df_all_mu.empty:
                df_all_mu['date'] = pd.to_datetime(df_all_mu['date'])
                mask = (df_all_mu['date'].dt.date >= d_start) & (df_all_mu['date'].dt.date <= d_end)
                df_mu_filtered = df_all_mu[mask]
            else:
                df_mu_filtered = pd.DataFrame()
        except Exception as e:
            st.error(f"Lỗi lấy dữ liệu IQC: {e}")
            df_mu_filtered = pd.DataFrame()

        # --- LẤY %BIAS TỪ EQA ---
        df_eqa_mu = db.get_eqa_data(current_test['id'])
        bias_pct = 0.0
        if not df_eqa_mu.empty:
            last_eqa = df_eqa_mu.iloc[-1]
            if last_eqa['ref_value'] != 0:
                bias_pct = abs((last_eqa['lab_value'] - last_eqa['ref_value']) / last_eqa['ref_value']) * 100
        
        # --- HÀM TÍNH CHI TIẾT ---
        def calculate_mu_level_logic(df, level_num, bias_p, mau_limit, lot_sd):
            if df.empty or 'level' not in df.columns:
                return None 
            
            df_lvl = df[df['level'] == level_num]
            
            if not df_lvl.empty and len(df_lvl) >= 2:
                mean_calc, sd_calc, cv_calc = get_stats_real(df_lvl)
                u_prec = sd_calc 
            else:
                # Fallback về SD của Lot nếu chưa đủ dữ liệu
                mean_calc = 0
                sd_calc = 0
                cv_calc = 0
                u_prec = lot_sd if lot_sd else 0
            
            # Tính u_bias dựa trên Mean thực tế
            bias_abs_val = (bias_p / 100) * mean_calc if mean_calc else 0
            u_bias = bias_abs_val 
            
            # Tính Độ KĐB đo tổng hợp (uc) và mở rộng (Ue)
            uc = np.sqrt(u_prec**2 + u_bias**2)
            ue = uc * 2 # k=2
            
            # MAU (Maximum Allowable Uncertainty)
            mau_abs = (mau_limit / 100) * mean_calc if mean_calc else 0
            pass_mau = ue <= mau_abs if mau_abs > 0 else False
            
            return {
                "n": len(df_lvl), "mean": mean_calc, "sd": sd_calc, "cv": cv_calc,
                "u_prec": u_prec, "u_bias": u_bias, "uc": uc, "ue": ue, 
                "mau_abs": mau_abs, "pass": pass_mau
            }

        # --- HÀM ĐÁNH GIÁ HIỆU NĂNG ---
        def get_performance_status(ue_pct, m_min, m_des, m_opt):
            if ue_pct <= m_opt:
                return "🌟 TỐI ƯU (Optimal)", "green"
            elif ue_pct <= m_des:
                return "✅ MONG MUỐN (Desirable)", "blue"
            elif ue_pct <= m_min:
                return "⚠️ TỐI THIỂU (Minimum)", "orange"
            else:
                return "❌ KHÔNG ĐẠT", "red"

        # --- HIỂN THỊ KẾT QUẢ ---
        c1, c2 = st.columns(2)

        # XỬ LÝ LEVEL 1
        with c1:
            st.markdown("#### 🔵 Level 1")
            if cur_lot_l1 is not None:
                target_sd1 = cur_lot_l1['sd'] if 'sd' in cur_lot_l1 else 0
                res_l1 = calculate_mu_level_logic(df_mu_filtered, 1, bias_pct, tea_limit, target_sd1)
                
                if res_l1 and res_l1['mean'] > 0:
                    ue_pct = (res_l1['ue'] / res_l1['mean']) * 100
                    status_text, color = get_performance_status(ue_pct, mau_min, mau_des, mau_opt)
                    
                    st.metric("Ue (k=2)", f"{res_l1['ue']:.4f}", f"{ue_pct:.2f}%")
                    st.markdown(f"Đánh giá: :{color}[**{status_text}**]")
                    
                    with st.expander("Chi tiết mục tiêu"):
                        st.write(f"- Tối ưu: ≤ {mau_opt:.2f}%")
                        st.write(f"- Mong muốn: ≤ {mau_des:.2f}%")
                        st.write(f"- Tối thiểu: ≤ {mau_min:.2f}%")
                else:
                    st.warning("Không đủ dữ liệu Level 1.")

        # XỬ LÝ LEVEL 2
        with c2:
            st.markdown("#### 🟠 Level 2")
            if cur_lot_l2 is not None:
                target_sd2 = cur_lot_l2['sd'] if 'sd' in cur_lot_l2 else 0
                res_l2 = calculate_mu_level_logic(df_mu_filtered, 2, bias_pct, tea_limit, target_sd2)
                
                if res_l2 and res_l2['mean'] > 0:
                    ue_pct_l2 = (res_l2['ue'] / res_l2['mean']) * 100
                    status_text_l2, color_l2 = get_performance_status(ue_pct_l2, mau_min, mau_des, mau_opt)
                    
                    st.metric("Ue (k=2)", f"{res_l2['ue']:.4f}", f"{ue_pct_l2:.2f}%")
                    st.markdown(f"Đánh giá: :{color_l2}[**{status_text_l2}**]")
                    
                    with st.expander("Chi tiết mục tiêu"):
                        st.write(f"- Tối ưu: ≤ {mau_opt:.2f}%")
                        st.write(f"- Mong muốn: ≤ {mau_des:.2f}%")
                        st.write(f"- Tối thiểu: ≤ {mau_min:.2f}%")
                else:
                    st.warning("Không đủ dữ liệu Level 2.")

# === TAB 5: SIX SIGMA & BÁO CÁO ===
with tabs[4]:
    st.header("5. Six Sigma, QGI & Báo Cáo tổng hợp")

    # 1. BỘ LỌC THỜI GIAN
    with st.expander("📅 Chọn khoảng thời gian báo cáo", expanded=True):
        c_d1, c_d2 = st.columns(2)
        start_d = c_d1.date_input("Từ ngày", datetime.now().replace(day=1))
        end_d = c_d2.date_input("Đến ngày", datetime.now())

    # 2. TÍNH TOÁN DATA
    df_all = db.get_iqc_data_continuous(current_test['id'])
    df_eqa = db.get_eqa_data(current_test['id'])
    
    # Lọc data theo ngày
    if not df_all.empty:
        # Chuyển đổi cột 'date' sang datetime nếu chưa phải
        df_all['date'] = pd.to_datetime(df_all['date'])
        df_all = df_all[(df_all['date'].dt.date >= start_d) & (df_all['date'].dt.date <= end_d)]
    
    tea = current_test['tea']
    
    # Lấy Bias từ EQA gần nhất trong khoảng thời gian (hoặc gần nhất overall)
    bias_pct = 0.0
    if not df_eqa.empty:
        last = df_eqa.iloc[-1]
        if last['ref_value'] != 0:
            bias_pct = abs((last['lab_value'] - last['ref_value'])/last['ref_value'])*100
    
# 2. TÍNH TOÁN DATA
    # Lấy TOÀN BỘ dữ liệu (Không lọc ngày ngay lập tức)
    df_full_history = db.get_iqc_data_continuous(current_test['id'])
    
    # Chuyển đổi cột 'date' sang datetime
    if not df_full_history.empty:
        df_full_history['date'] = pd.to_datetime(df_full_history['date'])

    # Tạo một bản sao ĐÃ LỌC để dùng cho tính toán Sigma/MU và hiển thị Dashboard
    if not df_full_history.empty:
        mask = (df_full_history['date'].dt.date >= start_d) & (df_full_history['date'].dt.date <= end_d)
        df_filtered = df_full_history[mask].copy()
    else:
        df_filtered = pd.DataFrame()

    df_eqa = db.get_eqa_data(current_test['id'])
    tea = current_test['tea']
    
    # Lấy Bias từ EQA gần nhất
    bias_pct = 0.0
    if not df_eqa.empty:
        last = df_eqa.iloc[-1]
        if last['ref_value'] != 0:
            bias_pct = abs((last['lab_value'] - last['ref_value'])/last['ref_value'])*100
    
    # --- TÍNH TOÁN SIGMA DỰA TRÊN DỮ LIỆU ĐÃ LỌC (df_filtered) ---
    sigma_results = {}
    sigma_plot_data = []
    
    c1, c2 = st.columns(2)
    
    for lvl in [1, 2]:
        df_lvl = df_filtered[df_filtered['level'] == lvl] if not df_filtered.empty else pd.DataFrame()
        
        cv = 0.0
        mean_val = 0.0
        n_count = len(df_lvl)
        
        if n_count >= 2:
            mean_val, sd_val, cv = get_stats_real(df_lvl)
            
        sigma = (tea - bias_pct) / cv if cv > 0 else 0
        qgi, qgi_reason = calculate_qgi(bias_pct, cv)
        
        sigma_results[lvl] = {
            'cv': round(cv, 2), 
            'bias': round(bias_pct, 2), 
            'sigma': round(sigma, 2), 
            'qgi': round(qgi, 2), 
            'reason': qgi_reason,
            'mean': round(mean_val, 4),
            'n': n_count,
            'sd': round(sd_val, 4) if n_count >= 2 else 0
        }
        
        if cv > 0:
            sigma_plot_data.append({'label': f"L{lvl}", 'bias': bias_pct, 'cv': cv})
        
# --- PHẦN HIỂN THỊ UI NÂNG CAO ---
        with c1 if lvl == 1 else c2:
            # Tạo khung bao quanh bằng st.container
            with st.container(border=True):
                st.markdown(f"### 🎯 Level {lvl}")
                
                # Hàng 1: Hiển thị Sigma lớn
                # Màu sắc: Sigma > 6 (Xanh dương), > 4 (Xanh lá), > 3 (Vàng), < 3 (Đỏ)
                if sigma >= 6:
                    st.success(f"**SIX SIGMA: {sigma:.2f} (Thế giới - World Class)**")
                elif sigma >= 4:
                    st.info(f"**SIX SIGMA: {sigma:.2f} (Tốt - Excellent)**")
                elif sigma >= 3:
                    st.warning(f"**SIX SIGMA: {sigma:.2f} (Tạm đạt - Marginal)**")
                else:
                    st.error(f"**SIX SIGMA: {sigma:.2f} (Cần cải tiến - Poor)**")

                # Hàng 2: Các chỉ số chi tiết
                col_a, col_b, col_c = st.columns(3)
                col_a.metric("CV (%)", f"{cv:.2f}%")
                col_b.metric("Bias (%)", f"{bias_pct:.2f}%")
                col_c.metric("TEa (%)", f"{tea}%")

                st.markdown("---")
                
                # Hàng 3: Phân tích QGI (Chỉ hiển thị khi Sigma < 6)
                if sigma < 6:
                    st.write("**Phân tích nguyên nhân (QGI):**")
                    qgi_val = sigma_results[lvl]['qgi']
                    
                    # Tạo màu sắc cho thanh tiến trình QGI
                    if qgi_val < 0.8:
                        st.error(f"QGI = {qgi_val:.2f} → {qgi_reason}")
                    elif 0.8 <= qgi_val <= 1.2:
                        st.warning(f"QGI = {qgi_val:.2f} → {qgi_reason}")
                    else:
                        st.error(f"QGI = {qgi_val:.2f} → {qgi_reason}")
                else:
                    st.write("✅ **Hiệu năng hoàn hảo, không cần phân tích QGI.**")
# 5. BẢNG TỔNG HỢP CÓ MÀU SẮC (Dưới biểu đồ)
    st.subheader("📋 Bảng tổng hợp hiệu năng")
    
    summary_data = []
    for l, res in sigma_results.items():
        summary_data.append({
            "Mức độ": f"Level {l}",
            "N": res['n'],
            "CV%": res['cv'],
            "Bias%": res['bias'],
            "Sigma": res['sigma'],
            "QGI": res['qgi'],
            "Đánh giá": "Đạt" if res['sigma'] >= 3 else "Không đạt"
        })
    
    df_summary = pd.DataFrame(summary_data)

    # Hàm tô màu cho cột Sigma
    def color_sigma(val):
        if val >= 6: color = '#b3e6ff' # Xanh dương nhạt
        elif val >= 4: color = '#c6efce' # Xanh lá nhạt
        elif val >= 3: color = '#ffeb9c' # Vàng nhạt
        else: color = '#ffc7ce' # Đỏ nhạt
        return f'background-color: {color}'

    # Hiển thị bảng đã được format
    st.dataframe(
        df_summary.style.applymap(color_sigma, subset=['Sigma'])
        .format({'CV%': "{:.2f}", 'Bias%': "{:.2f}", 'Sigma': "{:.2f}", 'QGI': "{:.2f}"}),
        use_container_width=True
    )                    

# --- 3. BIỂU ĐỒ SIX SIGMA METHOD DECISION CHART ---
    st.markdown("---")
    st.subheader("📈 Biểu đồ Method Decision Chart")
    
    # Vẽ biểu đồ và lưu vào biến fig_sigma
    fig_sigma = plot_sigma_chart(sigma_plot_data, tea)
    st.pyplot(fig_sigma)
    
 
# --- 4. XUẤT BÁO CÁO (Đã tích hợp MAU Biological Variation) ---
    st.markdown("---")
    if st.button("📥 Tải Báo Cáo Tổng Hợp (Excel)"):
        with st.spinner("Đang khởi tạo báo cáo..."):
            # 1. LJ Chart: Dùng df_filtered để vẽ đúng giai đoạn báo cáo
            img_lj_buffer = None
            if not df_filtered.empty:
                fig_lj = plot_levey_jennings(df_filtered, f"LJ Chart: {current_test['name']}", show_legend=False)
                if fig_lj:
                    img_lj_buffer = io.BytesIO()
                    fig_lj.savefig(img_lj_buffer, format='png', bbox_inches='tight')
                    img_lj_buffer.seek(0) 

            # 2. Sigma Chart
            img_sigma_buffer = None
            if fig_sigma:
                img_sigma_buffer = io.BytesIO()
                fig_sigma.savefig(img_sigma_buffer, format='png', bbox_inches='tight')
                img_sigma_buffer.seek(0)

            # 3. Tính toán MU và Biological Variation Limits
            # Lấy cvi, cvg từ current_test đã load từ Database
            cvi_val = float(current_test.get('cvi', 0.0))
            cvg_val = float(current_test.get('cvg', 0.0))
            
            # Tính 3 mức MAU cho báo cáo
            m_min = 0.75 * cvi_val + 1.65 * (0.375 * np.sqrt(cvi_val**2 + cvg_val**2))
            m_des = 0.5 * cvi_val + 1.65 * (0.25 * np.sqrt(cvi_val**2 + cvg_val**2))
            m_opt = 0.25 * cvi_val + 1.65 * (0.125 * np.sqrt(cvi_val**2 + cvg_val**2))
            mau_limits_input = [m_min, m_des, m_opt]

            mu_res = {}
            for lvl in [1, 2]:
                d = df_filtered[df_filtered['level'] == lvl] if not df_filtered.empty else pd.DataFrame()
                if len(d) >= 2:
                    mean_val = d['value'].mean()
                    sd_val = d['value'].std()
                    u_prec = sd_val
                    u_bias = (bias_pct / 100) * mean_val if mean_val else 0
                    uc = np.sqrt(u_prec**2 + u_bias**2)
                    
                    # MAU cũ theo TEa cài đặt (giữ để tham khảo nếu cần)
                    mau_tea = (current_test.get('tea', 10.0) / 100) * mean_val if mean_val else 0
                    
                    mu_res[lvl] = {
                        'u_prec': round(u_prec, 4), 
                        'u_bias': round(u_bias, 4),
                        'uc': round(uc, 4), 
                        'ue': round(uc * 2, 4), 
                        'mau': round(mau_tea, 4)
                    }
                else:
                    mu_res[lvl] = {}

            # 4. GỌI HÀM EXCEL: TRUYỀN THÊM mau_limits_input
            try:
                excel_data = generate_excel_report_comprehensive(
                    current_test, 
                    df_full_history,   # Dữ liệu gốc để tính Westgard
                    df_eqa, 
                    mu_res, 
                    sigma_results,
                    img_lj_buffer,   
                    img_sigma_buffer, 
                    (start_d, end_d),  # Khoảng thời gian báo cáo
                    mau_limits_input   # <--- THAM SỐ MỚI ĐÃ ĐƯỢC THÊM VÀO
                )
                
                st.download_button(
                    label="📂 Nhấn vào đây để tải file .xlsx",
                    data=excel_data,
                    file_name=f"Bao_cao_QLCL_{current_test['name']}_{start_d}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Lỗi khi tạo file Excel: {e}")
                # In chi tiết lỗi ra console để debug nếu cần
                print(f"Excel Error: {e}")

                
# === TAB 6: QUẢN TRỊ (ADMIN) ===

# Lấy mật khẩu quản trị hiện tại từ DB (Mặc định là 'admin123' nếu chưa thiết lập)
ADMIN_PASSWORD_KEY = "admin_password"
current_admin_pwd = db.get_setting(ADMIN_PASSWORD_KEY, "admin123")


# === TAB 6: QUẢN TRỊ (ADMIN) ===
with tabs[5]:
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
        
# TAB 7: IMPORT DỮ LIỆU
with tabs[6]:
    st.header("📥 Import kết quả IQC từ file Excel")
    
    # Chia sub-tabs để giao diện gọn gàng
    sub1, sub2 = st.tabs(["🚀 Import Dữ liệu", "🔗 Cấu hình Mapping"])
    
    with sub2:
        # Gọi hàm mapping đã định nghĩa
        manage_test_mapping() 
        
    with sub1:
        uploaded_file = st.file_uploader("Chọn file Excel kết quả", type=["xlsx", "xls"])
        
        if uploaded_file:
            try:
                df_raw = pd.read_excel(uploaded_file)
                # Danh sách cột bắt buộc phải có trong file Excel
                required_cols = ['Thời gian chạy', 'Máy xét nghiệm', 'Tên xét nghiệm', 'Kết quả', 'Lô', 'Mức QC']
                
                if all(c in df_raw.columns for c in required_cols):
                    df_import = df_raw[required_cols].copy()
                    df_import['Thời gian chạy'] = pd.to_datetime(df_import['Thời gian chạy'])
                    
                    st.write("### Xem trước dữ liệu:")
                    st.dataframe(df_import.head(5))
                    
                    if st.button("🚀 Xác nhận Import"):
                        with st.spinner("Đang xử lý..."):
                            # Gọi qua đối tượng db đã khởi tạo ở đầu file main.py
                            count, errors = db.import_iqc_from_dataframe(df_import)
                            
                            if count > 0:
                                st.success(f"Đã Import thành công {count} kết quả!")
                            if errors:
                                with st.expander("Chi tiết dòng lỗi/Chưa mapping"):
                                    for err in errors: st.warning(err)
                        st.rerun()
                else:
                    st.error(f"File Excel thiếu cột. Cần: {', '.join(required_cols)}")
            except Exception as e:
                st.error(f"Lỗi khi đọc file: {e}")
