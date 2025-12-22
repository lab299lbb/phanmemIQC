# File: db_module.py
import sqlite3
import pandas as pd
from datetime import datetime, date, timedelta

class DBManager:

    def __init__(self, db_path):
        self.db_path = db_path
        self.conn = sqlite3.connect(db_path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        # THÊM DÒNG DƯỚI ĐÂY:
        self.cursor = self.conn.cursor() 
        self.create_tables()

    def get_setting(self, key, default=None):
        sql = "SELECT value FROM settings WHERE key = ?"
        # Bây giờ self.cursor đã tồn tại và có thể sử dụng
        result = self.cursor.execute(sql, (key,)).fetchone()
        return result[0] if result else default

    def create_tables(self):
        c = self.conn.cursor()
        # 1. Bảng TESTS
        c.execute("""CREATE TABLE IF NOT EXISTS tests (
                id INTEGER PRIMARY KEY, name TEXT UNIQUE NOT NULL, unit TEXT,
                tea REAL, device TEXT, cvi REAL DEFAULT 0, cvg REAL DEFAULT 0)""")
        
        # 2. Bảng LOTS: Lưu thông tin chi tiết từng Lô, TÁCH BIỆT THEO LEVEL
        c.execute("""CREATE TABLE IF NOT EXISTS lots (
                id INTEGER PRIMARY KEY, test_id INTEGER, lot_number TEXT NOT NULL,
                level INTEGER NOT NULL, method TEXT, expiry_date TEXT,
                mean REAL, sd REAL, FOREIGN KEY (test_id) REFERENCES tests (id))""")
        
        # 3. Bảng IQC_DATA
        c.execute("""CREATE TABLE IF NOT EXISTS iqc_data (
                id INTEGER PRIMARY KEY, lot_id INTEGER, date TEXT, level INTEGER,
                value REAL, note TEXT, FOREIGN KEY (lot_id) REFERENCES lots (id))""")
        
        # 4. Bảng EQA_DATA
        c.execute("""CREATE TABLE IF NOT EXISTS eqa_data (
                id INTEGER PRIMARY KEY, test_id INTEGER, date TEXT, lab_value REAL,
                ref_value REAL, sd_group REAL, sample_id TEXT,
                FOREIGN KEY (test_id) REFERENCES tests (id))""")

	# 5. Bảng SETTINGS: Lưu các cài đặt chung (ví dụ: mật khẩu quản trị)
        c.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)
	# --- BỔ SUNG CHỈ MỤC (INDEX) ĐỂ TĂNG TỐC ĐỘ TRUY VẤN (QUAN TRỌNG VỚI DỮ LIỆU LỚN) ---
        
        # 1. Index cho bảng LOTS (tăng tốc độ tìm kiếm Lots theo Test)
        c.execute("CREATE INDEX IF NOT EXISTS idx_lots_test ON lots (test_id);")
        
        # 2. Index cho bảng IQC_DATA (tăng tốc độ JOIN theo Lot và sắp xếp/lọc theo Ngày)
        c.execute("CREATE INDEX IF NOT EXISTS idx_iqc_lot_date ON iqc_data (lot_id, date);")
        
        # 3. Index cho bảng EQA_DATA (tăng tốc độ tìm kiếm EQA theo Test)
        c.execute("CREATE INDEX IF NOT EXISTS idx_eqa_test ON eqa_data (test_id);")
        
        self.conn.commit()	
    def update_test(self, test_id, name, unit, device, tea, cvi, cvg):
        """Cập nhật thông tin xét nghiệm bao gồm cả TEa, CVi, CVg"""
        sql = """
            UPDATE tests 
            SET name = ?, 
                unit = ?, 
                device = ?, 
                tea = ?, 
                CVi = ?, 
                CVg = ? 
            WHERE id = ?
        """
        try:
            cursor = self.conn.cursor()
            cursor.execute(sql, (name, unit, device, tea, cvi, cvg, test_id))
            self.conn.commit()
            return True
        except Exception as e:
            print(f"Lỗi cập nhật Database: {e}")
            return False

    def delete_test(self, test_id):
        """Xóa Test VÀ TẤT CẢ dữ liệu (IQC, EQA, Lots) liên quan."""
        try:
            # Xóa các dữ liệu liên quan (theo thứ tự: iqc -> lots -> eqa -> tests)
            self.cursor.execute("DELETE FROM iqc_data WHERE lot_id IN (SELECT id FROM lots WHERE test_id = ?)", (test_id,))
            self.cursor.execute("DELETE FROM lots WHERE test_id = ?", (test_id,))
            self.cursor.execute("DELETE FROM eqa_data WHERE test_id = ?", (test_id,))
            self.cursor.execute("DELETE FROM tests WHERE id = ?", (test_id,))
            self.conn.commit()
            return True
        except Exception as e:
            print(f"LỖI DB: Không thể xóa Test ID {test_id}: {e}")
            return False
            
    # --- QUẢN LÝ THIẾT BỊ & TESTS ---
    def get_all_devices(self):
        try:
            c = self.conn.cursor()
            c.execute("SELECT DISTINCT device FROM tests WHERE device IS NOT NULL AND device != '' ORDER BY device")
            return [row[0] for row in c.fetchall()]
        except: return []

    def add_test(self, name, unit, tea, device, cvi=0, cvg=0):
        try:
            self.cursor.execute("INSERT INTO tests (name, unit, tea, device, cvi, cvg) VALUES (?, ?, ?, ?, ?, ?)", 
                      (name, unit, tea, device, cvi, cvg))
            self.conn.commit(); return True
        except: return False

    def get_all_tests(self):
        return pd.read_sql_query("SELECT * FROM tests", self.conn)

    def update_test_info(self, test_id, name, unit, tea, device, cvi, cvg):
        try:
            self.cursor.execute("UPDATE tests SET name=?, unit=?, tea=?, device=?, cvi=?, cvg=? WHERE id=?", 
                      (name, unit, tea, device, cvi, cvg, test_id))
            self.conn.commit(); return True
        except: return False

    def delete_test(self, test_id):
        try:
            self.cursor.execute("DELETE FROM eqa_data WHERE test_id=?", (test_id,))
            self.cursor.execute("SELECT id FROM lots WHERE test_id=?", (test_id,))
            lots = self.cursor.fetchall()
            for lot in lots: self.cursor.execute("DELETE FROM iqc_data WHERE lot_id=?", (lot[0],))
            self.cursor.execute("DELETE FROM lots WHERE test_id=?", (test_id,))
            self.cursor.execute("DELETE FROM tests WHERE id=?", (test_id,))
            self.conn.commit(); return True
        except: return False

    # --- QUẢN LÝ LOTS ---
    def add_lot(self, test_id, lot_number, level, method, expiry_date, mean, sd):
        try:
            exp_str = expiry_date.strftime('%Y-%m-%d') if isinstance(expiry_date, (datetime, pd.Timestamp, date)) else str(expiry_date)
            self.cursor.execute("INSERT INTO lots (test_id, lot_number, level, method, expiry_date, mean, sd) VALUES (?, ?, ?, ?, ?, ?, ?)",
                             (test_id, lot_number, level, method, exp_str, mean, sd))
            self.conn.commit(); return True
        except: return False

    def get_lots_for_test(self, test_id):
        return pd.read_sql_query(f"SELECT * FROM lots WHERE test_id={test_id} ORDER BY id DESC", self.conn)

    def update_lot_params(self, lot_id, lot_number, method, expiry_date, mean, sd):
        try:
            exp_str = expiry_date.strftime('%Y-%m-%d') if isinstance(expiry_date, (datetime, pd.Timestamp, date)) else str(expiry_date)
            self.cursor.execute("UPDATE lots SET lot_number=?, method=?, expiry_date=?, mean=?, sd=? WHERE id=?", 
                             (lot_number, method, exp_str, mean, sd, lot_id))
            self.conn.commit(); return True
        except: return False
            
    def delete_lot(self, lot_id):
        try:
            self.cursor.execute("DELETE FROM iqc_data WHERE lot_id=?", (lot_id,))
            self.cursor.execute("DELETE FROM lots WHERE id=?", (lot_id,))
            self.conn.commit(); return True
        except: return False
            
    # --- QUẢN LÝ IQC DATA ---
    def add_iqc(self, lot_id, dt, level, value, note):
        try:
            d_str = dt.strftime('%Y-%m-%d') if isinstance(dt, datetime) else str(dt)
            self.cursor.execute("INSERT INTO iqc_data (lot_id, date, level, value, note) VALUES (?, ?, ?, ?, ?)",
                      (lot_id, d_str, level, value, note))
            self.conn.commit(); return True
        except: return False
    def update_iqc_action(self, result_id, action_text):
        conn = self.create_connection()
        cursor = conn.cursor()
        cursor.execute("UPDATE iqc_results SET action = ? WHERE id = ?", (action_text, result_id))
        conn.commit()
        conn.close()
    def get_iqc_data_continuous(self, test_id, max_months=None):
        """
        Lấy IQC nối các Lot, kèm theo Mean/SD của Lot đó.
        Có thể giới hạn dữ liệu trong max_months gần nhất để tối ưu hiệu suất.
        """
        
        # Thêm điều kiện giới hạn ngày nếu có tham số
        date_filter = ""
        if max_months is not None and max_months > 0:
            # Tính toán ngày bắt đầu: max_months trước ngày hiện tại
            # Sử dụng 30.5 ngày/tháng để ước tính chính xác hơn 
            start_date = datetime.now() - timedelta(days=30.5 * max_months) 
            # Lọc theo cột date trong bảng iqc_data
            date_filter = f" AND i.date >= '{start_date.strftime('%Y-%m-%d')}'"

        query = f"""
            SELECT i.id, i.date, i.level, i.value, i.note, i.lot_id,
                l.lot_number, l.mean AS target_mean, l.sd AS target_sd, l.expiry_date
            FROM iqc_data i
            JOIN lots l ON i.lot_id = l.id
            WHERE l.test_id = {test_id}
            {date_filter}
            ORDER BY i.date ASC, i.id ASC
        """
        
        df = pd.read_sql_query(query, self.conn)
        
        if not df.empty:
            df['date'] = pd.to_datetime(df['date'])
            # Bổ sung cột Z-score, cần thiết cho biểu đồ và Westgard
            df['z_score'] = (df['value'] - df['target_mean']) / df['target_sd']
            # Xử lý trường hợp sd = 0
            df.loc[df['target_sd'] == 0, 'z_score'] = 0 
            
        return df

    def get_iqc_data_by_lot(self, lot_id):
        df = pd.read_sql_query(f"SELECT * FROM iqc_data WHERE lot_id={lot_id} ORDER BY date DESC", self.conn)
        if not df.empty: df['date'] = pd.to_datetime(df['date'])
        return df
        
    def update_iqc_data(self, iqc_id, dt, level, value, note):
        try:
            d_str = dt.strftime('%Y-%m-%d') if isinstance(dt, datetime) else str(dt)
            self.cursor.execute("UPDATE iqc_data SET date=?, level=?, value=?, note=? WHERE id=?", (d_str, level, value, note, iqc_id))
            self.conn.commit(); return True
        except: return False

    def delete_iqc_data(self, iqc_id):
        try:
            self.cursor.execute("DELETE FROM iqc_data WHERE id=?", (iqc_id,))
            self.conn.commit(); return True
        except: return False
    # Chạy đoạn này trong hàm khởi tạo hoặc terminal để nâng cấp bảng
    def upgrade_db():
        conn = sqlite3.connect("lab_database.db")
        cursor = conn.cursor()
        try:
            cursor.execute("ALTER TABLE iqc_results ADD COLUMN action TEXT DEFAULT ''")
            conn.commit()
        except:
            pass # Cột đã tồn tại
        finally:
            conn.close()
    def import_iqc_from_dataframe(self, df):
        """
        df: DataFrame đã được lọc đúng các cột cần thiết
        """
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        success_count = 0
        error_logs = []

        for _, row in df.iterrows():
            try:
                # 1. Tìm lot_id dựa trên Số Lô và Tên xét nghiệm
                # Lưu ý: Cần join với bảng tests để đảm bảo đúng xét nghiệm
                query_lot = """
                    SELECT l.id FROM iqc_lots l
                    JOIN tests t ON l.test_id = t.id
                    WHERE l.lot_number = ? AND t.name = ? AND l.level = ?
                """
                cursor.execute(query_lot, (str(row['Lô']), row['Tên xét nghiệm'], int(row['Mức QC'])))
                lot_res = cursor.fetchone()

                if lot_res:
                    lot_id = lot_res[0]
                    # 2. Chèn vào bảng iqc_results
                    # Kiểm tra tránh trùng lặp (nếu cùng ngày, cùng lot, cùng giá trị thì bỏ qua)
                    check_query = "SELECT id FROM iqc_results WHERE lot_id = ? AND date = ? AND value = ?"
                    cursor.execute(check_query, (lot_id, row['Thời gian chạy'], row['Kết quả']))
                    
                    if not cursor.fetchone():
                        cursor.execute("""
                            INSERT INTO iqc_results (lot_id, date, value, level, note)
                            VALUES (?, ?, ?, ?, ?)
                        """, (lot_id, row['Thời gian chạy'], row['Kết quả'], int(row['Mức QC']), f"Import từ máy {row['Máy xét nghiệm']}"))
                        success_count += 1
                else:
                    error_logs.append(f"Không tìm thấy Lot {row['Lô']} cho XN {row['Tên xét nghiệm']} Level {row['Mức QC']}")
            
            except Exception as e:
                error_logs.append(f"Lỗi tại dòng {row['Tên xét nghiệm']}: {str(e)}")

        conn.commit()
        conn.close()
        return success_count, error_logs
    def get_iqc_results_with_mapping(self, test_id):
        """
        Lấy tất cả kết quả IQC. 
        Dùng LEFT JOIN để không bị mất dữ liệu khi chưa có mapping.
        """
        query = """
            SELECT 
                r.*, 
                m.external_name 
            FROM iqc_results r
            LEFT JOIN iqc_lots l ON r.lot_id = l.id
            LEFT JOIN test_mapping m ON l.test_id = m.test_id
            WHERE l.test_id = ? OR r.lot_id IN (SELECT id FROM iqc_lots WHERE test_id = ?)
            ORDER BY r.date DESC
        """
        # Lưu ý: Thay 'iqc_lots' thành 'lots' nếu database của bạn đặt tên là 'lots'
        try:
            return pd.read_sql(query, self.conn, params=(int(test_id), int(test_id)))
        except:
            # Phương án dự phòng nếu tên bảng iqc_lots sai
            query_alt = query.replace("iqc_lots", "lots")
            return pd.read_sql(query_alt, self.conn, params=(int(test_id), int(test_id)))

    # --- QUẢN LÝ EQA DATA ---
    def add_eqa(self, test_id, dt, lab, ref, sd, sample_id):
        try:
            d_str = dt.strftime('%Y-%m-%d') if isinstance(dt, datetime) else str(dt)
            self.cursor.execute("INSERT INTO eqa_data (test_id, date, lab_value, ref_value, sd_group, sample_id) VALUES (?, ?, ?, ?, ?, ?)",
                      (test_id, d_str, lab, ref, sd, sample_id))
            self.conn.commit(); return True
        except: return False

    def get_eqa_data(self, test_id):
        df = pd.read_sql_query(f"SELECT * FROM eqa_data WHERE test_id={test_id} ORDER BY date ASC", self.conn)
        if not df.empty: df['date'] = pd.to_datetime(df['date'])
        return df
        
    def delete_eqa(self, eqa_id):
        try:
            self.cursor.execute("DELETE FROM eqa_data WHERE id=?", (eqa_id,))
            self.conn.commit(); return True
        except: return False

    def update_eqa(self, eqa_id, data):
        if not data: return False
        try:
            set_parts = [f"{key} = ?" for key in data.keys()]
            values = list(data.values()) + [eqa_id]
            self.cursor.execute(f"UPDATE eqa_data SET {', '.join(set_parts)} WHERE id = ?", values)
            self.conn.commit(); return True
        except: return False

    # --- HÀM QUẢN LÝ CÀI ĐẶT CHUNG (Settings) ---
    def get_setting(self, key, default=None):
        """Lấy giá trị của một cài đặt."""
        sql = "SELECT value FROM settings WHERE key = ?"
        result = self.cursor.execute(sql, (key,)).fetchone()
        return result[0] if result else default

    def set_setting(self, key, value):
        """Cài đặt hoặc cập nhật một giá trị cài đặt."""
        try:
            sql = "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)"
            self.cursor.execute(sql, (key, value))
            self.conn.commit()
            return True
        except Exception as e:
            print(f"LỖI DB: Không thể lưu cài đặt {key}: {e}")
            return False    
    # --- ADMIN / UTILS ---
    def execute_raw(self, sql):
        """Dùng cho tính năng Reset DB"""
        try:
            self.cursor.execute(sql)
            self.conn.commit(); return True
        except: return False

    def upgrade_database_for_pro_features():
        conn = sqlite3.connect("lab_data.db")
        cursor = conn.cursor()
        
        # 1. Bảng Mapping: Liên kết tên trên máy (AU640) với tên trong phần mềm
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS test_mapping (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                test_id INTEGER,
                external_name TEXT UNIQUE,
                FOREIGN KEY (test_id) REFERENCES tests (id)
            )
        """)
        
        # 2. Thêm cột Phê duyệt vào bảng iqc_results (nếu chưa có)
        try:
            cursor.execute("ALTER TABLE iqc_results ADD COLUMN is_approved INTEGER DEFAULT 0")
            cursor.execute("ALTER TABLE iqc_results ADD COLUMN approved_by TEXT")
            cursor.execute("ALTER TABLE iqc_results ADD COLUMN approved_at TEXT")
        except:
            pass # Cột đã tồn tại
            
        conn.commit()
        conn.close()

    upgrade_database_for_pro_features()