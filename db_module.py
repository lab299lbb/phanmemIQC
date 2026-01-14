# File: db_module.py
import sqlite3
import pandas as pd
from datetime import datetime, date, timedelta

import sqlite3

class DBManager:
    def __init__(self, db_path):
        self.db_path = db_path
        self.conn = sqlite3.connect(db_path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.cursor = self.conn.cursor() 
        self.create_tables()
        self.upgrade_tables() # Tự động vá lỗi thiếu cột khi chạy app

    def get_setting(self, key, default=None):
        sql = "SELECT value FROM settings WHERE key = ?"
        result = self.cursor.execute(sql, (key,)).fetchone()
        return result[0] if result else default

    def create_tables(self):
        c = self.conn.cursor()
        
        # 1. Bảng TESTS
        c.execute("""CREATE TABLE IF NOT EXISTS tests (
                id INTEGER PRIMARY KEY AUTOINCREMENT, 
                name TEXT UNIQUE NOT NULL, 
                unit TEXT,
                tea REAL, 
                device TEXT, 
                cvi REAL DEFAULT 0, 
                cvg REAL DEFAULT 0,
                last_mu_review TEXT)""")
        
        # 2. Bảng LOTS
        c.execute("""CREATE TABLE IF NOT EXISTS lots (
                id INTEGER PRIMARY KEY AUTOINCREMENT, 
                test_id INTEGER, 
                lot_number TEXT NOT NULL,
                level INTEGER NOT NULL, 
                method TEXT, 
                expiry_date TEXT,
                mean REAL, 
                sd REAL, 
                FOREIGN KEY (test_id) REFERENCES tests (id))""")
        
        # 3. Bảng IQC_RESULTS (Thống nhất tên iqc_results)
        c.execute("""CREATE TABLE IF NOT EXISTS iqc_results (
                id INTEGER PRIMARY KEY AUTOINCREMENT, 
                lot_id INTEGER, 
                date TEXT, 
                level INTEGER,
                value REAL, 
                note TEXT, 
                FOREIGN KEY (lot_id) REFERENCES lots (id))""")
        
        # 4. Bảng EQA_RESULTS
        c.execute("""CREATE TABLE IF NOT EXISTS eqa_results (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                test_id INTEGER,
                date TEXT,
                lab_value REAL,
                ref_value REAL,
                sd_group REAL,
                sdi REAL,
                program_name TEXT,
                FOREIGN KEY (test_id) REFERENCES tests (id))""")
        
        # 5. Bảng SETTINGS
        c.execute("""CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT)""")
        
        # 6. Bảng MAPPING
        c.execute("""CREATE TABLE IF NOT EXISTS test_mapping (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                test_id INTEGER,
                external_name TEXT,
                FOREIGN KEY (test_id) REFERENCES tests (id))""")

        # --- CHỈ MỤC (INDEX) - Sửa lỗi iqc_data thành iqc_results ở đây ---
        c.execute("CREATE INDEX IF NOT EXISTS idx_lots_test ON lots (test_id);")
        c.execute("CREATE INDEX IF NOT EXISTS idx_iqc_lot_date ON iqc_results (lot_id, date);")
        c.execute("CREATE INDEX IF NOT EXISTS idx_eqa_res_test ON eqa_results (test_id);")
        
        self.conn.commit()

    def upgrade_tables(self):
        """Tự động thêm các cột còn thiếu vào DB cũ nếu cần"""
        c = self.conn.cursor()
        try:
            # Thêm cột cho eqa_results nếu chưa có
            columns = [info[1] for info in c.execute("PRAGMA table_info(eqa_results)").fetchall()]
            if 'sdi' not in columns:
                c.execute("ALTER TABLE eqa_results ADD COLUMN sdi REAL")
            if 'sd_group' not in columns:
                c.execute("ALTER TABLE eqa_results ADD COLUMN sd_group REAL")
            
            # Thêm cột cho tests nếu chưa có
            test_columns = [info[1] for info in c.execute("PRAGMA table_info(tests)").fetchall()]
            if 'last_mu_review' not in test_columns:
                c.execute("ALTER TABLE tests ADD COLUMN last_mu_review TEXT")
                
            self.conn.commit()
        except Exception as e:
            print(f"Lỗi khi nâng cấp DB: {e}")

    def upgrade_tables(self):
        """Hàm tự động vá database cũ: Thêm các cột thiếu mà không cần xóa file DB"""
        c = self.conn.cursor()
        
        # Danh sách các cột cần kiểm tra bổ sung
        upgrades = [
            ("eqa_results", "sd_group", "REAL"),
            ("eqa_results", "sdi", "REAL"),
            ("tests", "last_mu_review", "TEXT")
        ]
        
        for table, column, col_type in upgrades:
            try:
                c.execute(f"ALTER TABLE {table} ADD COLUMN {column} {col_type}")
            except sqlite3.OperationalError:
                pass # Cột đã tồn tại, không làm gì cả
        
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
        """Xóa một Lot khỏi hệ thống"""
        try:
            cursor = self.conn.cursor()
            # Lưu ý: Cân nhắc việc xóa Lot có thể ảnh hưởng đến dữ liệu IQC đã nhập
            sql = "DELETE FROM lots WHERE id = ?"
            cursor.execute(sql, (lot_id,))
            self.conn.commit()
            return True
        except Exception as e:
            print(f"Lỗi khi xóa Lot: {e}")
            return False
    def update_lot(self, lot_id, lot_number, mean, sd, expiration_date):
            """Cập nhật thông tin một Lot đã có"""
            try:
                cursor = self.conn.cursor()
                sql = """UPDATE lots 
                         SET lot_number = ?, mean = ?, sd = ?, expiration_date = ? 
                         WHERE id = ?"""
                cursor.execute(sql, (lot_number, mean, sd, expiration_date, lot_id))
                self.conn.commit()
                return True
            except Exception as e:
                print(f"Lỗi khi cập nhật Lot: {e}")
                return False

    def get_test_by_name(self, name):
        # Giả sử bạn dùng SQLite
        query = "SELECT id FROM tests WHERE name = ?"
        result = self.conn.execute(query, (name,)).fetchone()
        return {'id': result[0]} if result else None              
    # --- QUẢN LÝ IQC DATA ---
    def add_iqc_data(self, lot_id, dt, level, value, note):
        try:
            # Ép mọi thứ về định dạng chuẩn ISO: YYYY-MM-DD HH:mm:ss
            if isinstance(dt, str):
                # Nếu là chuỗi, thử chuyển sang datetime trước để chuẩn hóa
                dt_obj = pd.to_datetime(dt, dayfirst=True)
                d_str = dt_obj.strftime('%Y-%m-%d %H:%M:%S')
            else:
                d_str = dt.strftime('%Y-%m-%d %H:%M:%S')

            query = "INSERT INTO iqc_results (lot_id, date, value, level, note) VALUES (?, ?, ?, ?, ?)"
            self.cursor.execute(query, (lot_id, d_str, value, level, note))
            self.conn.commit()
            return True
        except Exception as e:
            print(f"Lỗi chuẩn hóa ngày: {e}")
            return False
                   
    def update_iqc_action(self, result_id, action_text):
        conn = self.create_connection()
        cursor = conn.cursor()
        cursor.execute("UPDATE iqc_results SET action = ? WHERE id = ?", (action_text, result_id))
        conn.commit()
        conn.close()
    def get_iqc_data_continuous(self, test_id, max_months=None):
        """
        Lấy dữ liệu IQC. Nếu có max_months, chỉ lấy dữ liệu trong khoảng tháng đó.
        """
        if max_months:
            query = f"""
                SELECT r.*, l.lot_number 
                FROM iqc_results r
                JOIN lots l ON r.lot_id = l.id
                WHERE l.test_id = ? 
                AND r.date >= date('now', '-{max_months} months')
                ORDER BY r.date ASC
            """
        else:
            query = """
                SELECT r.*, l.lot_number 
                FROM iqc_results r
                JOIN lots l ON r.lot_id = l.id
                WHERE l.test_id = ?
                ORDER BY r.date ASC
            """
        return pd.read_sql(query, self.conn, params=(test_id,))

    def get_iqc_data_filtered(test_id, d_start, d_end):
        """
        Truy vấn dữ liệu IQC trong khoảng thời gian xác định và sắp xếp theo ngày.
        """
        # Chuyển đổi date object sang string ISO để truy vấn SQL
        s_date = d_start.strftime('%Y-%m-%d')
        e_date = d_end.strftime('%Y-%m-%d')
        
        # Giả sử bạn dùng SQL, hãy điều chỉnh theo thư viện DB của bạn
        query = f"""
            SELECT date, value, level 
            FROM iqc_results 
            WHERE test_id = '{test_id}' 
            AND date BETWEEN '{s_date}' AND '{e_date}'
            ORDER BY date ASC
        """
        # Trả về DataFrame từ database
        return db.query_to_dataframe(query)
    def get_iqc_data_by_lot(self, lot_id):
        """
        Truy vấn toàn bộ kết quả IQC dựa trên lot_id.
        Bổ sung cột level để phục vụ phân tích Westgard.
        """
        query = """
            SELECT 
                id, 
                date, 
                value, 
                level,  -- Bắt buộc phải có cột này
                note 
            FROM iqc_results 
            WHERE lot_id = ? 
            ORDER BY date DESC
        """
        try:
            # Sử dụng self.conn đã được thiết lập trong class DBManager
            return pd.read_sql(query, self.conn, params=(lot_id,))
        except Exception as e:
            print(f"Lỗi truy vấn: {e}")
            return pd.DataFrame()

    def get_iqc_data_by_lot_full(self, lot_id):
        """Lấy toàn bộ dữ liệu của một Lot, không phân biệt nguồn gốc"""
        query = """
            SELECT 
                id, 
                date, 
                value, 
                note 
            FROM iqc_results 
            WHERE lot_id = ? 
            ORDER BY date DESC
        """
        return pd.read_sql(query, self.conn, params=(lot_id,))
        
    def update_iqc_data(self, iqc_id, note, dt, level, value):
            """Cập nhật kết quả IQC vào bảng iqc_results"""
            try:
                # Chuẩn hóa ngày tháng về chuỗi ISO để lưu vào SQLite
                if isinstance(dt, (pd.Timestamp, datetime)):
                    d_str = dt.strftime('%Y-%m-%d %H:%M:%S')
                else:
                    d_str = str(dt)

                # SỬA TÊN BẢNG THÀNH iqc_results ĐỂ ĐỒNG BỘ VỚI HÀM IMPORT
                sql = "UPDATE iqc_results SET date=?, level=?, value=?, note=? WHERE id=?"
                self.cursor.execute(sql, (d_str, level, value, note, iqc_id))
                
                self.conn.commit()
                return True
            except Exception as e:
                print(f"Lỗi SQL: {e}") # In lỗi ra console để debug
                return False
  
    def delete_iqc_result(self, row_id):
        try:
            # Đảm bảo row_id là số nguyên
            clean_id = int(row_id)
            self.cursor.execute("DELETE FROM iqc_results WHERE id = ?", (clean_id,))
            self.conn.commit()
            
            if self.cursor.rowcount > 0:
                return True
            else:
                print(f"Cảnh báo: Không tìm thấy dòng có ID {clean_id} để xóa.")
                return False
        except Exception as e:
            print(f"Lỗi SQL xóa IQC: {e}")
            return False

   # def delete_iqc_data(self, iqc_id):
    #    try:
     #       self.cursor.execute("DELETE FROM iqc_data WHERE id=?", (iqc_id,))
      #      self.conn.commit(); return True
       # except: return False

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
        success_count = 0
        errors = []

        for _, row in df.iterrows():
            try:
                # 1. Chuẩn hóa dữ liệu đầu vào
                ext_name = str(row['Tên xét nghiệm']).strip()
                lot_num = str(row['Lô']).strip()
                level = int(row['Mức QC'])
                value = float(row['Kết quả'])
                
                # CHUẨN HÓA NGÀY THÁNG: Chuyển về dạng YYYY-MM-DD HH:MM:SS
                run_date_raw = pd.to_datetime(row['Thời gian chạy'])
                run_date_iso = run_date_raw.strftime('%Y-%m-%d %H:%M:%S')

                # 2. Tra cứu test_id từ Mapping
                self.cursor.execute("SELECT test_id FROM test_mapping WHERE external_name = ?", (ext_name,))
                mapping_res = self.cursor.fetchone()
                
                if not mapping_res:
                    errors.append(f"Chưa mapping tên: {ext_name}")
                    continue
                
                test_id = mapping_res[0]

                # 3. Tra cứu lot_id (Sử dụng LIKE để linh hoạt hơn với số Lô)
                self.cursor.execute("""
                    SELECT id FROM lots 
                    WHERE test_id = ? AND lot_number LIKE ? AND level = ?
                """, (test_id, f"%{lot_num}%", level))
                lot_res = self.cursor.fetchone()

                if not lot_res:
                    errors.append(f"Không tìm thấy Lô {lot_num} (Mức {level}) cho XN này.")
                    continue
                
                lot_id = lot_res[0]

                # 4. KIỂM TRA TRÙNG LẶP (Dựa trên ngày chuẩn ISO)
                self.cursor.execute("""
                    SELECT id FROM iqc_results 
                    WHERE lot_id = ? AND date = ? AND value = ?
                """, (lot_id, run_date_iso, value))
                
                if self.cursor.fetchone():
                    continue 

                # 5. Chèn vào bảng kết quả với ngày chuẩn ISO
                self.cursor.execute("""
                    INSERT INTO iqc_results (lot_id, date, value, level, note)
                    VALUES (?, ?, ?, ?, ?)
                """, (lot_id, run_date_iso, value, level, f"Import từ máy {row.get('Máy xét nghiệm', 'Excel')}"))
                
                success_count += 1

            except Exception as e:
                errors.append(f"Lỗi dòng {row.get('Tên xét nghiệm', 'N/A')}: {str(e)}")

        self.conn.commit()
        return success_count, errors

    def get_iqc_results_all_sources(self, test_id):
        """Lấy tất cả kết quả IQC của một xét nghiệm, bao gồm cả nhập tay và máy"""
        query = """
            SELECT 
                r.id, 
                r.date, 
                r.value, 
                r.level, 
                r.note, 
                r.lot_id,
                m.external_name -- Tên máy (nếu có)
            FROM iqc_results r
            INNER JOIN lots l ON r.lot_id = l.id
            LEFT JOIN test_mapping m ON l.test_id = m.test_id 
                 AND (r.note LIKE '%' || m.external_name || '%' OR r.note = 'Import từ Excel')
            WHERE l.test_id = ?
            ORDER BY r.date DESC
        """
        # Lưu ý: Nếu bạn không lưu tên máy vào note, hãy bỏ phần AND trong LEFT JOIN
        # Đơn giản nhất là dùng query dưới đây:
        query = """
            SELECT r.*, l.lot_number
            FROM iqc_results r
            JOIN lots l ON r.lot_id = l.id
            WHERE l.test_id = ?
            ORDER BY r.date DESC
        """
        return pd.read_sql(query, self.conn, params=(test_id,))


    # db_module.py
    def debug_all_iqc_data(self):
        """Truy vấn không lọc để tìm dữ liệu thất lạc"""
        query = """
            SELECT 
                r.id, 
                l.lot_number, 
                t.name AS test_name, 
                r.date, 
                r.value, 
                r.level, 
                r.note
            FROM iqc_results r
            LEFT JOIN lots l ON r.lot_id = l.id
            LEFT JOIN tests t ON l.test_id = t.id
            ORDER BY r.date DESC
        """
        return pd.read_sql(query, self.conn)

    def add_mapping(self, test_id, external_name):
        """Lưu một liên kết mapping mới"""
        self.cursor.execute(
            "INSERT OR REPLACE INTO test_mapping (test_id, external_name) VALUES (?, ?)",
            (test_id, external_name)
        )
        self.conn.commit()

    def get_all_mappings(self):
        """Lấy danh sách mapping kèm tên xét nghiệm nội bộ để hiển thị"""
        query = """
            SELECT 
                m.id, 
                t.name as internal_name, 
                m.external_name,
                m.test_id
            FROM test_mapping m
            JOIN tests t ON m.test_id = t.id
        """
        return pd.read_sql(query, self.conn)

    def update_mapping(self, mapping_id, new_external_name):
        """Cập nhật lại tên máy cho một mapping đã tồn tại"""
        self.cursor.execute(
            "UPDATE test_mapping SET external_name = ? WHERE id = ?",
            (new_external_name, mapping_id)
        )
        self.conn.commit()

    def delete_mapping(self, mapping_id):
        """Xóa bỏ một liên kết mapping"""
        self.cursor.execute("DELETE FROM test_mapping WHERE id = ?", (mapping_id,))
        self.conn.commit()

    def get_unmapped_tests(self, excel_test_names):
        """
        excel_test_names: List các tên xét nghiệm lấy từ cột 'Tên xét nghiệm' của file Excel
        Trả về: List các tên chưa có trong bảng test_mapping
        """
        # Lấy tất cả tên đã được map trong DB
        self.cursor.execute("SELECT external_name FROM test_mapping")
        mapped_names = [row[0] for row in self.cursor.fetchall()]
        
        # Tìm sự khác biệt
        unmapped = [name for name in excel_test_names if name not in mapped_names]
        
        # Loại bỏ trùng lặp trong danh sách trả về
        return list(set(unmapped))
    
    # --- QUẢN LÝ EQA DATA ---
    def add_eqa(self, data):
        try:
            query = """
                INSERT INTO eqa_results (test_id, date, lab_value, ref_value, sd_group, sdi, program_name)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """
            params = (
                data['test_id'], data['date'], data['lab_value'], 
                data['ref_value'], data['sd_group'], data['sdi'], data['program_name']
            )
            self.cursor.execute(query, params)
            self.conn.commit()
            return True
        except Exception as e:
            print(f"Error adding EQA: {e}")
            return False
    def get_eqa_data(self, test_id):
        # Đổi eqa_data thành eqa_results
        query = "SELECT * FROM eqa_results WHERE test_id = ? ORDER BY date ASC"
        return pd.read_sql_query(query, self.conn, params=(test_id,))
        
    def delete_eqa(self, eqa_id):
        try:
            # Ép kiểu eqa_id về int một lần nữa cho chắc chắn
            clean_id = int(eqa_id)
            self.cursor.execute("DELETE FROM eqa_results WHERE id = ?", (clean_id,))
            self.conn.commit()
            
            # Kiểm tra xem có dòng nào thực sự bị xóa không
            if self.cursor.rowcount > 0:
                return True
            else:
                print(f"Không tìm thấy ID {clean_id} trong DB")
                return False
        except Exception as e:
            print(f"Lỗi SQL: {e}")
            return False

    def update_eqa(self, eqa_id, data):
        if not data: return False
        try:
            # SỬA TẠI ĐÂY: Đổi eqa_data thành eqa_results
            set_parts = [f"{key} = ?" for key in data.keys()]
            values = list(data.values()) + [eqa_id]
            self.cursor.execute(f"UPDATE eqa_results SET {', '.join(set_parts)} WHERE id = ?", values)
            self.conn.commit()
            return True
        except Exception as e:
            print(f"Lỗi cập nhật DB: {e}") # In lỗi ra console để kiểm tra
            return False
   
    def import_eqa_from_dataframe(self, df):
        success_count = 0
        errors = []
        # Làm sạch tên cột: xóa khoảng trắng và chuyển về chữ thường để so sánh
        df.columns = [str(c).strip() for c in df.columns]

        # NHẬN DIỆN CỘT NGOÀI VÒNG LẶP (Để tăng hiệu suất)
        cols = df.columns
        lab_col = next((c for c in cols if any(k in c.lower() for k in ['phòng xét nghiệm', 'kết quả', 'lab', 'pxn'])), None)
        ref_col = next((c for c in cols if any(k in c.lower() for k in ['mục tiêu', 'tham chiếu', 'target', 'ref'])), None)
        # Mở rộng từ khóa tìm kiếm SD để không bị bỏ sót
        sd_col = next((c for c in cols if 'sd' in c.lower() or 'độ lệch' in c.lower()), None)
        name_col = next((c for c in cols if 'tên' in c.lower() and 'nghiệm' in c.lower()), None)
        prog_col = next((c for c in cols if any(k in c.lower() for k in ['chương trình', 'mã', 'đợt', 'program'])), None)
        date_col = next((c for c in cols if 'ngày' in c.lower()), None)

        for index, row in df.iterrows():
            try:
                # 1. Kiểm tra các cột bắt buộc
                if not name_col or not lab_col or not ref_col:
                    errors.append(f"Dòng {index+2}: Thiếu cột Tên xét nghiệm hoặc Kết quả/Tham chiếu.")
                    continue

                # 2. Lấy giá trị và xử lý kiểu dữ liệu
                ext_name = str(row[name_col]).strip()
                lab_val = pd.to_numeric(row[lab_col], errors='coerce')
                ref_val = pd.to_numeric(row[ref_col], errors='coerce')
                # Xử lý SD Nhóm: nếu trống hoặc lỗi thì để 0 thay vì None để tránh lỗi hiển thị
                sd_group = pd.to_numeric(row[sd_col], errors='coerce') if sd_col else 0
                if pd.isna(sd_group): sd_group = 0
                
                program_name = str(row[prog_col]) if prog_col and not pd.isna(row[prog_col]) else "EQA Import"

                # 3. Tính toán SDI (Z-Score)
                sdi_val = 0.0
                if sd_group and sd_group > 0:
                    sdi_val = (lab_val - ref_val) / sd_group

                # 4. Mapping và lấy test_id
                self.cursor.execute("SELECT test_id FROM test_mapping WHERE external_name = ?", (ext_name,))
                mapping_res = self.cursor.fetchone()
                if not mapping_res:
                    errors.append(f"Dòng {index+2}: Chưa map tên '{ext_name}'")
                    continue
                
                test_id = mapping_res[0]
                
                # Xử lý ngày tháng an toàn
                if date_col and not pd.isna(row[date_col]):
                    res_date = pd.to_datetime(row[date_col]).strftime('%Y-%m-%d')
                else:
                    res_date = datetime.now().strftime('%Y-%m-%d')

                # 5. LỆNH INSERT (Đảm bảo lưu đủ sd_group và sdi)
                self.cursor.execute("""
                    INSERT INTO eqa_results (test_id, date, lab_value, ref_value, sd_group, sdi, program_name)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (test_id, res_date, float(lab_val), float(ref_val), float(sd_group), float(sdi_val), program_name))
                
                success_count += 1
            except Exception as e:
                errors.append(f"Lỗi dòng {index+2}: {str(e)}")

        self.conn.commit()
        return success_count, errors
    def upgrade_eqa_table(self):
            """Hàm này đảm bảo bảng luôn có cột sd_group và sdi"""
            try:
                # Thử thêm cột, nếu đã có rồi SQLite sẽ báo lỗi và nhảy vào khối 'except'
                self.cursor.execute("ALTER TABLE eqa_results ADD COLUMN sd_group REAL")
                self.cursor.execute("ALTER TABLE eqa_results ADD COLUMN sdi REAL")
                self.conn.commit()
            except sqlite3.OperationalError:
                # Lỗi này xảy ra khi cột đã tồn tại, chúng ta có thể bỏ qua
                pass

    # --- HÀM QUẢN LÝ CÀI ĐẶT CHUNG (Settings) ---
    def get_setting(self, key, default=None):
        """Lấy giá trị của một cài đặt."""
        sql = "SELECT value FROM settings WHERE key = ?"
        result = self.cursor.execute(sql, (key,)).fetchone()
        return result[0] if result else default
    def calculate_rms_bias(df_eqa):
        """Tính toán RMS Bias từ lịch sử EQA (ISO/TS 20914)"""
        if df_eqa is None or df_eqa.empty or len(df_eqa) < 2:
            return 0.0
        # Tính Bias % cho từng kỳ: ((Lab - Ref) / Ref) * 100
        biases = ((df_eqa['lab_value'] - df_eqa['ref_value']) / df_eqa['ref_value']) * 100
        rms_bias_pct = np.sqrt(np.mean(biases**2))
        return rms_bias_pct

    def get_mu_target_value(standard, test_data, sub_type=None):
        """Trả về mục tiêu MAU (%) dựa trên tiêu chuẩn lựa chọn"""
        if standard == "BV (Biological Variation)":
            cvi = float(test_data.get('cvi', 0.0))
            cvg = float(test_data.get('cvg', 0.0))
            if cvi == 0: return float(test_data.get('tea', 10.0))
            # Công thức TEa/MAU dựa trên Milan Model
            if sub_type == "Tối ưu": return 0.25 * cvi + 1.65 * (0.125 * np.sqrt(cvi**2 + cvg**2))
            if sub_type == "Tối thiểu": return 0.75 * cvi + 1.65 * (0.375 * np.sqrt(cvi**2 + cvg**2))
            return 0.5 * cvi + 1.65 * (0.25 * np.sqrt(cvi**2 + cvg**2)) # Mong muốn
        
        elif standard == "CLIA": return float(test_data.get('clia_limit', 10.0))
        elif standard == "RCPA": return float(test_data.get('rcpa_limit', 8.0))
        return float(test_data.get('tea', 10.0))
    def update_mu_review(self, test_id, review_date): # Thêm 'self' ở đây
            """Cập nhật ngày xem xét MU cuối cùng cho xét nghiệm"""
            try:
                # Sử dụng self.conn hoặc self.create_connection() tùy theo cấu trúc class của bạn
                cursor = self.conn.cursor() 
                sql = "UPDATE tests SET last_mu_review = ? WHERE id = ?"
                cursor.execute(sql, (review_date, test_id))
                self.conn.commit()
                return True
            except Exception as e:
                print(f"Error updating MU review: {e}")
                return False

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
