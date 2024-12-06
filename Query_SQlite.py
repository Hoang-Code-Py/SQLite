import sqlite3
import pandas as pd

def export_sql_to_excel(sql_file, query, output_file):
    """
    Xuất dữ liệu từ SQLite sang Excel nhanh chóng.
    
    Args:
        sql_file (str): Đường dẫn file SQLite.
        query (str): Câu lệnh SQL để lọc dữ liệu.
        output_file (str): Đường dẫn file Excel xuất ra.
    """
    # Kết nối đến SQLite
    conn = sqlite3.connect(sql_file)
    
    try:
        # Thực thi query và đọc dữ liệu vào DataFrame
        print("Đang truy vấn dữ liệu từ SQLite...")
        df = pd.read_sql_query(query, conn)
        
        # Ghi dữ liệu vào file Excel
        print("Đang ghi dữ liệu vào Excel...")
        df.to_excel(output_file, index=False)
        print(f"Hoàn tất! Dữ liệu đã được ghi vào: {output_file}")
    except Exception as e:
        print(f"Lỗi: {e}")
    finally:
        conn.close()

# Sử dụng
sql_file = r"D:\Code\SQL Application\Convert Excel To SQL\Output\Compare.db"
query = '''
SELECT distinct Manufacturer
from Combined_Table WHERE Feature_Group like 'Function in details' 
and Function like '%Most popular%'
''' 

output_file = r"D:\Code\SQL Application\Convert Excel To SQL\Output\Query_SQLi.xlsx"
export_sql_to_excel(sql_file, query, output_file)
