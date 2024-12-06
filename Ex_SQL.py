import pandas as pd
import sqlite3
import os

def convert_SQL(file_Excel, file_SQL):
    """
    Gộp dữ liệu từ nhiều file Excel vào một bảng SQLite, thêm cột 'Filename' và sử dụng cột 'Manufacturer' làm khóa chính.
    
    Args:
        file_Excel (str): Đường dẫn đến thư mục chứa các file Excel.
        file_SQL (str): Đường dẫn đến file SQLite sẽ tạo ra.
    """
    # Kết nối hoặc tạo file SQLite
    sql_dir = os.path.dirname(file_SQL)
    if not os.path.exists(sql_dir):
        os.makedirs(sql_dir)
        print(f"Đã tạo thư mục: {sql_dir}")
    connect_SQL = sqlite3.connect(file_SQL)

    combined_table_name = "Combined_Table"  # Tên bảng duy nhất
    try:
        # Duyệt qua tất cả các file trong thư mục
        for root, folders, files in os.walk(file_Excel):
            for file in files:
                filename = os.path.join(root, file)
                
                # Kiểm tra file Excel
                if filename.endswith('.xlsx') and not filename.startswith('~$'):
                    try:
                        print(f"Đang xử lý file: {filename}")
                        # Đọc tất cả các sheet trong file Excel
                        excel_data = pd.read_excel(filename, sheet_name=None,header=0)
                        
                        for sheet_name, df in excel_data.items():
                            # Làm sạch tên cột
                            df.columns = [col.replace(" ", "_") for col in df.columns]
                            
                            # Đảm bảo cột 'Manufacturer' tồn tại
                            if 'Manufacturer' not in df.columns:
                                print(f"Sheet '{sheet_name}' trong file '{file}' không có cột 'Manufacturer'. Bỏ qua.")
                                continue
                            
                            # Thêm cột 'Filename' chứa tên file hiện tại
                            df['Filename'] = os.path.basename(filename)
                            
                            # Đặt cột 'Manufacturer' làm index
                            df.set_index('Manufacturer', inplace=True)
                            
                            # Gộp dữ liệu vào bảng SQLite
                            df.to_sql(combined_table_name, connect_SQL, if_exists='append', index=True, index_label='Manufacturer')
                            print(f"Dữ liệu từ sheet '{sheet_name}' trong file '{file}' đã được thêm vào bảng '{combined_table_name}'.")

                    except Exception as e:
                        print(f"Lỗi khi xử lý file {filename}: {e}")
    
    except Exception as e:
        print(f"Lỗi tổng quát: {e}")
    
    finally:
        # Đóng kết nối SQLite
        connect_SQL.close()
        print("Đã đóng kết nối với SQLite.")

def main():
    print("Bắt đầu chuyển đổi Excel sang SQLite...")
    file_Excel = r'D:\Code\SQL Application\Convert Excel To SQL\Input'
    file_SQL = r'D:\Code\SQL Application\Convert Excel To SQL\Output\Manufacturer.db'
    convert_SQL(file_Excel, file_SQL)
    print("Hoàn thành quá trình chuyển đổi.")

if __name__ == "__main__":
    main()
