import pandas as pd
import pyodbc

# Đọc dữ liệu từ file Excel
file_path = 'C:/Users/Admin/Downloads/ketqua.xlsx'
df = pd.read_excel(file_path)

# Kết nối đến SQL Server bằng Windows Authentication
conn_str = (
    r'DRIVER={SQL Server};'
    r'SERVER=XPS7590;'  # Thay đổi tên server của bạn
    r'DATABASE=UNIS_DATA;'  # Thay đổi tên cơ sở dữ liệu của bạn
    r'Trusted_Connection=yes;'  # Sử dụng Windows Authentication
)
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# Cập nhật dữ liệu trong bảng SQL Server
for index, row in df.iterrows():
    sql_update_query = """
    UPDATE gia_ban_le
    SET GIA = ?
    WHERE Macn = ? AND Masp = ? AND ĐVT = ? AND Loai = ?
    """
    values = (row['GIA'], row['Macn'], row['Masp'], row['ĐVT'], row['Loai'])
    cursor.execute(sql_update_query, values)

# Commit các thay đổi và đóng kết nối
conn.commit()
cursor.close()
conn.close()

print("Cập nhật dữ liệu thành công!")
