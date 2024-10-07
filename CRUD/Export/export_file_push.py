import pandas as pd
import pyodbc
import os

# Thay đổi đường dẫn lưu file
output_directory = r'C:\Users\Admin\Downloads'  # Thay đổi đường dẫn tới thư mục bạn muốn lưu file

# Kết nối đến cơ sở dữ liệu SQL Server với xác thực Windows
conn_str = (
    r'DRIVER={SQL Server};'
    r'SERVER=XPS7590;'
    r'DATABASE=UNIS_DATA;'
    r'Trusted_Connection=yes;'
)

# Kết nối đến cơ sở dữ liệu
conn = pyodbc.connect(conn_str)

# Truy vấn dữ liệu
query = "SELECT *  FROM file_pushgia3"
df = pd.read_sql(query, conn)

# Đóng kết nối
conn.close()

# Lưu file Excel cho từng dòng
for index, row in df.iterrows():
    filename = f"{row['Cong Ty Trien Khai']}_{row['Ten Chi Nhanh']}.xlsx"
    filepath = os.path.join(output_directory, filename)  # Tạo đường dẫn đầy đủ
    row.to_frame().T.to_excel(filepath, index=False, engine='openpyxl')
    print(f"Đã lưu file: {filepath}")
