import pandas as pd
import pyodbc
import os

# Kết nối đến SQL Server
conn_str = (
    r'DRIVER={ODBC Driver 17 for SQL Server};'
    r'SERVER=XPS7590;'  # Thay thế với tên máy chủ của bạn
    r'DATABASE=UNIS_DATA;'  # Thay thế với tên cơ sở dữ liệu của bạn
    r'TRUSTED_CONNECTION=yes;'
)
conn = pyodbc.connect(conn_str)

# Đọc dữ liệu từ bảng
query = "SELECT * FROM Phantichgia_state1"
df = pd.read_sql(query, conn)

# Đóng kết nối
conn.close()

# Lấy danh sách các mã chi nhánh
branch_codes = df['Macn'].unique()

# Nhận đường dẫn đầu ra từ người dùng
output_directory = input("Nhập đường dẫn thư mục để lưu file Excel: ")

# Kiểm tra nếu thư mục không tồn tại thì thông báo lỗi và thoát chương trình
if not os.path.exists(output_directory):
    print(f"Thư mục '{output_directory}' không tồn tại.")
    exit()

# Tạo thư mục nếu chưa tồn tại
os.makedirs(output_directory, exist_ok=True)

for branch_code in branch_codes:
    # Lọc dữ liệu theo mã chi nhánh
    branch_df = df[df['Macn'] == branch_code]

    # Đặt tên file dựa trên mã chi nhánh
    file_name = f'{branch_code}.xlsx'

    # Tạo đường dẫn đầy đủ tới file
    file_path = os.path.join(output_directory, file_name)

    # Xuất dữ liệu ra file Excel
    branch_df.to_excel(file_path, index=False, engine='openpyxl')

    print(f'Data for branch code {branch_code} exported to {file_path}')
