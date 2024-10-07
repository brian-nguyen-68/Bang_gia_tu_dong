import pandas as pd
import pyodbc

# Đọc file Excel
df = pd.read_excel('G:/.shortcut-targets-by-id/17iVJ0hhkx3ZIc5d3mWBvN1A4R-QCnV4r/06.2024/01.GIABAN/BangGiaTuDong/BangDoKhungTrenDuoi.xlsx')

# Lấy tên các cột chứa giá trị (bắt đầu từ cột thứ tư)
value_columns = df.columns[3:]

# Tạo danh sách để lưu kết quả
result = []

# Chuyển đổi dữ liệu
for index, row in df.iterrows():
    for col in value_columns:
        result.append({
            'BGTD_QuiCachBangGia': row['BGTD_QuiCachBangGia'],
            'Macn': int(col),  # Chuyển đổi tên cột sang số nguyên
            'Tuden': int(row[col]) if pd.notna(row[col]) else 0  # Chuyển đổi sang int, thay NaN bằng 0
        })

# Tạo DataFrame từ kết quả
result_df = pd.DataFrame(result)

# Kiểm tra kiểu dữ liệu
print(result_df.dtypes)  # In ra kiểu dữ liệu để kiểm tra

# Kết nối với SQL Server bằng Windows Authentication
conn = pyodbc.connect('DRIVER={SQL Server};'
                      'SERVER=XPS7590;'
                      'DATABASE=UNIS_DATA;'
                      'Trusted_Connection=yes;')  # Sử dụng Windows Authentication

# Tạo cursor
cursor = conn.cursor()

# Xóa bảng nếu đã tồn tại
cursor.execute("IF OBJECT_ID('dbo.bangdokhungtrenduoi', 'U') IS NOT NULL DROP TABLE dbo.bangdokhungtrenduoi")
conn.commit()

# Tạo bảng mới
cursor.execute('''
CREATE TABLE bangdokhungtrenduoi (
    BGTD_QuiCachBangGia NVARCHAR(50),
    Macn INT,
    Tuden INT
)
''')
conn.commit()

# Lưu DataFrame vào SQL Server
for index, row in result_df.iterrows():
    cursor.execute('''
        INSERT INTO bangdokhungtrenduoi (BGTD_QuiCachBangGia, Macn, Tuden)
        VALUES (?, ?, ?)''',
        row['BGTD_QuiCachBangGia'], row['Macn'], row['Tuden'])

conn.commit()

# Đóng kết nối
conn.close()
