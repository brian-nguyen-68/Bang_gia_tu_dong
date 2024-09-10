import pandas as pd
import pyodbc

# Kết nối đến SQL Server
conn_str = (
    r'DRIVER={ODBC Driver 17 for SQL Server};'
    r'SERVER=XPS7590;'  # Thay thế với tên máy chủ của bạn
    r'DATABASE=UNIS_DATA;'  # Thay thế với tên cơ sở dữ liệu của bạn
    r'TRUSTED_CONNECTION=yes;'
)
conn = pyodbc.connect(conn_str)

# Đọc dữ liệu từ bảng gốc testkg2
query = "SELECT * FROM testkg2"
df = pd.read_sql(query, conn)

# In tất cả các tên cột để xác nhận
print("Tên các cột trong bảng:", df.columns.tolist())

# Sắp xếp dữ liệu theo cột Khunggiaban
df_sorted = df.sort_values(by=['Macn','Kichthuoc', 'Dacdiem'])

# Tạo cột Khunggia_test dựa trên thứ hạng của Khunggiaban
# Sử dụng phương pháp phân nhóm để gán cùng giá trị KG cho các giá trị giống nhau
df_sorted['Khunggia_test'] = df_sorted.groupby(['Macn','Kichthuoc', 'Dacdiem']).ngroup() + 1

# Chuyển đổi thứ hạng thành định dạng 'KG' + số thứ hạng
df_sorted['Khunggia_test'] = df_sorted['Khunggia_test'].apply(lambda x: f'KG{x}')

# Tạo bảng mới test_khunggia nếu đã tồn tại thì xóa và tạo lại
cursor = conn.cursor()

# Xóa bảng nếu đã tồn tại
drop_table_query = "IF OBJECT_ID('test_khunggia', 'U') IS NOT NULL DROP TABLE test_khunggia"
cursor.execute(drop_table_query)
conn.commit()

# Tạo bảng mới
create_table_query = """
CREATE TABLE test_khunggia (
    Macn INT,
    Masp NVARCHAR(250),
    Mahang NVARCHAR(250),
    Dacdiem NVARCHAR(250),
    Kichthuoc NVARCHAR(250),
    dvt NVARCHAR(250),
    Khunggiaban NVARCHAR(250),
    GBLL1KGT INT,
    GBLL1KGC INT,
    Giachinhsachkgt INT,
    Giachinhsachkgc INT,
    Giaomkho200 INT,
    Giasicont INT,
    Khunggia_test NVARCHAR(250)
)
"""
cursor.execute(create_table_query)
conn.commit()

# Chèn dữ liệu vào bảng mới test_khunggia
insert_query = """
INSERT INTO test_khunggia (Macn, Masp, Mahang, Dacdiem, Kichthuoc, dvt, Khunggiaban, GBLL1KGT, GBLL1KGC, Giachinhsachkgt, Giachinhsachkgc, Giaomkho200, Giasicont, Khunggia_test)
VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
"""

# Chèn từng dòng vào bảng mới
for index, row in df_sorted.iterrows():
    print(f'Inserting row {index} with Masp={row["Masp"]}')  # Debugging line
    cursor.execute(insert_query,
                   row['Macn'], row['Masp'], row['Mahang'], row['Dacdiem'], row['Kichthuoc'],
                   row['dvt'], row['Khunggiaban'], row['GBLL1KGT'], row['GBLL1KGC'],
                   row['Giachinhsachkgt'], row['Giachinhsachkgc'], row['Giaomkho200'],
                   row['Giasicont'], row['Khunggia_test'])

conn.commit()
cursor.close()
conn.close()
