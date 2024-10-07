import pyodbc

# Thông tin kết nối
server = 'XPS7590'
database = 'UNIS_DATA'

# Tạo kết nối với Windows Authentication
conn = pyodbc.connect(f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes;')

# Thực hiện truy vấn
query = """
UPDATE A
SET 
    A.[Tiền Giảm giá] = CAST(B.[Tiền Giảm giá] AS DECIMAL(37, 4)) * CAST(A.[Quydoidonvi] AS DECIMAL(37, 4)),
    A.[% Giảm giá] = B.[% Giảm giá],
    A.[Đơn giá sau chiết khấu] = CAST(A.[Đơn giá niêm yết] AS DECIMAL(37, 4)) - 
        (CAST(A.[Đơn giá niêm yết] AS DECIMAL(37, 4)) * B.[% Giảm giá] + 
        CAST(B.[Tiền Giảm giá] AS DECIMAL(37, 4)) * CAST(A.[Quydoidonvi] AS DECIMAL(37, 4)))
FROM file_pushgia3 AS A
JOIN Discount AS B ON A.[Macn] = B.[Macn] AND A.[Mahang] = B.[Mahang]
WHERE A.[Vùng bán hàng] = B.[Vùng bán hàng] 
  AND A.[LOAI BAN HANG] = B.[LOAI BAN HANG] 
  AND A.[HTTT] = B.[HTTT];
"""

# Thực hiện cập nhật
with conn.cursor() as cursor:
    cursor.execute(query)
    conn.commit()  # Lưu các thay đổi

# Đóng kết nối
conn.close()
