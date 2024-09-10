import pandas as pd

# Đọc dữ liệu từ file Excel gốc
input_file = 'C:/Users/Admin/Downloads/giaphantich.xlsx'  # Thay thế bằng tên file của bạn
df = pd.read_excel(input_file, sheet_name='Giaphantich')

# Tạo một DataFrame rỗng để chứa dữ liệu mới
new_data = []

# Lặp qua từng hàng trong DataFrame gốc
for index, row in df.iterrows():
    macn = row['Macn']
    mahang = row['Mahang']
    kichthuoc = row['Kichthuoc']
    dvt = row['dvt']
    Khunggiaban = row['Khunggiaban']
    Discount_rate = row['Discount_rate']
    Discount = row['Discount']

    # Danh sách các loại giá và giá
    price_types = {
        'GBLL1KGT': row['GBLL1KGT'],
        'GBLL1KGC': row['GBLL1KGC'],
        'Giachinhsachkgc': row['Giachinhsachkgc'],
        'Giachinhsachkgt': row['Giachinhsachkgt'],
        'Giaomkho200': row['Giaomkho200'],
        'Giasicont': row['Giasicont']
    }

    # Thêm dữ liệu cho mỗi loại giá
    for loai, gia in price_types.items():
        # Chỉ thêm dữ liệu nếu giá không phải là NaN
        if pd.notna(gia):
            new_data.append([
                macn, mahang, kichthuoc, dvt, Khunggiaban,Discount_rate, Discount, loai, gia
            ])

# Chuyển đổi dữ liệu mới thành DataFrame
new_df = pd.DataFrame(new_data,
                      columns=['Macn', 'Mahang', 'Kichthuoc', 'dvt',
                               'Khunggiaban', 'Discount_rate',
                               'Discount', 'Loai', 'GIA'])

# Lưu DataFrame mới vào file Excel
output_file = 'C:/Users/Admin/Downloads/giaphantich_demo1.xlsx'  # Thay thế bằng tên file bạn muốn lưu
new_df.to_excel(output_file, index=False)
