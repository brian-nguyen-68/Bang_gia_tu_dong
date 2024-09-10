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
    dacdiem = row['Dacdiem']
    kichthuoc = row['Kichthuoc']
    dvt = row['dvt']
    khunggiaban = row['Khunggiaban']
    Tuden = row['Tuden']
    Pl_banggia = row['Pl_banggia']
    Pl_oplat = row['Pl_oplat']
    Pl_phankhucgach = row['Pl_phankhucgach']

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
                macn, mahang, dacdiem, kichthuoc, dvt, khunggiaban, Pl_banggia, Tuden,
                Pl_oplat, Pl_phankhucgach, loai, gia
            ])

# Chuyển đổi dữ liệu mới thành DataFrame
new_df = pd.DataFrame(new_data,
                      columns=['Macn', 'Mahang', 'Dacdiem', 'Kichthuoc', 'dvt',
                               'Khunggiaban', 'Pl_banggia', 'Tuden', 'Pl_oplat',
                               'Pl_phankhucgach', 'Loai', 'GIA'])

# Lưu DataFrame mới vào file Excel
output_file = 'C:/Users/Admin/Downloads/giaphantich_demo1.xlsx'  # Thay thế bằng tên file bạn muốn lưu
new_df.to_excel(output_file, index=False)
