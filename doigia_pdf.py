import pandas as pd

# Đọc dữ liệu từ các file Excel
b1 = pd.read_excel('C:/Users/Admin/Downloads/060.xlsx', sheet_name='Banle', header=6)  # Bảng giá cũ
b2 = pd.read_excel('C:/Users/Admin/Downloads/giamoi.xlsx', sheet_name='Banle', header=0)  # Bảng giá mới

# Tạo từ điển ánh xạ giá từ bảng b2
value_map = b2.set_index(['Masanpham', 'Key'])['Value'].to_dict()

# Danh sách các dòng mới với giá cập nhật
new_rows = []

# Danh sách các dòng giữ nguyên
remaining_rows = []

# Duyệt qua từng dòng của bảng giá cũ
for index, row in b1.iterrows():
    if pd.notna(row['Mã sản phẩm']):
        masanpham_list = [mp.strip() for mp in row['Mã sản phẩm'].split(',')]
        updated_masanpham_list = []
        for masanpham in masanpham_list:
            value1 = value_map.get((masanpham, 'Value1'), None)
            value2 = value_map.get((masanpham, 'Value2'), None)

            if value1 and value1 != row['> 50 M2']:
                new_rows.append({
                    'STT': row['STT'],
                    'DO': row['DO'],
                    'Kích thước': row['Kích thước'],
                    'Đặc điểm': row['Đặc điểm'],
                    'Khung Giá': row['Khung Giá'],
                    'Mã sản phẩm': masanpham,
                    'ĐVT': row['ĐVT'],
                    '> 50 M2': value1,
                    '< 50 M2': row['< 50 M2'],
                    'Ghi chú': 'Giá mới'
                })

            if value2 and value2 != row['< 50 M2']:
                new_rows.append({
                    'STT': row['STT'],
                    'DO': row['DO'],
                    'Kích thước': row['Kích thước'],
                    'Đặc điểm': row['Đặc điểm'],
                    'Khung Giá': row['Khung Giá'],
                    'Mã sản phẩm': masanpham,
                    'ĐVT': row['ĐVT'],
                    '> 50 M2': row['> 50 M2'],
                    '< 50 M2': value2,
                    'Ghi chú': 'Giá mới'
                })
            else:
                updated_masanpham_list.append(masanpham)

        if updated_masanpham_list:
            remaining_rows.append({
                'STT': row['STT'],
                'DO': row['DO'],
                'Kích thước': row['Kích thước'],
                'Đặc điểm': row['Đặc điểm'],
                'Khung Giá': row['Khung Giá'],
                'Mã sản phẩm': ', '.join(updated_masanpham_list),
                'ĐVT': row['ĐVT'],
                '> 50 M2': row['> 50 M2'],
                '< 50 M2': row['< 50 M2'],
                'Ghi chú': row['Ghi chú']
            })

# Tạo DataFrame từ danh sách các dòng mới và các dòng giữ nguyên
df_new_rows = pd.DataFrame(new_rows)
df_remaining_rows = pd.DataFrame(remaining_rows)

# Ghi dữ liệu vào file Excel mới
with pd.ExcelWriter('C:/Users/Admin/Downloads/giamoi_sau_cap_nhat.xlsx', engine='openpyxl') as writer:
    df_remaining_rows.to_excel(writer, sheet_name='Banle', index=False, startrow=0)
    df_new_rows.to_excel(writer, sheet_name='Banle', index=False, startrow=len(df_remaining_rows)+2)
