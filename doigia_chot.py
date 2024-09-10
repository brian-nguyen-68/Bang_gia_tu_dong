import pandas as pd
import re

# Đọc dữ liệu từ file Excel
file_path = 'C:/Users/Admin/Downloads/Z00_FORM LAM GIA.xlsx'

# Xử lý sheet 'giabanle'
df_giabanle = pd.read_excel(file_path, sheet_name='giabanle', header=[0, 1])
df_giabanle.columns = ['_'.join(col).strip() for col in df_giabanle.columns.values]
print("Cột hiện tại trong 'giabanle':", df_giabanle.columns)

# Xác định các cột loại hàng
pattern = r'^\d{3}_GBLL1K'  # Biểu thức chính quy để tìm các cột bắt đầu bằng 3 chữ số và '_GBLL1K'
loai_cols = [col for col in df_giabanle.columns if re.match(pattern, col)]

# Chuyển đổi dữ liệu từ dạng bảng rộng thành dạng bảng dài
id_vars_giabanle = ['Mahang_Mahang', 'dvt_dvt']
for var in id_vars_giabanle:
    if var not in df_giabanle.columns:
        print(f"Cột '{var}' không có trong DataFrame.")
        exit()

melted_df_giabanle = df_giabanle.melt(id_vars=id_vars_giabanle, value_vars=loai_cols, var_name='Loai', value_name='GIA')
melted_df_giabanle['Macn'] = melted_df_giabanle['Loai'].apply(lambda x: x.split('_')[0])
melted_df_giabanle['Loai'] = melted_df_giabanle['Loai'].apply(lambda x: x.split('_')[1])
melted_df_giabanle.rename(columns={'Mahang_Mahang': 'Mahang', 'dvt_dvt': 'dvt'}, inplace=True)
melted_df_giabanle = melted_df_giabanle[['Macn', 'Mahang', 'dvt', 'Loai', 'GIA']]

# Xử lý sheet 'giachinhsach'
df_giachinhsach = pd.read_excel(file_path, sheet_name='giachinhsach', header=[0, 1])
df_giachinhsach.columns = ['_'.join(col).strip() for col in df_giachinhsach.columns.values]
print("Cột hiện tại trong 'giachinhsach':", df_giachinhsach.columns)

# Xác định các cột loại hàng trong 'giachinhsach'
pattern_giachinhsach = r'^\d{3}_Giachinhsach'  # Biểu thức chính quy để tìm các cột bắt đầu bằng 3 chữ số và '_Giachinhsach'
loai_cols_giachinhsach = [col for col in df_giachinhsach.columns if re.match(pattern_giachinhsach, col)]

# Chuyển đổi dữ liệu từ dạng bảng rộng thành dạng bảng dài
id_vars_giachinhsach = ['Mahang_Mahang', 'dvt_dvt']
for var in id_vars_giachinhsach:
    if var not in df_giachinhsach.columns:
        print(f"Cột '{var}' không có trong DataFrame.")
        exit()

melted_df_giachinhsach = df_giachinhsach.melt(id_vars=id_vars_giachinhsach, value_vars=loai_cols_giachinhsach, var_name='Loai', value_name='GIA')
melted_df_giachinhsach['Macn'] = melted_df_giachinhsach['Loai'].apply(lambda x: x.split('_')[0])
melted_df_giachinhsach['Loai'] = melted_df_giachinhsach['Loai'].apply(lambda x: x.split('_')[1])
melted_df_giachinhsach.rename(columns={'Mahang_Mahang': 'Mahang', 'dvt_dvt': 'dvt'}, inplace=True)
melted_df_giachinhsach = melted_df_giachinhsach[['Macn', 'Mahang', 'dvt', 'Loai', 'GIA']]

# Xử lý sheet 'giasicont'
df_giasicont = pd.read_excel(file_path, sheet_name='giasicont', header=[0, 1])
df_giasicont.columns = ['_'.join(col).strip() for col in df_giasicont.columns.values]
print("Cột hiện tại trong 'giasicont':", df_giasicont.columns)

# Xác định các cột loại hàng trong 'giasicont'
pattern_giasicont = r'^\d{3}_Giasicont'  # Biểu thức chính quy để tìm các cột bắt đầu bằng 3 chữ số và '_Giasicont'
loai_cols_giasicont = [col for col in df_giasicont.columns if re.match(pattern_giasicont, col)]

# Chuyển đổi dữ liệu từ dạng bảng rộng thành dạng bảng dài
id_vars_giasicont = ['Mahang_Mahang', 'dvt_dvt']
for var in id_vars_giasicont:
    if var not in df_giasicont.columns:
        print(f"Cột '{var}' không có trong DataFrame.")
        exit()

melted_df_giasicont = df_giasicont.melt(id_vars=id_vars_giasicont, value_vars=loai_cols_giasicont, var_name='Loai', value_name='GIA')
melted_df_giasicont['Macn'] = melted_df_giasicont['Loai'].apply(lambda x: x.split('_')[0])
melted_df_giasicont['Loai'] = melted_df_giasicont['Loai'].apply(lambda x: x.split('_')[1])
melted_df_giasicont.rename(columns={'Mahang_Mahang': 'Mahang', 'dvt_dvt': 'dvt'}, inplace=True)
melted_df_giasicont = melted_df_giasicont[['Macn', 'Mahang', 'dvt', 'Loai', 'GIA']]

# Xử lý sheet 'giaomkho200'
df_giaomkho200 = pd.read_excel(file_path, sheet_name='giaomkho200', header=[0, 1])
df_giaomkho200.columns = ['_'.join(col).strip() for col in df_giaomkho200.columns.values]
print("Cột hiện tại trong 'giasicont':", df_giaomkho200.columns)

# Xác định các cột loại hàng trong 'giaomkho200'
pattern_giaomkho200 = r'^\d{3}_Giaomkho200'  # Biểu thức chính quy để tìm các cột bắt đầu bằng 3 chữ số và '_Giasicont'
loai_cols_giaomkho200 = [col for col in df_giaomkho200.columns if re.match(pattern_giaomkho200, col)]

# Chuyển đổi dữ liệu từ dạng bảng rộng thành dạng bảng dài
id_vars_giaomkho200 = ['Mahang_Mahang', 'dvt_dvt']
for var in id_vars_giaomkho200:
    if var not in df_giaomkho200.columns:
        print(f"Cột '{var}' không có trong DataFrame.")
        exit()

melted_df_giaomkho200 = df_giaomkho200.melt(id_vars=id_vars_giaomkho200, value_vars=loai_cols_giaomkho200, var_name='Loai', value_name='GIA')
melted_df_giaomkho200['Macn'] = melted_df_giaomkho200['Loai'].apply(lambda x: x.split('_')[0])
melted_df_giaomkho200['Loai'] = melted_df_giaomkho200['Loai'].apply(lambda x: x.split('_')[1])
melted_df_giaomkho200.rename(columns={'Mahang_Mahang': 'Mahang', 'dvt_dvt': 'dvt'}, inplace=True)
melted_df_giaomkho200 = melted_df_giaomkho200[['Macn', 'Mahang', 'dvt', 'Loai', 'GIA']]

# Lưu kết quả vào file Excel mới với các sheet tương ứng
output_file_path = 'C:/Users/Admin/Downloads/ketqua.xlsx'
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    melted_df_giabanle.to_excel(writer, sheet_name='giabanle', index=False)
    melted_df_giachinhsach.to_excel(writer, sheet_name='giachinhsach', index=False)
    melted_df_giasicont.to_excel(writer, sheet_name='giasicont', index=False)
    melted_df_giaomkho200.to_excel(writer, sheet_name='giaomkho200', index=False)

print("Dữ liệu đã được lưu vào file Excel mới.")
