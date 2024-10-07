import pandas as pd
import os

# Đường dẫn đến thư mục chứa các file Excel
folder_path = 'G:/.shortcut-targets-by-id/13BJtPs_07_3RGiW58ePrNXtMl6-oIwgn/02.Folder_Gia_Nguon/Gia2024/Pushgia2024/PUSH GIA THANG 09.2024/UNICHEMI'

# Danh sách để lưu các DataFrame đã chỉnh sửa
df_list = []

# Lặp qua từng file trong thư mục
for file_name in os.listdir(folder_path):
    if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
        file_path = os.path.join(folder_path, file_name)  # Tạo đường dẫn đầy đủ

        # Đọc file Excel và tìm sheet "BANG GIA PUSH"
        try:
            df = pd.read_excel(file_path, sheet_name='BANG GIA PUSH')
        except ValueError:
            continue  # Nếu không có sheet, bỏ qua file này

        # Lọc dữ liệu: chỉ lấy những dòng không có ĐVT là 'Hộp'
        df_filtered = df[df.iloc[:, 5] != 'Hộp'].copy()  # Sử dụng .copy() để tránh cảnh báo

        # Lấy tên chi nhánh từ tên file (giả sử tên chi nhánh nằm giữa các dấu '_')
        branch_name = file_name.split('_')[1]  # Lấy phần thứ hai trong tên file

        # Thêm cột Tenchinhanh với tên chi nhánh
        df_filtered['Tenchinhanh'] = branch_name

        # Thêm DataFrame đã chỉnh sửa vào danh sách
        df_list.append(df_filtered)

# Gộp tất cả các DataFrame lại thành một DataFrame lớn
if df_list:
    result_df = pd.concat(df_list, ignore_index=True)

    # Lưu lại DataFrame vào file Excel mới
    result_df.to_excel('C:/Users/Admin/Downloads/unichemi_sppt.xlsx', index=False)  # Đường dẫn lưu file mới
else:
    print("Không có dữ liệu nào được lọc.")
