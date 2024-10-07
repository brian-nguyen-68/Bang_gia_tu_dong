import pandas as pd
import os

# Đường dẫn đến thư mục chứa các file Excel
folder_path = 'G:/.shortcut-targets-by-id/13BJtPs_07_3RGiW58ePrNXtMl6-oIwgn/02.Folder_Gia_Nguon/Gia2024/Pushgiavanchuyen2024/UNILUX'  # Thay 'duongdan/thu_muc_excel' bằng đường dẫn đến thư mục của bạn

# Danh sách để lưu các DataFrame đã chỉnh sửa
df_list = []

# Lặp qua từng file trong thư mục
for file_name in os.listdir(folder_path):
    if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
        file_path = os.path.join(folder_path, file_name)  # Tạo đường dẫn đầy đủ

        # Đọc file Excel
        excel_file = pd.ExcelFile(file_path)

        # Lặp qua từng sheet trong file Excel
        for sheet_name in excel_file.sheet_names:
            # Đọc dữ liệu từ sheet
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            # Thêm cột Tenchinhanh với tên sheet
            df['Tenchinhanh'] = sheet_name

            # Đổi vị trí cột Tenchinhanh về đầu bảng
            df = df[['Tenchinhanh'] + [col for col in df.columns if col != 'Tenchinhanh']]

            # Thêm DataFrame đã chỉnh sửa vào danh sách
            df_list.append(df)

# Gộp tất cả các DataFrame lại thành một DataFrame lớn
result_df = pd.concat(df_list, ignore_index=True)

# Lưu lại DataFrame vào file Excel mới
result_df.to_excel('C:/Users/Admin/Downloads/unilux.xlsx', index=False)  # Thay 'duongdan/file_moi.xlsx' bằng đường dẫn bạn muốn lưu
