import pyodbc

# Thông tin kết nối
server = 'XPS7590'
database = 'UNIS_DATA'

# Tạo kết nối với Windows Authentication
conn = pyodbc.connect(f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes;')

# Thực hiện truy vấn
query = """
USE UNIS_DATA;

----- Tạo bảng file_pushgia1-----
DROP TABLE IF EXISTS file_pushgia1;

-- Truy vấn chính để lấy giá dựa trên các điều kiện ưu tiên
SELECT DISTINCT
    p.[Macn],
    b1.[Cong Ty Trien Khai],
    b1.[Ten Chi Nhanh],
    p.[Mahang],
    p.[Kichthuoc],
    p.[dvt],
    p.[GIA],
    p.[Loai],
    b.[Dac Diem],
    b.[Quydoidonvi],
    b.[Vien/Hop],
	b.[M2/Vien],
	COALESCE(b2.[Tuden], 0) AS Tuden,
	CASE 
        WHEN LEFT(b.[Ma Sale], 2) IN ('31', '32', '33', '34') THEN 
			CONCAT(
				LEFT(b.[Ma Sale], 3), 
				'L1.', 
			SUBSTRING(
					b.[Ma Sale], 
					CHARINDEX('.', b.[Ma Sale], CHARINDEX('.', b.[Ma Sale])) + 1,
					LEN(b.[Ma Sale]) - CHARINDEX('.', b.[Ma Sale], CHARINDEX('.', b.[Ma Sale]) + 1) 
				))
        ELSE 
			CONCAT(
			LEFT(b.[Ma Sale], 3), 
			'L1.', 
			b.[Nhom Gia],
			SUBSTRING(
				b.[Ma Sale], 
				CHARINDEX('.', b.[Ma Sale], CHARINDEX('.', b.[Ma Sale])) + 1,
				LEN(b.[Ma Sale]) - CHARINDEX('.', b.[Ma Sale], CHARINDEX('.', b.[Ma Sale]) + 1) 
			))   
    END AS [Mặt hàng] ,

    1 AS HTTT,
    '0' AS [Biên độ giảm],
    '0' AS [Biên độ tăng],
    CASE 
        WHEN p.Loai IN ('GBLL1KGT', 'GBLL1KGC', 'Giachinhsachkgc', 'Giachinhsachkgt') THEN 1
        WHEN p.Loai = 'Giasicont' THEN 2
		WHEN P.Loai = 'Giacatlo' THEN 3
        WHEN p.Loai = 'Giaomkho200' THEN 4
        ELSE NULL
    END AS [LOAI BAN HANG],
   CAST(N'Hộp' AS nvarchar(255)) AS [ĐVT_1],
    CASE 
        WHEN p.Loai IN ('Giachinhsachkgc', 'Giachinhsachkgt') THEN
            CASE 
                WHEN p.GIA IS NULL OR p.GIA = '0' THEN '0'
                ELSE p.GIA
            END
        ELSE p.GIA
    END AS [Đơn giá niêm yết_chuaquydoi],
    '' AS [Tên hàng hóa, dịch vụ],
	 CAST(NULL AS FLOAT) AS [% Giảm giá],
     CAST(NULL AS DECIMAL(37, 4)) AS [Tiền Giảm giá],
	  b.[Nhom Gia],
    '' AS [Kích thước],
    '' AS [Khung Giá],
    '' AS [Mã SP],
    '' AS [Loại],
    '' AS [Gạch Ốp/Lát],
    '' AS [ĐVT],
    '' AS [TỪ],
    '' AS [ĐẾN],
    '' AS [GIÁ BÁN],
    '' AS [GIÁ CHÍNH SÁCH],
    '' AS [GIÁ SÀN]
INTO file_pushgia1
FROM [UNIS_DATA].[dbo].[file_pdf_gia] p
LEFT JOIN [UNIS_DATA].[dbo].[BangDoMaHang] b ON p.[Mahang] = b.[Ma Sale]
LEFT JOIN 
    [UNIS_DATA].[dbo].[bangdokhungtrenduoi] b2 ON CONCAT(b2.[Macn], b2.[BGTD_QuiCachBangGia]) = CONCAT(p.[Macn], p.[Kichthuoc])
LEFT JOIN [UNIS_DATA].[dbo].[BangDoChiNhanh] b1 ON p.[Macn] = b1.[Ma Chi Nhanh];
----- Tạo bảng file_pushgia2-----
DROP TABLE IF EXISTS file_pushgia2;
SELECT [Cong Ty Trien Khai],
[Ten Chi Nhanh],[Macn],[dvt],[Quydoidonvi],[Vien/Hop],[M2/Vien]
      ,[Mặt hàng],
	  CASE 
           WHEN [LOAI BAN HANG] IN ('1', '3', '4') THEN 1
           WHEN [LOAI BAN HANG] = '2' THEN 2
           ELSE NULL -- hoặc giá trị mặc định khác nếu cần
       END AS [Vùng bán hàng]
      ,[LOAI BAN HANG]
	  ,[HTTT]
	  ,[Tên hàng hóa, dịch vụ]
	  ,[ĐVT_1]
	  ,[Loai]
	  ,CASE 
           WHEN file_pushgia1.[dvt] = 'M' THEN 
               CAST(file_pushgia1.[Quydoidonvi] AS DECIMAL(18, 2)) * 
               CAST(file_pushgia1.[Đơn giá niêm yết_chuaquydoi] AS DECIMAL(18, 2))
           WHEN file_pushgia1.[dvt] = 'V' THEN 
               CAST(file_pushgia1.[Vien/Hop] AS DECIMAL(18, 2)) * 
               CAST(file_pushgia1.[Đơn giá niêm yết_chuaquydoi] AS DECIMAL(18, 2))
           ELSE 
               CAST(file_pushgia1.[Đơn giá niêm yết_chuaquydoi] AS DECIMAL(18, 2))
       END AS [Đơn giá niêm yết]
	  ,[% Giảm giá]
      ,[Tiền Giảm giá]
	  ,CASE 
           WHEN file_pushgia1.[dvt] = 'M' THEN 
               CAST(file_pushgia1.[Quydoidonvi] AS DECIMAL(18, 2)) * 
               CAST(file_pushgia1.[Đơn giá niêm yết_chuaquydoi] AS DECIMAL(18, 2))
           WHEN file_pushgia1.[dvt] = 'V' THEN 
               CAST(file_pushgia1.[Vien/Hop] AS DECIMAL(18, 2)) * 
               CAST(file_pushgia1.[Đơn giá niêm yết_chuaquydoi] AS DECIMAL(18, 2))
           ELSE 
               CAST(file_pushgia1.[Đơn giá niêm yết_chuaquydoi] AS DECIMAL(18, 2))
       END AS [Đơn giá sau chiết khấu],

	  CASE 
			WHEN  LEFT([file_pushgia1].[Mahang], 2) IN ('31', '32', '33', '34') THEN 0
			ELSE 
				CASE
					WHEN Loai IN ('GBLL1KGT', 'Giachinhsachkgt', 'Giasicont') THEN 0
					WHEN Loai IN ('GBLL1KGC', 'Giachinhsachkgc') THEN 
						CASE 
							WHEN dvt = 'H' THEN Tuden
							ELSE Tuden / (TRY_CAST(file_pushgia1.[M2/Vien] AS DECIMAL(18,2))*TRY_CAST(file_pushgia1.[Vien/Hop] AS DECIMAL(18,2)))
						END
					WHEN Loai = 'Giaomkho200' THEN 200
					ELSE 0 
				END 
	  END AS [TU],
	   CASE 
			WHEN  LEFT([file_pushgia1].[Mahang], 2) IN ('31', '32', '33', '34') THEN 10000
			ELSE 
				CASE
					WHEN Loai IN ('GBLL1KGT', 'Giachinhsachkgt') THEN
						CASE 
							WHEN dvt = 'H' THEN Tuden
							ELSE Tuden / (TRY_CAST(file_pushgia1.[M2/Vien] AS DECIMAL(18,2))*TRY_CAST(file_pushgia1.[Vien/Hop] AS DECIMAL(18,2)))
						END
					WHEN Loai IN ('GBLL1KGC', 'Giachinhsachkgc', 'Giaomkho200', 'Giasicont') THEN 10000
					ELSE 0 
				END 
	  END AS [DEN],
	  [Nhom Gia]
	   ,[Biên độ giảm]
	   ,[Biên độ tăng]
      ,[Kích thước]
      ,[Khung Giá]
      ,[Mã SP]
      ,[Loại]
      ,[Gạch Ốp/Lát]
      ,[ĐVT]
      ,[TỪ]
      ,[ĐẾN]
      ,[GIÁ BÁN]
      ,[GIÁ CHÍNH SÁCH]
      ,[GIÁ SÀN]
INTO file_pushgia2
FROM [UNIS_DATA].[dbo].[file_pushgia1]
----- Tạo bảng file_pushgia3-----
DROP TABLE IF EXISTS file_pushgia3;
-- Tạo bảng phân loại và ký tự
;WITH Classification AS (
    SELECT t.*,
           CASE
               WHEN t.Loai IN ('GBLL1KGC', 'GBLL1KGT') THEN 1
               WHEN t.Loai IN ('Giachinhsachkgc', 'Giachinhsachkgt') THEN 1
               WHEN t.Loai = 'Giaomkho200' THEN 4
			   WHEN t.Loai = 'Giacatlo' THEN 3
               WHEN t.Loai = 'Giasicont' THEN 2
               ELSE NULL -- Hoặc giá trị mặc định nếu cần
           END AS KyTu
    FROM file_pushgia2 t
)
, Ranked AS (
    SELECT *,
           ROW_NUMBER() OVER (
               PARTITION BY Macn, [Mặt hàng], KyTu
               ORDER BY 
                   CASE
                       -- Ưu tiên cặp 'Giachinhsachkgc', 'Giachinhsachkgt' nếu KyTu = 1
                       WHEN KyTu = 1 AND Loai IN ('Giachinhsachkgc', 'Giachinhsachkgt') THEN 1
                       -- Ưu tiên cặp 'GBLL1KGC', 'GBLL1KGT' nếu KyTu = 1
                       WHEN KyTu = 1 AND Loai IN ('GBLL1KGC', 'GBLL1KGT') THEN 2
                       -- Các ký tự khác hoặc các loại khác
                       ELSE 3
                   END
           ) AS RowNum
    FROM Classification
)
-- Truy vấn chính để lấy dữ liệu dựa trên các điều kiện ưu tiên và lưu vào bảng mới
SELECT * 
INTO file_pushgia3  -- Tên bảng mới để lưu kết quả
FROM Ranked
WHERE (KyTu = 1 AND RowNum <= 2) -- Lấy cả hai loại ưu tiên cho KyTu = 1
   OR (KyTu <> 1)              -- Lấy tất cả các ký tự khác
ORDER BY [Mặt hàng]; -- Sắp xếp kết quả
--------------------Tạo bảng file_pdf_gia_sppt---------------------------------------
DROP TABLE IF EXISTS file_pdf_gia_sppt;
SELECT DISTINCT 
    b1.[Cong Ty Trien Khai],b1.[Ten Chi Nhanh],ps.[Macn], ps.[ĐVT] AS dvt,b.[Quydoidonvi],b.[Vien/Hop],b.[M2/Vien], ps.[Mặt hàng], 
	 CASE 
           WHEN ps.[Loai] IN ('GBL', 'Giacatlo') THEN 1
           ELSE NULL -- hoặc giá trị mặc định khác nếu cần
     END AS [Vùng bán hàng]
	,
		CASE 
			WHEN ps.[Loai] = 'GBL' THEN 1
			WHEN ps.[Loai] = 'Giacatlo' THEN 3
			ELSE NULL
		END AS [LOAI BAN HANG],
	    1 AS HTTT,
	    '' AS [Tên hàng hóa, dịch vụ],
	    ps.[ĐVT] AS [ĐVT_1],
	    ps.[Loai],
	    ps.[Đơn giá niêm yết],
	    CAST(NULL AS FLOAT) AS [% Giảm giá],
        CAST(NULL AS DECIMAL(37, 4)) AS [Tiền Giảm giá],
	    ps.[Đơn giá sau chiết khấu],
	    0 AS TU,
	    CAST(10000 AS DECIMAL(37, 4)) AS DEN,
	     b.[Nhom Gia],
		'0' AS [Biên độ giảm],
		'0' AS [Biên độ tăng],
		'' AS [Kích thước],
		'' AS [Khung Giá],
		'' AS [Mã SP],
		'' AS [Loại],
		'' AS [Gạch Ốp/Lát],
		'' AS [ĐVT],
		'' AS [TỪ],
		'' AS [ĐẾN],
		'' AS [GIÁ BÁN],
		'' AS [GIÁ CHÍNH SÁCH],
		'' AS [GIÁ SÀN]

INTO file_pdf_gia_sppt
FROM [Phantichgia_sppt] ps
LEFT JOIN [BangDoMaHang] b ON ps.[Mặt hàng] = b.[Ma Phan Mem]
LEFT JOIN [BangDoChiNhanh] b1 ON ps.[Macn] = b1.[Ma Chi Nhanh];
--------------------Add Bảng phân tích SPPT vào file Push---------------------------------------
USE UNIS_DATA
INSERT INTO [file_pushgia3] ([Cong Ty Trien Khai],[Ten Chi Nhanh],[Macn],[dvt],[Quydoidonvi],[Vien/Hop],[M2/Vien]
      ,[Mặt hàng],[Vùng bán hàng],[LOAI BAN HANG],[HTTT],[Tên hàng hóa, dịch vụ],[ĐVT_1],[Loai],[Đơn giá niêm yết]
      ,[% Giảm giá],[Tiền Giảm giá],[Đơn giá sau chiết khấu],[TU],[DEN],[Nhom Gia],[Biên độ giảm],[Biên độ tăng]
      ,[Kích thước],[Khung Giá],[Mã SP],[Loại],[Gạch Ốp/Lát],[ĐVT],[TỪ],[ĐẾN],[GIÁ BÁN],[GIÁ CHÍNH SÁCH],[GIÁ SÀN])

SELECT [Cong Ty Trien Khai],[Ten Chi Nhanh],[Macn],[dvt],[Quydoidonvi],[Vien/Hop],[M2/Vien]
      ,[Mặt hàng],[Vùng bán hàng],[LOAI BAN HANG],[HTTT],[Tên hàng hóa, dịch vụ],[ĐVT_1],[Loai],[Đơn giá niêm yết]
      ,[% Giảm giá],[Tiền Giảm giá],[Đơn giá sau chiết khấu],[TU],[DEN],[Nhom Gia],[Biên độ giảm],[Biên độ tăng]
      ,[Kích thước],[Khung Giá],[Mã SP],[Loại],[Gạch Ốp/Lát],[ĐVT],[TỪ],[ĐẾN],[GIÁ BÁN],[GIÁ CHÍNH SÁCH],[GIÁ SÀN]
FROM [file_pdf_gia_sppt];

"""

# Thực hiện cập nhật
with conn.cursor() as cursor:
    cursor.execute(query)
    conn.commit()  # Lưu các thay đổi

# Đóng kết nối
conn.close()