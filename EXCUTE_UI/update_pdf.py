import pyodbc

# Thông tin kết nối
server = 'XPS7590'
database = 'UNIS_DATA'

# Tạo kết nối với Windows Authentication
conn = pyodbc.connect(f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes;')

# Thực hiện truy vấn
query = """
USE UNIS_DATA
DROP TABLE IF EXISTS file_pdf_gia;
SELECT DISTINCT 
		p.[Macn], p.[Mahang], p.[dvt], p.[GIA], p.[Loai],
       b.[Dac Diem], b.[Quydoidonvi],b.[Vien/Hop],b.[M2/Vien],
    SUBSTRING(b.[Ma Sale], CHARINDEX('.',b.[Ma Sale]) + 1, LEN(b.[Ma Sale])) AS [Ma_1],
	b.[BGTD_CaoCap/PhoThong],b.[BGTD_Op/Lat],b.BGTD_QuiCachBangGia AS Kichthuoc,b.[Bo Mau],b.[BGTD_BoMau] AS Dacdiemchia,p.Khunggiaban,
	CASE 
           WHEN  p.[dvt] = 'M' THEN 
               p.[GIA]
           WHEN  p.[dvt] = 'H' THEN 
                CAST(p.[GIA] AS DECIMAL(18, 2))/ 
               CAST(b.[Quydoidonvi] AS DECIMAL(18, 2))
           ELSE 
               CAST(p.[GIA] AS DECIMAL(18, 2))/ 
               CAST(b.[M2/Vien] AS DECIMAL(18, 2))
    END AS [GIA_SOSANH]
INTO file_pdf_gia
FROM [UNIS_DATA].[dbo].[Phantichgia_state1] p
LEFT JOIN [UNIS_DATA].[dbo].[BangDoMaHang] b ON p.[Mahang] = b.[Ma Sale];

-----Chia khung gia---------------------------------
-- Tính giá trung bình cho từng nhóm, với điều kiện xử lý riêng biệt cho 'Bộ Đậm Nhạt' và 'Bộ Thân Viền'
DROP TABLE IF EXISTS Banle_pdf;
-- Tính giá trung bình cho từng nhóm, với điều kiện xử lý riêng biệt cho 'Bộ Đậm Nhạt' và 'Bộ Thân Viền'
;WITH AvgPrice AS (
    SELECT DISTINCT
        Macn,
        Kichthuoc,
        Dacdiemchia,
        [Bo Mau],
        AVG(GIA) AS AvgGIA
    FROM
        file_pdf_gia
    WHERE
        Dacdiemchia IN ('Bộ Đậm Nhạt', 'Bộ Thân Viền', 'Bộ Đậm Nhạt Matt', 'Bộ Siêu Bóng', 'Bộ Thân Viền Matt') AND Loai IN ('GBLL1KGT','GBLL1KGC')
    GROUP BY
        Macn,
        Kichthuoc,
        Dacdiemchia,
        [Bo Mau]

    UNION ALL

    SELECT DISTINCT
        Macn,
        Kichthuoc,
        'Khác' AS Dacdiemchia,
        [Bo Mau],
        AVG(GIA_SOSANH) AS AvgGIA
    FROM
        file_pdf_gia
    WHERE
        Dacdiemchia NOT IN ('Bộ Đậm Nhạt', 'Bộ Thân Viền', 'Bộ Đậm Nhạt Matt', 'Bộ Siêu Bóng', 'Bộ Thân Viền Matt') AND Loai IN ('GBLL1KGT','GBLL1KGC')
    GROUP BY
        Macn,
        Kichthuoc,
        [Bo Mau]
),

RankedPrices AS (
    SELECT DISTINCT
        Macn,
        Kichthuoc,
        Dacdiemchia,
        [Bo Mau],
        AvgGIA,
        DENSE_RANK() OVER (
            PARTITION BY Macn, Kichthuoc, Dacdiemchia
            ORDER BY AvgGIA ASC
        ) AS Rank
    FROM
        AvgPrice
),

ClassifiedPrices AS (
    SELECT DISTINCT
        t.Macn,
        t.Kichthuoc,
        t.Loai,
        t.Mahang,
		t.[Dac Diem],
        t.Khunggiaban,
        t.[BGTD_CaoCap/PhoThong],
        t.[BGTD_Op/Lat],
        t.dvt,
        t.Quydoidonvi,
        t.[Vien/Hop],
        t.Ma_1,
        t.Dacdiemchia,
        t.[Bo Mau],
        t.GIA,
        r.Rank,
        CASE
            WHEN r.Rank IS NULL THEN 'CCG'
            WHEN r.Dacdiemchia IN ('Bộ Đậm Nhạt', 'Bộ Thân Viền', 'Bộ Đậm Nhạt Matt', 'Bộ Siêu Bóng', 'Bộ Thân Viền Matt')  THEN
                CASE
                    WHEN r.Rank = 1 THEN 'KG1'
                    WHEN r.Rank = 2 THEN 'KG2'
                    WHEN r.Rank = 3 THEN 'KG3' 
					WHEN r.Rank = 4 THEN 'KG4'
                    ELSE 'KG5'
                END
            ELSE
                 CASE
                    WHEN r.Rank = 1 THEN 'KG1'
                    WHEN r.Rank = 2 THEN 'KG2'
                    WHEN r.Rank = 3 THEN 'KG3' 
					WHEN r.Rank = 4 THEN 'KG4'
                    ELSE 'KG5'
                END
        END AS KHUNG_GIA_1,
        ROW_NUMBER() OVER (PARTITION BY t.Macn, t.Kichthuoc ORDER BY r.Rank) AS RowNum
    FROM
        file_pdf_gia t
    
    LEFT JOIN RankedPrices r ON 
        t.Macn = r.Macn
        AND t.Kichthuoc = r.Kichthuoc
        AND (
            (t.Dacdiemchia = r.Dacdiemchia AND r.Dacdiemchia IN ('Bộ Đậm Nhạt', 'Bộ Thân Viền', 'Bộ Đậm Nhạt Matt', 'Bộ Siêu Bóng', 'Bộ Thân Viền Matt'))
            OR (r.Dacdiemchia = 'Khác')
        )
        AND t.[Bo Mau] = r.[Bo Mau]
	WHERE
        Loai IN ('GBLL1KGT','GBLL1KGC')
)

-- Hiển thị kết quả phân loại khung giá
SELECT DISTINCT
        ClassifiedPrices.Macn,
        Kichthuoc,
        Mahang,
		[Dac Diem],
		[BGTD_CaoCap/PhoThong],
		[BGTD_Op/Lat],
		dvt,
		Quydoidonvi,
		[Vien/Hop],
		Ma_1,
		Dacdiemchia,
        [Bo Mau],
		Loai,
        GIA,
		Khunggiaban,
		KHUNG_GIA_1,
		COALESCE(b2.[Tuden], 0) AS Tuden
INTO Banle_pdf
FROM
    ClassifiedPrices
LEFT JOIN 
    [UNIS_DATA].[dbo].[bangdokhungtrenduoi] b2 ON CONCAT(b2.[Macn], b2.[BGTD_QuiCachBangGia]) = CONCAT(ClassifiedPrices.[Macn],ClassifiedPrices.[Kichthuoc])
ORDER BY
	
   Kichthuoc,
	[Bo Mau],
	GIA,
    ClassifiedPrices.Macn,
    Dacdiemchia

;

; 

DROP TABLE IF EXISTS Chinhsach_pdf;
-- Tính giá trung bình cho từng nhóm, với điều kiện xử lý riêng biệt cho 'Bộ Đậm Nhạt' và 'Bộ Thân Viền'
;WITH AvgPrice_chinhsach AS (
    SELECT DISTINCT
        Macn,
        Kichthuoc,
        Dacdiemchia,
        [Bo Mau],
        AVG(GIA) AS AvgGIA
    FROM
        file_pdf_gia
    WHERE
        Dacdiemchia IN ('Bộ Đậm Nhạt', 'Bộ Thân Viền', 'Bộ Đậm Nhạt Matt', 'Bộ Siêu Bóng', 'Bộ Thân Viền Matt') AND Loai IN ('Giachinhsachkgc','Giachinhsachkgt')
    GROUP BY
        Macn,
        Kichthuoc,
        Dacdiemchia,
        [Bo Mau]

    UNION ALL

    SELECT DISTINCT
        Macn,
        Kichthuoc,
        'Khác' AS Dacdiemchia,
        [Bo Mau],
        AVG(GIA_SOSANH) AS AvgGIA
    FROM
        file_pdf_gia
    WHERE
        Dacdiemchia NOT IN ('Bộ Đậm Nhạt', 'Bộ Thân Viền', 'Bộ Đậm Nhạt Matt', 'Bộ Siêu Bóng', 'Bộ Thân Viền Matt') AND Loai IN ('Giachinhsachkgc','Giachinhsachkgt')
    GROUP BY
        Macn,
        Kichthuoc,
        [Bo Mau]
),

RankedPrices_chinhsach AS (
    SELECT DISTINCT
        Macn,
        Kichthuoc,
        Dacdiemchia,
        [Bo Mau],
        AvgGIA,
        DENSE_RANK() OVER (
            PARTITION BY Macn, Kichthuoc, Dacdiemchia
            ORDER BY AvgGIA ASC
        ) AS Rank
    FROM
        AvgPrice_chinhsach
),

ClassifiedPrices_chinhsach AS (
    SELECT DISTINCT
        t.Macn,
        t.Kichthuoc,
        t.Loai,
        t.Mahang,
		t.[Dac Diem],
        t.Khunggiaban,
        t.[BGTD_CaoCap/PhoThong],
        t.[BGTD_Op/Lat],
        t.dvt,
        t.Quydoidonvi,
        t.[Vien/Hop],
        t.Ma_1,
        t.Dacdiemchia,
        t.[Bo Mau],
        t.GIA,
        r.Rank,
        CASE
            WHEN r.Rank IS NULL THEN 'CCG'
            WHEN r.Dacdiemchia IN ('Bộ Đậm Nhạt', 'Bộ Thân Viền', 'Bộ Đậm Nhạt Matt', 'Bộ Siêu Bóng', 'Bộ Thân Viền Matt')  THEN
                 CASE
                    WHEN r.Rank = 1 THEN 'KG1'
                    WHEN r.Rank = 2 THEN 'KG2'
                    WHEN r.Rank = 3 THEN 'KG3' 
					WHEN r.Rank = 4 THEN 'KG4'
                    ELSE 'KG5'
                END
            ELSE
                 CASE
                    WHEN r.Rank = 1 THEN 'KG1'
                    WHEN r.Rank = 2 THEN 'KG2'
                    WHEN r.Rank = 3 THEN 'KG3' 
					WHEN r.Rank = 4 THEN 'KG4'
                    ELSE 'KG5'
                END
        END AS KHUNG_GIA_1,
        ROW_NUMBER() OVER (PARTITION BY t.Macn, t.Kichthuoc ORDER BY r.Rank) AS RowNum
    FROM
        file_pdf_gia t
    
    LEFT JOIN RankedPrices_chinhsach r ON 
        t.Macn = r.Macn
        AND t.Kichthuoc = r.Kichthuoc
        AND (
            (t.Dacdiemchia = r.Dacdiemchia AND r.Dacdiemchia IN ('Bộ Đậm Nhạt', 'Bộ Thân Viền', 'Bộ Đậm Nhạt Matt', 'Bộ Siêu Bóng', 'Bộ Thân Viền Matt'))
            OR (r.Dacdiemchia = 'Khác')
        )
        AND t.[Bo Mau] = r.[Bo Mau]
	WHERE
        Loai IN ('Giachinhsachkgc','Giachinhsachkgt')
)

-- Hiển thị kết quả phân loại khung giá
SELECT DISTINCT
		ClassifiedPrices_chinhsach.Macn,
        Kichthuoc,
        Mahang,
		[Dac Diem],
		[BGTD_CaoCap/PhoThong],
		[BGTD_Op/Lat],
		dvt,
		Quydoidonvi,
		[Vien/Hop],
		Ma_1,
		Dacdiemchia,
        [Bo Mau],
		Loai,
        GIA,
		Khunggiaban,
		KHUNG_GIA_1,
		COALESCE(b2.[Tuden], 0) AS Tuden
INTO Chinhsach_pdf
FROM
    ClassifiedPrices_chinhsach 
LEFT JOIN 
    [UNIS_DATA].[dbo].[bangdokhungtrenduoi] b2 ON CONCAT(b2.[Macn], b2.[BGTD_QuiCachBangGia]) = CONCAT(ClassifiedPrices_chinhsach.[Macn],ClassifiedPrices_chinhsach.[Kichthuoc])
ORDER BY
	
   Kichthuoc,
	[Bo Mau],
	GIA,
    ClassifiedPrices_chinhsach.Macn,
    Dacdiemchia

;

; 
"""

# Thực hiện cập nhật
with conn.cursor() as cursor:
    cursor.execute(query)
    conn.commit()  # Lưu các thay đổi

# Đóng kết nối
conn.close()