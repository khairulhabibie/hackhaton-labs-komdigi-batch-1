import pandas as pd

def normalize_tanggal_transaksi(input_xlsx_path: str, output_xlsx_path: str) -> None:
    try:
        df = pd.read_excel(input_xlsx_path, sheet_name='transaksi')
    except FileNotFoundError:
        raise FileNotFoundError(f"File tidak ditemukan: {input_xlsx_path}")
    except ValueError:
        raise ValueError("Sheet 'transaksi' tidak ditemukan di file input.")

    # cari kolom tanggal
    tanggal_cols = [c for c in df.columns if 'tanggal' in c.lower()]
    if not tanggal_cols:
        raise KeyError("Kolom tanggal tidak ditemukan. Pastikan nama kolom mengandung 'tanggal'.")
    tanggal_col = tanggal_cols[0]

    # mapping bulan (Indonesia + Inggris + singkatan campur)
    month_map = {
        'jan': '01', 'januari': '01', 'january': '01',
        'feb': '02', 'februari': '02', 'february': '02',
        'mar': '03', 'maret': '03', 'march': '03',
        'apr': '04', 'april': '04',
        'mei': '05', 'may': '05',
        'jun': '06', 'juni': '06', 'june': '06',
        'jul': '07', 'juli': '07', 'july': '07',
        'agu': '08', 'agus': '08', 'agustus': '08', 'aug': '08', 'august': '08',
        'sep': '09', 'sept': '09', 'september': '09',
        'okt': '10', 'oct': '10', 'oktober': '10', 'october': '10',
        'nov': '11', 'november': '11',
        'des': '12', 'dec': '12', 'desember': '12', 'december': '12'
    }

    def clean_date(x):
        if pd.isna(x):
            return x

        s = str(x).strip().lower()
        # bersihkan karakter pengganggu
        s = s.replace(',', ' ').replace('.', ' ').replace('/', ' ').replace('-', ' ')
        s = s.replace("‘", ' ').replace("’", ' ').replace("'", ' ')
        s = ' '.join(s.split())  # hapus spasi ganda

        tokens = s.split()
        day, month, year = None, None, None

        # cari tahun
        for t in tokens:
            if t.isdigit():
                val = int(t)
                if val > 1900:  # tahun 4 digit
                    year = str(val)
                elif val < 100:  # tahun 2 digit (misal '24)
                    year = f"20{val:02d}"

        # cari bulan
        for t in tokens:
            key = t[:3]
            if key in month_map:
                month = month_map[key]
                break

        # cari hari
        for t in tokens:
            if t.isdigit() and 1 <= int(t) <= 31:
                if year is None or str(int(t)) not in year:
                    day = t
                    break

        # deteksi format seperti "2024 6 nov" (tahun dulu)
        if tokens and tokens[0].isdigit() and len(tokens) >= 3:
            if int(tokens[0]) > 1900 and tokens[2][:3] in month_map:
                year = tokens[0]
                day = tokens[1]
                month = month_map[tokens[2][:3]]

        # fallback (tanggal tunggal)
        if not day and tokens and tokens[0].isdigit():
            day = tokens[0]

        # pastikan semua field terisi
        try:
            day = f"{int(day):02d}" if day else "01"
            month = month if month else "01"
            year = year if year else "1900"
            return f"{day}-{month}-{year}"
        except:
            return x

    # transformasi
    df[tanggal_col] = df[tanggal_col].apply(clean_date)

    # simpan dengan sheet yang sama
    with pd.ExcelWriter(output_xlsx_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='transaksi')


if __name__ == "__main__":
    input_file = "penjualan_dqmart_01.xlsx"
    output_file = "penjualan_dqmart_01_output.xlsx"
    normalize_tanggal_transaksi(input_file, output_file)
