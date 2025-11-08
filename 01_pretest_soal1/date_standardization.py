import pandas as pd
import os

def normalize_tanggal_transaksi(input_xlsx_path: str, output_xlsx_path: str) -> None:
    # cek file
    if not os.path.exists(input_xlsx_path):
        raise FileNotFoundError(f"File tidak ditemukan: {input_xlsx_path}")

    # baca excel
    df = pd.read_excel(input_xlsx_path, sheet_name='transaksi')

    # cek kolom tanggal
    tanggal_cols = [c for c in df.columns if 'tanggal' in c.lower()]
    if not tanggal_cols:
        raise KeyError("Kolom tanggal tidak ditemukan. Pastikan ada kolom seperti 'Tanggal Transaksi'.")
    tanggal_col = tanggal_cols[0]

    # Mapping
    month_map = {
        'jan': '01', 'januari': '01', 'january': '01',
        'feb': '02', 'februari': '02', 'february': '02',
        'mar': '03', 'maret': '03', 'march': '03',
        'apr': '04', 'april': '04',
        'mei': '05', 'may': '05',
        'jun': '06', 'juni': '06', 'june': '06',
        'jul': '07', 'juli': '07', 'july': '07',
        'agu': '08', 'agustus': '08', 'aug': '08', 'august': '08',
        'sep': '09', 'sept': '09', 'september': '09',
        'okt': '10', 'oct': '10', 'oktober': '10', 'october': '10',
        'nov': '11', 'november': '11',
        'des': '12', 'dec': '12', 'desember': '12', 'december': '12'
    }

    def clean_date(x):
        if pd.isna(x):
            return x
        s = str(x).strip().lower()
        s = s.replace(',', ' ').replace('.', ' ').replace("'", ' ')
        s = s.replace("‘", ' ').replace("’", ' ').replace('/', '-')
        s = ' '.join(s.split())

        if '-' in s and len(s.split('-')) == 3:
            parts = s.split('-')
            try:
                if len(parts[0]) == 4:
                    y, m, d = parts
                elif len(parts[2]) == 4:
                    d, m, y = parts
                else:
                    return s
                return f"{int(d):02d}-{int(m):02d}-{int(y):04d}"
            except:
                pass

        tokens = s.split()
        day, month, year = None, None, None

        for t in tokens:
            if t.isdigit() and (len(t) == 4 or len(t) == 2):
                val = int(t)
                year = f"20{val:02d}" if val < 100 else str(val)

        for t in tokens:
            key = t[:3]
            if key in month_map:
                month = month_map[key]
                break

        for t in tokens:
            if t.isdigit() and int(t) <= 31:
                if year is None or t not in year:
                    day = t
                    break

        if tokens and tokens[0].isdigit() and len(tokens[0]) >= 2 and not tokens[0][:3] in month_map:
            if len(tokens) >= 3 and tokens[1].isdigit() and tokens[2][:3] in month_map:
                year = tokens[0]
                day = tokens[1]
                month = month_map[tokens[2][:3]]

        try:
            day = f"{int(day):02d}" if day else "01"
            month = month if month and month.isdigit() else "01"
            year = year if year else "1900"
            return f"{day}-{month}-{year}"
        except:
            return x

    print(f"🔍 Kolom tanggal terdeteksi: '{tanggal_col}'")
    df[tanggal_col] = df[tanggal_col].apply(clean_date)
    df.to_excel(output_xlsx_path, index=False)
    print(f"✅ File output berhasil dibuat: {output_xlsx_path}")


# ============================================================
# Menjalankan file
# ============================================================
if __name__ == "__main__":
    input_file = "penjualan_dqmart_01-beta.xlsx"
    output_file = "penjualan_dqmart_01-beta_output.xlsx"

    normalize_tanggal_transaksi(input_file, output_file)
