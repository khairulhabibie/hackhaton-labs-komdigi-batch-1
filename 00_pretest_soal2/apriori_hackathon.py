import pandas as pd
from mlxtend.frequent_patterns import apriori, association_rules

def run_analysis(input_xlsx_path: str, output_xlsx_path: str) -> None:
    print("Membaca file transaksi...")
    # 1. Baca data dari Excel
    df = pd.read_excel(input_xlsx_path, sheet_name="Transaksi")
    df = df[['Kode Transaksi', 'Nama Produk']]

    # 2. Buat data pivot transaksi vs produk
    print("Menyiapkan data transaksi...")
    basket = (df.groupby(['Kode Transaksi', 'Nama Produk'])['Nama Produk']
                .count().unstack().fillna(0))
    basket = (basket > 0).astype(bool)  # Konversi ke tipe bool agar sesuai rekomendasi mlxtend

    # 3. Jalankan Apriori
    print("Menjalankan algoritma Apriori...")
    frequent_itemsets = apriori(basket, min_support=0.05, use_colnames=True)
    rules = association_rules(frequent_itemsets, metric="confidence", min_threshold=0.4)

    if rules.empty:
        print("Tidak ada kombinasi produk yang memenuhi syarat support dan confidence.")
        return

    # 4. Buat kolom 'Products' berisi kombinasi produk yang disortir
    print("Menggabungkan kombinasi produk...")
    rules['Products'] = rules.apply(
        lambda r: ";".join(sorted(set(list(r['antecedents']) + list(r['consequents'])))),
        axis=1
    )

    # 5. Gabungkan kombinasi produk yang sama (ambil max lift & confidence)
    agg_rules = (
        rules.groupby('Products')
        .agg(Maximum_Lift=('lift', 'max'),
             Maximum_Confidence=('confidence', 'max'))
        .reset_index()
    )

    # 6. Urutkan hasil berdasarkan lift dan confidence (descending)
    agg_rules = agg_rules.sort_values(
        by=['Maximum_Lift', 'Maximum_Confidence'],
        ascending=[False, False]
    ).reset_index(drop=True)

    # 7. Tambahkan kolom Packaging Set ID
    agg_rules.insert(0, 'Packaging Set ID', range(1, len(agg_rules) + 1))

    # 8. Simpan hasil ke Excel
    print("Menyimpan hasil ke file Excel...")
    agg_rules.to_excel(output_xlsx_path, index=False)
    print(f"Analisis selesai! File hasil disimpan sebagai: {output_xlsx_path}")

# Contoh pemanggilan langsung
if __name__ == "__main__":
    run_analysis("transaksi_dqmart.xlsx", "product_packaging.xlsx")
