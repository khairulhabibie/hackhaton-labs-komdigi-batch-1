import pandas as pd
from mlxtend.frequent_patterns import apriori, association_rules

def run_analysis(input_xlsx_path: str, output_xlsx_path: str) -> None:
    print("Membaca file transaksi...")

    # 1. Baca data
    df = pd.read_excel(input_xlsx_path, sheet_name="Transaksi", dtype=str)
    df = df[['Kode Transaksi', 'Nama Produk']].dropna()
    df['Nama Produk'] = df['Nama Produk'].astype(str).str.strip()

    # 2. Bentuk basket
    print("Menyiapkan data transaksi...")
    basket = (
        df.groupby(['Kode Transaksi', 'Nama Produk'])['Nama Produk']
        .count().unstack().fillna(0)
    )
    basket = (basket > 0).astype(bool)

    # 3. Jalankan Apriori
    print("Menjalankan Apriori...")
    frequent_itemsets = apriori(basket, min_support=0.05, use_colnames=True)
    rules = association_rules(frequent_itemsets, metric="confidence", min_threshold=0.4)

    if rules.empty:
        print("Tidak ada kombinasi produk yang memenuhi syarat.")
        return

    # 4. Gabungkan dan urutkan produk (A-Z)
    print("Menggabungkan kombinasi produk...")
    def combine_and_sort_products(antecedents, consequents):
        combined = list(set(list(antecedents) + list(consequents)))
        combined_sorted = sorted(combined)
        return ";".join(combined_sorted)

    rules['Products'] = rules.apply(
        lambda r: combine_and_sort_products(r['antecedents'], r['consequents']),
        axis=1
    )

    # 5. Agregasi kombinasi unik
    agg_rules = (
        rules.groupby('Products', as_index=False)
        .agg({
            'lift': 'max',
            'confidence': 'max'
        })
    )

    # 6. Ubah nama kolom sesuai format
    agg_rules.rename(columns={
        'lift': 'Maximum Lift',
        'confidence': 'Maximum Confidence'
    }, inplace=True)

    # 7. Urutkan hasil
    agg_rules.sort_values(
        by=['Maximum Lift', 'Maximum Confidence'],
        ascending=[False, False],
        inplace=True
    )
    agg_rules.reset_index(drop=True, inplace=True)

    # 8. Tambahkan ID
    agg_rules.insert(0, 'Packaging Set ID', range(1, len(agg_rules) + 1))

    # 9. Simpan hasil (tanpa index, sheet name tepat "Packaging")
    print("Menyimpan hasil ke Excel (sheet: Packaging)...")
    with pd.ExcelWriter(output_xlsx_path) as writer:
        agg_rules.to_excel(writer, index=False, sheet_name='Packaging')

    print(f"Analisis selesai! File disimpan sebagai: {output_xlsx_path}")


if __name__ == "__main__":
    run_analysis("transaksi_dqmart.xlsx", "product_packaging.xlsx")