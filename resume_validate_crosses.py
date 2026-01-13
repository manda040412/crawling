import pandas as pd

# file utama dan autosave terakhir
main_file = "List spare parts-Anugerah Auto.xlsx"   # file master semua item
autosave_file = "autosave_1100_115011.xlsx"         # file autosave terakhir kamu

print("📘 Membaca data utama & autosave...")

main_df = pd.read_excel(main_file)
auto_df = pd.read_excel(autosave_file)

# pastikan nama kolom sesuai
main_df['ItemCode'] = main_df['ItemCode'].astype(str).str.strip()
auto_df['Item Code'] = auto_df['Item Code'].astype(str).str.strip()

# cari item yang belum pernah diproses
done_codes = auto_df['Item Code'].unique().tolist()
remaining_df = main_df[~main_df['ItemCode'].isin(done_codes)]

print(f"✅ Total data di autosave  : {len(done_codes)} item")
print(f"🔁 Sisa yang belum diproses: {len(remaining_df)} item\n")

output = "resume_list.xlsx"
remaining_df.to_excel(output, index=False)

print(f"📄 Daftar sisa item disimpan ke: {output}")
print("\n👉 Sekarang ubah EXCEL_PATH di script utama jadi 'resume_list.xlsx'")
