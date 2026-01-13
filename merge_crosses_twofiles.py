import pandas as pd
from datetime import datetime

# ================================
# 📘 1. BACA FILE
# ================================
file1 = "Jikiu_Crosses_FinalPairs_FULL_20260113_121129.xlsx"
file2 = "250_20260113_154857.xlsx"

print("📂 Membaca file...")
df1 = pd.read_excel(file1)
df2 = pd.read_excel(file2)

print(f"✅ File1: {len(df1)} baris, File2: {len(df2)} baris")

# ================================
# 🧹 2. SAMAKAN STRUKTUR KOLOM
# ================================
df1.columns = df1.columns.str.strip()
df2.columns = df2.columns.str.strip()

# Normalisasi nama kolom jadi lowercase
df1.columns = df1.columns.str.lower()
df2.columns = df2.columns.str.lower()

# Pastikan kolom penting ada
essential_cols = ["item code", "owner", "number"]
for col in essential_cols:
    if col not in df1.columns or col not in df2.columns:
        raise ValueError(f"❌ Kolom '{col}' tidak ditemukan di salah satu file.")

# ================================
# 🔗 3. GABUNGKAN
# ================================
merged = pd.concat([df1, df2], ignore_index=True)
print(f"🔄 Total gabungan awal: {len(merged)} baris")

# ================================
# 🧹 4. HAPUS DUPLIKAT
# ================================
before = len(merged)
merged.drop_duplicates(subset=["item code", "owner", "number"], inplace=True)
after = len(merged)
print(f"🧹 Menghapus {before - after} duplikat")

# ================================
# 🗂️ 5. URUTKAN
# ================================
if "item code" in merged.columns:
    merged.sort_values(by="item code", inplace=True)

# ================================
# 💾 6. SIMPAN HASIL
# ================================
output_file = f"Jikiu_Crosses_Merged_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
merged.to_excel(output_file, index=False)

print(f"\n✅ Selesai! File disimpan sebagai: {output_file}")
print(f"📊 Total baris akhir: {len(merged)}")
