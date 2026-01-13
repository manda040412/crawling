import pandas as pd
from datetime import datetime

# ================================
# 📘 1. BACA FILE
# ================================
cross_file = "Jikiu_Crosses_Merged_20260113_165618.xlsx"
validation_file = "validation_FULL_20260113_133316.xlsx"

print("📂 Membaca file...")
df_cross = pd.read_excel(cross_file)
df_val = pd.read_excel(validation_file)

print(f"✅ Cross data: {len(df_cross)} baris, Validation: {len(df_val)} baris")

# ================================
# 🧹 2. NORMALISASI KOLOM
# ================================
# Ubah nama kolom ke huruf kecil semua
df_cross.columns = df_cross.columns.str.strip().str.lower()
df_val.columns = df_val.columns.str.strip().str.lower()

# Deteksi nama kolom 'item code'
cross_item_col = next((c for c in df_cross.columns if "item" in c and "code" in c), None)
val_item_col = next((c for c in df_val.columns if "item" in c and "code" in c), None)

if not cross_item_col or not val_item_col:
    raise ValueError("❌ Tidak menemukan kolom Item Code di salah satu file!")

# ================================
# 🔗 3. AMBIL KOLOM STATUS DAN DETAILS
# ================================
needed_cols = ["status", "details", val_item_col]
df_val_sub = df_val[[c for c in needed_cols if c in df_val.columns]].copy()

# ================================
# 🔗 4. GABUNGKAN TANPA UBAH URUTAN
# ================================
merged = pd.merge(
    df_cross,
    df_val_sub,
    how="left",
    left_on=cross_item_col,
    right_on=val_item_col
)

# Hapus kolom duplikat item code dari validation
if val_item_col in merged.columns and val_item_col != cross_item_col:
    merged.drop(columns=[val_item_col], inplace=True)

# ================================
# 📦 5. SIMPAN HASIL FINAL
# ================================
output_file = f"Jikiu_Crosses_Merged_Status_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
merged.to_excel(output_file, index=False)

print("\n✅ Proses selesai tanpa mengubah urutan data!")
print(f"💾 File disimpan sebagai: {output_file}")
print(f"📊 Total baris akhir: {len(merged)}")
