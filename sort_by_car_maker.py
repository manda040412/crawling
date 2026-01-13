import pandas as pd
import os
from datetime import datetime

# ================================
# ⚙️ 1. CARI FILE TERBARU OTOMATIS
# ================================
prefix = "Jikiu_Crosses_Merged_Status_"
latest_file = None
latest_time = 0

for f in os.listdir("."):
    if f.startswith(prefix) and f.endswith(".xlsx"):
        mtime = os.path.getmtime(f)
        if mtime > latest_time:
            latest_file = f
            latest_time = mtime

if not latest_file:
    raise FileNotFoundError(f"❌ Tidak ada file dengan prefix '{prefix}' di folder ini.")

print(f"📂 File terbaru ditemukan: {latest_file}")

# ================================
# 📘 2. BACA FILE
# ================================
df = pd.read_excel(latest_file)
print(f"📊 Total baris awal: {len(df)}")

# ================================
# 🧹 3. NORMALISASI KOLOM
# ================================
df.columns = df.columns.str.strip().str.lower()

# Deteksi kolom 'Car Maker Name'
car_maker_col = next((c for c in df.columns if "car maker" in c), None)
if not car_maker_col:
    raise ValueError("❌ Tidak ditemukan kolom 'Car Maker Name' di file ini!")

print(f"🔍 Kolom pengurutan: {car_maker_col}")

# ================================
# 🔢 4. SORT DATA BERDASARKAN CAR MAKER
# ================================
# Urut berdasarkan Car Maker Name (A–Z)
# Baris yang kosong (NaN) akan diletakkan di bawah
df_sorted = df.sort_values(by=car_maker_col, ascending=True, na_position='last')

# ================================
# 💾 5. SIMPAN HASIL
# ================================
output_file = f"Jikiu_Crosses_SortedByCarMaker_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
df_sorted.to_excel(output_file, index=False)

print(f"\n✅ Selesai! Data telah diurutkan berdasarkan '{car_maker_col}'.")
print(f"💾 File disimpan sebagai: {output_file}")
print(f"📊 Total baris akhir: {len(df_sorted)}")
