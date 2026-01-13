import pandas as pd
import glob
from datetime import datetime

print("📦 Menggabungkan semua hasil autosave...")

# cari semua file autosave
files = sorted(glob.glob("autosave_*.xlsx"))

if not files:
    print("⚠️ Tidak ada file autosave ditemukan di folder ini.")
    exit()

all_data = []

for f in files:
    try:
        df = pd.read_excel(f)
        print(f"   📄 {f}  ({len(df)} baris)")
        all_data.append(df)
    except Exception as e:
        print(f"   ⚠️ Gagal baca {f}: {e}")

# gabungkan semua hasil
merged = pd.concat(all_data, ignore_index=True)

# hapus duplikat (berdasarkan kombinasi Item Code + Owner + Number)
if all(col in merged.columns for col in ["Item Code", "Owner", "Number"]):
    merged.drop_duplicates(subset=["Item Code", "Owner", "Number"], inplace=True)

# hapus baris kosong total
merged.dropna(how='all', inplace=True)

# reset index
merged.reset_index(drop=True, inplace=True)

# simpan hasil final
output = f"Jikiu_Crosses_Merged_Clean_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
merged.to_excel(output, index=False)

print(f"\n✅ Semua autosave berhasil digabung!")
print(f"📁 File akhir disimpan sebagai: {output}")
print(f"📊 Total baris unik: {len(merged)}")
