import pandas as pd
from datetime import datetime

# ================================
# 📘 1. BACA KEDUA FILE
# ================================
cross_file = "Jikiu_Crosses_FinalPairs_FULL_20260113_121129.xlsx"
validation_file = "validation_FULL_20260113_133316.xlsx"

print("📂 Membaca file...")
df_cross = pd.read_excel(cross_file)
df_val = pd.read_excel(validation_file)

print(f"✅ Crosses: {len(df_cross)} baris, Validation: {len(df_val)} baris")

# ================================
# 🧹 2. BERSIHKAN KOLOM
# ================================
# Normalisasi nama kolom jadi huruf kecil semua tanpa spasi
df_cross.columns = df_cross.columns.str.strip().str.lower()
df_val.columns = df_val.columns.str.strip().str.lower()

# Coba cari kolom item code di kedua file
cross_item_col = next((c for c in df_cross.columns if "item" in c and "code" in c), None)
val_item_col = next((c for c in df_val.columns if "item" in c and "code" in c), None)

if not cross_item_col or not val_item_col:
    raise ValueError("❌ Tidak menemukan kolom 'Item Code' di salah satu file!")

# Rename supaya seragam
df_cross.rename(columns={cross_item_col: "ItemCode"}, inplace=True)
df_val.rename(columns={val_item_col: "ItemCode"}, inplace=True)

# Bersihkan karakter dan spasi
df_cross["ItemCode"] = df_cross["ItemCode"].astype(str).str.strip().str.replace(r"\.", "", regex=True)
df_val["ItemCode"] = df_val["ItemCode"].astype(str).str.strip().str.replace(r"\.", "", regex=True)

# ================================
# 🔗 3. GABUNGKAN DATA
# ================================
merged = pd.merge(
    df_cross,
    df_val,
    on="ItemCode",
    how="left",
    suffixes=("_cross", "_val")
)

print(f"🔄 Gabungan awal: {len(merged)} baris")

# ================================
# 🧠 4. PILIH KOLOM SESUAI STRUKTUR
# ================================
final_cols = [
    "Brand", "ItemCode", "Car Maker Name", "Car Model Name", "Car Chassis Name",
    "Car EngineDesc Name", "Car Vehicle Name", "Year From", "Year To", "OEM No.",
    "Part Description", "Alias Name", "Print Description", "Owner", "Number"
]

# Coba temukan nama kolom aslinya dari file validation
def get_col(df, possible_names):
    for name in possible_names:
        for col in df.columns:
            if name.lower() in col:
                return col
    return None

def pick(df, key):
    col = get_col(df, [key])
    return df[col] if col else None

final_df = pd.DataFrame({
    "Brand": pick(merged, "brand"),
    "ItemCode": merged["ItemCode"],
    "Car Maker Name": pick(merged, "car maker"),
    "Car Model Name": pick(merged, "car model"),
    "Car Chassis Name": pick(merged, "chassis"),
    "Car EngineDesc Name": pick(merged, "engine"),
    "Car Vehicle Name": pick(merged, "vehicle"),
    "Year From": pick(merged, "year from"),
    "Year To": pick(merged, "year to"),
    "OEM No.": pick(merged, "oem"),
    "Part Description": pick(merged, "description"),
    "Alias Name": pick(merged, "alias"),
    "Print Description": pick(merged, "print"),
    "Owner": pick(merged, "owner"),
    "Number": pick(merged, "number")
})

# ================================
# 🧹 5. HAPUS DUPLIKAT
# ================================
before = len(final_df)
final_df.drop_duplicates(subset=["ItemCode", "Owner", "Number"], inplace=True)
after = len(final_df)
print(f"🧹 Menghapus {before - after} baris duplikat")

# ================================
# 📦 6. SIMPAN HASIL FINAL
# ================================
output_file = f"Final_Jikiu_Crosses_Clean_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
final_df.to_excel(output_file, index=False)

print(f"\n✅ Proses selesai! File disimpan sebagai: {output_file}")
print(f"📊 Total baris akhir: {len(final_df)}")
