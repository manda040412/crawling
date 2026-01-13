import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime

# ===============================
# KONFIGURASI
# ===============================
EXCEL_PATH = '250.xlsx'
OUTPUT_FILE = f'250_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'

print("📘 Membaca file Excel...")
df = pd.read_excel(EXCEL_PATH)
item_codes = df['ItemCode'].astype(str).tolist()  # semua item
print(f"🔍 Total {len(item_codes)} kode akan diproses.\n")

# ===============================
# SETUP SELENIUM
# ===============================
options = webdriver.ChromeOptions()
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-blink-features=AutomationControlled")
# options.add_argument("--headless")  # bisa aktifkan kalau mau tanpa UI
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 15)

results = []

# ===============================
# PARSER CROSSES
# ===============================
def parse_crosses_from_text(text):
    """Ambil pasangan Owner ↔ Number dari teks halaman Jikiu."""
    if "Crosses" not in text:
        return []

    section = text.split("Crosses", 1)[-1]
    lines = [ln.strip() for ln in section.split("\n") if ln.strip()]

    start_idx = None
    for i, ln in enumerate(lines):
        if ln.lower().startswith("owner"):
            start_idx = i + 1
            break
    if start_idx is None:
        return []

    valid_lines = []
    for ln in lines[start_idx:]:
        if any(stop in ln.lower() for stop in ["application", "brand", "vehicle", "datsun", "nissan »"]):
            break
        if ln and ln.lower() not in ["owner", "number"]:
            valid_lines.append(ln)

    pairs = []
    i = 0
    while i < len(valid_lines) - 1:
        owner = valid_lines[i].strip()
        number = valid_lines[i + 1].strip()
        if (
            len(owner) > 1
            and len(number) > 1
            and not owner.lower().startswith("number")
            and not number.lower().startswith("owner")
        ):
            pairs.append({"Owner": owner, "Number": number})
            i += 2
        else:
            i += 1

    return pairs

# ===============================
# LOOP ITEM
# ===============================
for i, code in enumerate(item_codes, 1):
    print(f"\n[{i}/{len(item_codes)}] 🔎 Cek kode: {code}")

    # Restart browser tiap 100 item biar stabil
    if i % 100 == 0:
        print("♻️ Restarting Chrome session untuk menjaga stabilitas...")
        try:
            driver.quit()
        except:
            pass
        time.sleep(3)
        driver = webdriver.Chrome(options=options)
        wait = WebDriverWait(driver, 15)

    # Auto retry up to 3x if timeout
    success = False
    for attempt in range(3):
        try:
            driver.set_page_load_timeout(60)
            driver.get("https://www.jikiu.com/catalogue")
            success = True
            break
        except Exception as e:
            print(f"⚠️ Percobaan ke-{attempt+1} gagal load halaman ({e}). Ulang...")
            try:
                driver.quit()
            except:
                pass
            time.sleep(5)
            driver = webdriver.Chrome(options=options)
            wait = WebDriverWait(driver, 15)
    if not success:
        print(f"❌ Gagal load halaman untuk {code}, skip.")
        results.append({"Item Code": code, "Owner": "", "Number": ""})
        continue

    try:
        # input pencarian
        search_box = wait.until(EC.presence_of_element_located((By.ID, "part_no")))
        search_box.clear()
        search_box.send_keys(code)
        search_box.send_keys(Keys.ENTER)
        time.sleep(3)

        # cek apakah ditemukan
        if "No data found" in driver.page_source or "0 result" in driver.page_source:
            print("   ✗ Tidak ditemukan di katalog.")
            results.append({"Item Code": code, "Owner": "", "Number": ""})
            continue

        print("   ✓ Data ditemukan, ambil Crosses...")
        page_text = driver.find_element(By.TAG_NAME, "body").text
        crosses_pairs = parse_crosses_from_text(page_text)
        print(f"   ↳ Total pasangan ditemukan: {len(crosses_pairs)}")

        if crosses_pairs:
            for p in crosses_pairs:
                results.append({
                    "Item Code": code,
                    "Owner": p["Owner"],
                    "Number": p["Number"]
                })
        else:
            results.append({"Item Code": code, "Owner": "", "Number": ""})

    except Exception as e:
        print(f"   ❌ Error: {e}")
        results.append({"Item Code": code, "Owner": "", "Number": ""})

        # Auto-recover kalau Chrome crash
        if "Failed to establish a new connection" in str(e) or "Max retries exceeded" in str(e):
            print("🔁 ChromeDriver terputus, mencoba restart...")
            try:
                driver.quit()
            except:
                pass
            time.sleep(5)
            driver = webdriver.Chrome(options=options)
            wait = WebDriverWait(driver, 15)

    # autosave tiap 50 item
    if i % 50 == 0:
        temp_file = f"autosave_{i}_{datetime.now().strftime('%H%M%S')}.xlsx"
        pd.DataFrame(results).to_excel(temp_file, index=False)
        print(f"💾 Autosave progress: {temp_file}")

    time.sleep(2)

# ===============================
# SIMPAN HASIL
# ===============================
try:
    driver.quit()
except:
    pass

result_df = pd.DataFrame(results)

# Gabungkan hasil dengan data asli (jika kolom cocok)
try:
    merged_df = pd.merge(result_df, df, left_on='Item Code', right_on='ItemCode', how='left')
except Exception as e:
    print(f"⚠️ Gagal merge detail tambahan: {e}")
    merged_df = result_df

# Urutan kolom yang diinginkan
ordered_cols = [
    'Item Code', 'Owner', 'Number', 'Brand', 'ItemCode',
    'Car Maker Name', 'Car Model Name', 'Car Chassis Name',
    'Car EngineDesc Name', 'Car Vehicle Name', 'Year From', 'Year To',
    'OEM No.', 'Part Description', 'Alias Name', 'Print Description'
]

# 🔧 hanya ambil kolom yang memang ada
available_cols = [c for c in ordered_cols if c in merged_df.columns]
merged_df = merged_df[available_cols]

# Simpan hasil
merged_df.to_excel(OUTPUT_FILE, index=False)
print(f"\n✅ Selesai! Semua hasil disimpan di: {OUTPUT_FILE}")
