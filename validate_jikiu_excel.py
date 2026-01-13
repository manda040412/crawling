import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re
from datetime import datetime
import os

# --- KONFIGURASI ---
EXCEL_PATH = 'List spare parts-Anugerah Auto.xlsx'
SAVE_INTERVAL = 50  # Simpan progres setiap 50 item
OUTPUT_FINAL = f'validation_FULL_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
BACKUP_FILE = 'backup_validation_progress.xlsx'

# 1. Membaca file Excel
print(f"[{datetime.now().strftime('%H:%M:%S')}] Membaca file Excel...")
try:
    excel_df = pd.read_excel(EXCEL_PATH)
    # Ambil SEMUA item tanpa limit [:100]
    item_codes = excel_df['ItemCode'].astype(str).tolist()
    print(f"Total {len(item_codes)} item ditemukan dalam file.\n")
except Exception as e:
    print(f"Gagal membaca file Excel: {e}")
    exit()

# 2. Setup Chrome Driver
print("Menyiapkan browser (Headless Mode)...")
options = webdriver.ChromeOptions()
options.add_argument('--headless')  # Berjalan di background (lebih cepat)
options.add_argument('--window-size=1920,1080')
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_experimental_option("excludeSwitches", ["enable-automation"])

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 10)

results = []
found_count = 0
not_found_count = 0
error_count = 0

print("="*60)
print(f"MULAI VALIDASI TOTAL {len(item_codes)} ITEM")
print("="*60)

start_time = time.time()

# 3. Looping Utama
try:
    for i, original_code in enumerate(item_codes, 1):
        code = str(original_code).strip().rstrip('.')
        progress = f"[{i}/{len(item_codes)}]"
        
        # Retry logic sederhana jika koneksi timeout
        retry = 0
        success_fetch = False
        
        while retry < 2 and not success_fetch:
            try:
                driver.get('https://www.jikiu.com/catalogue')
                search_box = wait.until(EC.presence_of_element_located((By.ID, "part_no")))
                
                search_box.clear()
                search_box.send_keys(code)
                search_box.send_keys(Keys.ENTER)
                
                # Tunggu sejenak agar hasil muncul (AJAX load)
                time.sleep(1.5)
                
                page_text = driver.find_element(By.TAG_NAME, 'body').text
                success_fetch = True
            except Exception:
                retry += 1
                time.sleep(2)

        if not success_fetch:
            status, details, jikiu_code = "ERROR", "Connection Timeout", ""
            error_count += 1
        else:
            # Analisis Hasil
            if "No data found!" in page_text or "0 result" in page_text.lower():
                status = "NOT FOUND"
                not_found_count += 1
                details, jikiu_code = "", ""
                print(f"{progress} {code:15} ✗ Tidak ditemukan")
                
            elif "Search Result for" in page_text or code.upper() in page_text.upper():
                status = "FOUND"
                found_count += 1
                
                jikiu_match = re.search(r'Returns JIKIU - (\w+)', page_text)
                jikiu_code = jikiu_match.group(1) if jikiu_match else ""
                
                # Deteksi Tipe Item
                types = ["BALL JOINT", "TIE ROD END", "STABILIZER LINK", "RACK END", "IDLER ARM"]
                item_type = "OTHER"
                for t in types:
                    if t in page_text.upper():
                        item_type = t
                        break
                
                details = f"{item_type} - {jikiu_code}" if jikiu_code else item_type
                print(f"{progress} {code:15} ✓ Ditemukan ({item_type})")
            else:
                status = "CHECK MANUAL"
                details, jikiu_code = "Ambiguous Response", ""
                print(f"{progress} {code:15} ? Cek Manual")

        # Simpan ke list hasil
        results.append({
            'No': i,
            'Item Code': original_code,
            'Status': status,
            'JIKIU Code': jikiu_code,
            'Details': details,
            'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })

        # 4. Auto-Save berkala
        if i % SAVE_INTERVAL == 0:
            pd.DataFrame(results).to_excel(BACKUP_FILE, index=False)
            elapsed = time.time() - start_time
            print(f"--- AUTO-SAVE: {i} item selesai. Durasi: {elapsed/60:.1f} menit ---")

except KeyboardInterrupt:
    print("\nProses dihentikan paksa oleh pengguna. Menyimpan data yang ada...")

# 5. Finalisasi Data
print("\n" + "="*60)
print("PROSES SELESAI")
final_df = pd.DataFrame(results)

# Mapping kembali kolom lain dari file asli (opsional)
try:
    other_columns = [col for col in excel_df.columns if col != 'ItemCode']
    for col in other_columns:
        mapping = dict(zip(excel_df['ItemCode'].astype(str), excel_df[col]))
        final_df[col] = final_df['Item Code'].astype(str).map(mapping)
except:
    pass

# Simpan hasil akhir
final_df.to_excel(OUTPUT_FINAL, index=False)
if os.path.exists(BACKUP_FILE):
    os.remove(BACKUP_FILE) # Hapus backup jika sukses sampai akhir

end_time = time.time()
print(f"Total Waktu: {(end_time - start_time)/60:.2f} menit")
print(f"Ditemukan: {found_count} | Tidak: {not_found_count} | Error: {error_count}")
print(f"File disimpan: {OUTPUT_FINAL}")

driver.quit()