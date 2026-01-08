import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re
from datetime import datetime

EXCEL_PATH = 'List spare parts-Anugerah Auto.xlsx'

# Read Excel file
print("Membaca file Excel...")
excel_df = pd.read_excel(EXCEL_PATH)

# Ambil 100 item pertama
item_codes = excel_df['ItemCode'].astype(str).tolist()[:100]
print(f"Total {len(item_codes)} item akan dicek\n")

# Setup Chrome driver
print("Membuka browser...")
options = webdriver.ChromeOptions()
options.add_argument('--window-size=1920,1080')
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_experimental_option("excludeSwitches", ["enable-automation"])

# Untuk lebih cepat, bisa pakai headless
# options.add_argument('--headless')

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 5)  # Timeout lebih pendek untuk lebih cepat

results = []
found_count = 0
not_found_count = 0
error_count = 0

print("="*60)
print("MULAI VALIDASI 100 ITEM")
print("="*60)

start_time = time.time()

for i, original_code in enumerate(item_codes, 1):
    # Bersihkan kode
    code = str(original_code).strip().rstrip('.')
    
    # Progress indicator
    progress = f"[{i:3d}/{len(item_codes)}]"
    
    try:
        # Buka halaman catalogue
        driver.get('https://www.jikiu.com/catalogue')
        time.sleep(1)  # Tunggu singkat
        
        # Cari search box
        try:
            search_box = wait.until(EC.presence_of_element_located((By.ID, "part_no")))
        except:
            # Coba refresh jika elemen tidak ditemukan
            driver.refresh()
            time.sleep(1)
            search_box = driver.find_element(By.ID, "part_no")
        
        # Input kode
        search_box.clear()
        search_box.send_keys(code)
        
        # Tekan Enter
        search_box.send_keys(Keys.ENTER)
        time.sleep(2)  # Tunggu hasil
        
        # Ambil text halaman
        page_text = driver.find_element(By.TAG_NAME, 'body').text
        
        # Analisis cepat
        if "No data found!" in page_text or "0 result" in page_text.lower():
            status = "NOT FOUND"
            not_found_count += 1
            print(f"{progress} {code:15} ✗ Tidak ditemukan")
            details = ""
            jikiu_code = ""
            
        elif "Search Result for" in page_text or code.upper() in page_text.upper():
            status = "FOUND"
            found_count += 1
            
            # Coba extract JIKIU code
            jikiu_match = re.search(r'Returns JIKIU - (\w+)', page_text)
            jikiu_code = jikiu_match.group(1) if jikiu_match else ""
            
            # Cek tipe item
            if "BALL JOINT" in page_text:
                item_type = "BALL JOINT"
            elif "TIE ROD END" in page_text:
                item_type = "TIE ROD END"
            elif "STABILIZER LINK" in page_text:
                item_type = "STABILIZER LINK"
            else:
                item_type = "OTHER"
            
            details = f"{item_type} - {jikiu_code}" if jikiu_code else item_type
            print(f"{progress} {code:15} ✓ Ditemukan ({item_type})")
            
        else:
            # Kasus ambigu, cek manual
            status = "CHECK MANUAL"
            print(f"{progress} {code:15} ? Perlu dicek manual")
            details = "Response tidak jelas"
            jikiu_code = ""
        
        results.append({
            'No': i,
            'Item Code': original_code,
            'Cleaned Code': code,
            'Status': status,
            'JIKIU Code': jikiu_code,
            'Details': details,
            'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })
        
    except Exception as e:
        error_count += 1
        error_msg = str(e)[:80]
        print(f"{progress} {code:15} ❌ Error: {error_msg}")
        
        results.append({
            'No': i,
            'Item Code': original_code,
            'Cleaned Code': code,
            'Status': 'ERROR',
            'JIKIU Code': '',
            'Details': error_msg,
            'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })
    
    # Delay antar request (untuk menghindari block)
    if i < len(item_codes):  # Tidak delay untuk item terakhir
        time.sleep(0.5)

# Hitung waktu eksekusi
end_time = time.time()
execution_time = end_time - start_time

print("\n" + "="*60)
print("HASIL VALIDASI 100 ITEM")
print("="*60)
print(f"✓ DITEMUKAN     : {found_count} item")
print(f"✗ TIDAK DITEMUKAN: {not_found_count} item")
print(f"❌ ERROR         : {error_count} item")
print(f"⏱ WAKTU EKSEKUSI: {execution_time:.1f} detik")
print(f"📊 RATA-RATA     : {execution_time/len(item_codes):.2f} detik/item")

# Simpan ke Excel
result_df = pd.DataFrame(results)

# Gabungkan dengan data asli jika ada kolom lain
try:
    # Cari kolom selain ItemCode
    other_columns = [col for col in excel_df.columns if col != 'ItemCode']
    if other_columns:
        # Gabungkan berdasarkan kode
        for col in other_columns:
            # Buat mapping dari data asli
            mapping = dict(zip(excel_df['ItemCode'].astype(str), excel_df[col]))
            result_df[col] = result_df['Item Code'].astype(str).map(mapping)
except:
    pass

# Urutkan kolom
column_order = ['No', 'Item Code', 'Cleaned Code', 'Status', 'JIKIU Code', 'Details', 'Timestamp']
other_cols = [col for col in result_df.columns if col not in column_order]
result_df = result_df[column_order + other_cols]

# Simpan ke file
output_file = f'validation_100_items_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
result_df.to_excel(output_file, index=False)

print(f"\n📁 Hasil disimpan di: {output_file}")

# Tampilkan 10 item pertama yang ditemukan
found_items = result_df[result_df['Status'] == 'FOUND'].head(10)
if len(found_items) > 0:
    print("\n🔍 CONTOH ITEM YANG DITEMUKAN:")
    print("-" * 50)
    for _, row in found_items.iterrows():
        print(f"{row['Item Code']} → {row['Details']}")

# Tampilkan 10 item yang tidak ditemukan
not_found_items = result_df[result_df['Status'] == 'NOT FOUND'].head(10)
if len(not_found_items) > 0:
    print("\n❌ CONTOH ITEM YANG TIDAK DITEMUKAN:")
    print("-" * 50)
    for _, row in not_found_items.iterrows():
        print(f"{row['Item Code']}")

driver.quit()
print("\n✅ Validasi selesai!")