import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import re
from datetime import datetime

EXCEL_PATH = 'List spare parts-Anugerah Auto.xlsx'

# ==============================================
# 1️⃣ BACA DATA EXCEL
# ==============================================
print("📘 Membaca file Excel...")
excel_df = pd.read_excel(EXCEL_PATH)

# Ambil maksimal 100 item
item_codes = excel_df['ItemCode'].astype(str).tolist()[:100]
print(f"Total {len(item_codes)} item akan dicek\n")

# ==============================================
# 2️⃣ SETUP SELENIUM
# ==============================================
print("🚀 Membuka browser...")
options = webdriver.ChromeOptions()
options.add_argument('--window-size=1920,1080')
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_experimental_option("excludeSwitches", ["enable-automation"])
# options.add_argument('--headless')

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 15)

results = []
found_count = 0
not_found_count = 0
error_count = 0

print("="*60)
print("MULAI VALIDASI 100 ITEM")
print("="*60)

start_time = time.time()

# ==============================================
# 3️⃣ LOOP CEK SETIAP ITEM
# ==============================================
for i, original_code in enumerate(item_codes, 1):
    code = str(original_code).strip().rstrip('.')
    progress = f"[{i:3d}/{len(item_codes)}]"

    try:
        driver.get('https://www.jikiu.com/catalogue')
        time.sleep(1)

        # Cari box input
        try:
            search_box = wait.until(EC.presence_of_element_located((By.ID, "part_no")))
        except:
            driver.refresh()
            time.sleep(1)
            search_box = driver.find_element(By.ID, "part_no")

        search_box.clear()
        search_box.send_keys(code)
        search_box.send_keys(Keys.ENTER)
        time.sleep(2)

        page_text = driver.find_element(By.TAG_NAME, 'body').text

        # ==============================
        # CEK STATUS DATA
        # ==============================
        if "No data found!" in page_text or "0 result" in page_text.lower():
            status = "NOT FOUND"
            not_found_count += 1
            jikiu_code = ""
            details = ""
            crosses_str = ""
            print(f"{progress} {code:15} ✗ Tidak ditemukan")

        elif "Search Result for" in page_text or code.upper() in page_text.upper():
            status = "FOUND"
            found_count += 1

            # Deteksi jenis part
            if "BALL JOINT" in page_text:
                item_type = "BALL JOINT"
            elif "TIE ROD END" in page_text:
                item_type = "TIE ROD END"
            elif "STABILIZER LINK" in page_text:
                item_type = "STABILIZER LINK"
            elif "LOWER ARM" in page_text or "BUSHING" in page_text:
                item_type = "LOWER ARM BUSHING"
            elif "RACK END" in page_text:
                item_type = "RACK END"
            else:
                item_type = "OTHER"

            print(f"{progress} {code:15} ✓ Ditemukan ({item_type})")

            # ==============================
            # CARI DAN KLIK LINK PRODUK (KOLUM PART)
            # ==============================
            crosses_all = []

            try:
                part_links = wait.until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'table a'))
                )

                for link in part_links:
                    href = link.get_attribute("href")
                    if href and "/product/" in href:
                        driver.execute_script("window.open(arguments[0]);", href)
                        driver.switch_to.window(driver.window_handles[-1])
                        time.sleep(3)

                        # ==============================================
                        # AMBIL DATA CROSSES DI HALAMAN DETAIL
                        # ==============================================
                        try:
                            # Tunggu halaman detail dimuat
                            wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
                            time.sleep(3)
                            
                            # ==============================================
                            # STRATEGI 1: Cari tab Crosses dan klik
                            # ==============================================
                            try:
                                # Cari semua elemen yang mungkin tab Crosses
                                crosses_tabs = driver.find_elements(By.XPATH, 
                                    "//*[contains(translate(text(), 'CROSSES', 'crosses'), 'crosses') or "
                                    "contains(translate(@class, 'CROSSES', 'crosses'), 'crosses') or "
                                    "contains(translate(@id, 'CROSSES', 'crosses'), 'crosses')]")
                                
                                for tab in crosses_tabs:
                                    if tab.is_displayed() and tab.is_enabled():
                                        print(f"     → Menemukan tab Crosses: {tab.text[:30]}")
                                        driver.execute_script("arguments[0].scrollIntoView(true);", tab)
                                        time.sleep(1)
                                        driver.execute_script("arguments[0].click();", tab)
                                        time.sleep(3)
                                        break
                            except Exception as e:
                                print(f"     → Tidak menemukan tab Crosses: {str(e)[:50]}")
                            
                            # ==============================================
                            # STRATEGI 2: Cari tabel Crosses langsung
                            # ==============================================
                            crosses_found = False
                            
                            # Coba berbagai selector untuk tabel Crosses
                            table_selectors = [
                                'table.detail_plate-crosses',
                                'table.crosses',
                                '.detail_plate-crosses table',
                                '.crosses-table',
                                '.detail_table-crosses',
                                '.detail_plate table',
                                'table[data-type="crosses"]',
                                '.detail_plate.detail_plate-crosses table',
                                '.crosses table',
                                'table.crosses-table'
                            ]
                            
                            for selector in table_selectors:
                                try:
                                    table = driver.find_element(By.CSS_SELECTOR, selector)
                                    if table.is_displayed():
                                        print(f"     → Menemukan tabel Crosses: {selector}")
                                        crosses_found = True
                                        
                                        # Ambil semua baris dari tabel
                                        rows = table.find_elements(By.TAG_NAME, 'tr')
                                        print(f"     → Jumlah baris ditemukan: {len(rows)}")
                                        
                                        for row_idx, row in enumerate(rows):
                                            # Skip header
                                            if row_idx == 0:
                                                continue
                                                
                                            # Coba ambil data dengan berbagai cara
                                            # Cara 1: Ambil dari td cells
                                            cells = row.find_elements(By.TAG_NAME, 'td')
                                            if len(cells) >= 2:
                                                owner = cells[0].text.strip()
                                                number = cells[1].text.strip()
                                                if owner and number and owner.upper() not in ['OWNER', 'OWNER:', 'PEMILIK']:
                                                    crosses_all.append(f"{owner}={number}")
                                                    print(f"       [{row_idx}] {owner}={number}")
                                            
                                            # Cara 2: Ambil dari th/td kombinasi
                                            elif len(cells) == 1:
                                                row_text = row.text.strip()
                                                if row_text and 'Owner' not in row_text and len(row_text) > 5:
                                                    # Coba split dengan spasi atau tab
                                                    parts = re.split(r'\s{2,}|\t', row_text)
                                                    if len(parts) >= 2:
                                                        owner = parts[0].strip()
                                                        number = parts[1].strip()
                                                        if owner and number:
                                                            crosses_all.append(f"{owner}={number}")
                                                            print(f"       [{row_idx}] {owner}={number}")
                                            
                                            # Cara 3: Coba dengan div dalam row
                                            divs = row.find_elements(By.TAG_NAME, 'div')
                                            if len(divs) >= 2 and not cells:
                                                owner = divs[0].text.strip()
                                                number = divs[1].text.strip()
                                                if owner and number:
                                                    crosses_all.append(f"{owner}={number}")
                                                    print(f"       [{row_idx}] {owner}={number}")
                                        
                                        break
                                except Exception as e:
                                    continue
                            
                            # ==============================================
                            # STRATEGI 3: Jika tabel tidak ditemukan, cari div-based layout
                            # ==============================================
                            if not crosses_found:
                                print("     → Mencari layout Crosses berbasis div...")
                                
                                # Coba cari container Crosses
                                container_selectors = [
                                    '.detail_plate-crosses',
                                    '.crosses-container',
                                    '.crosses_list',
                                    '[id*="cross"]',
                                    '[class*="cross"]',
                                    '.detail-plate-crosses',
                                    '.crosses-list',
                                    '.crosses-content'
                                ]
                                
                                for selector in container_selectors:
                                    try:
                                        container = driver.find_element(By.CSS_SELECTOR, selector)
                                        if container.is_displayed():
                                            print(f"     → Menemukan container Crosses: {selector}")
                                            
                                            # Cari semua pasangan owner/number
                                            owners = container.find_elements(By.CSS_SELECTOR, 
                                                '.detail_field, .owner, [class*="owner"], .crosses-owner, .w200')
                                            numbers = container.find_elements(By.CSS_SELECTOR, 
                                                '.detail_value, .number, [class*="number"], .crosses-number')
                                            
                                            if owners and numbers:
                                                print(f"     → Found {len(owners)} owners and {len(numbers)} numbers")
                                                min_len = min(len(owners), len(numbers))
                                                for i in range(min_len):
                                                    owner_text = owners[i].text.strip()
                                                    number_text = numbers[i].text.strip()
                                                    if owner_text and number_text and 'Owner' not in owner_text:
                                                        crosses_all.append(f"{owner_text}={number_text}")
                                                        print(f"       [{i}] {owner_text}={number_text}")
                                            
                                            # Coba ambil dari grid layout
                                            grid_items = container.find_elements(By.CSS_SELECTOR, '.row, .grid-item, .crosses-item')
                                            for item in grid_items:
                                                item_text = item.text.strip()
                                                if item_text and 'Owner' not in item_text and len(item_text) > 5:
                                                    # Coba split
                                                    parts = item_text.split('\n')
                                                    if len(parts) >= 2:
                                                        crosses_all.append(f"{parts[0]}={parts[1]}")
                                                    
                                            break
                                    except:
                                        continue
                            
                            # ==============================================
                            # STRATEGI 4: Ambil dari semua teks di halaman
                            # ==============================================
                            if not crosses_all:
                                print("     → Mencari Crosses dari seluruh teks halaman...")
                                
                                # Ambil semua teks halaman
                                page_text = driver.find_element(By.TAG_NAME, 'body').text
                                
                                # Cari bagian yang mengandung kata "Crosses"
                                lines = page_text.split('\n')
                                in_crosses_section = False
                                crosses_section_lines = []
                                
                                for line in lines:
                                    line = line.strip()
                                    if 'crosses' in line.lower():
                                        in_crosses_section = True
                                        continue
                                    
                                    if in_crosses_section and line:
                                        # Skip header
                                        if 'owner' in line.lower() or 'number' in line.lower():
                                            continue
                                        
                                        crosses_section_lines.append(line)
                                
                                # Parse lines dari section Crosses
                                for line in crosses_section_lines:
                                    if '=' in line:
                                        crosses_all.append(line)
                                    elif len(line) > 5:
                                        # Split by multiple spaces
                                        parts = re.split(r'\s{2,}', line)
                                        if len(parts) >= 2:
                                            crosses_all.append(f"{parts[0]}={parts[1]}")
                            
                            # ==============================================
                            # STRATEGI 5: Cari dengan XPath spesifik
                            # ==============================================
                            if not crosses_all:
                                try:
                                    print("     → Mencari dengan XPath spesifik...")
                                    # XPath untuk elemen yang mungkin berisi data Crosses
                                    xpath_elements = driver.find_elements(By.XPATH, 
                                        "//div[contains(@class, 'crosses')]//div[contains(@class, 'row')] | "
                                        "//div[contains(@class, 'crosses')]//div[contains(@class, 'item')] | "
                                        "//table[contains(@class, 'crosses')]//tr")
                                    
                                    for elem in xpath_elements:
                                        elem_text = elem.text.strip()
                                        if elem_text and len(elem_text) > 5:
                                            if '=' in elem_text:
                                                crosses_all.append(elem_text)
                                            else:
                                                parts = elem_text.split()
                                                if len(parts) >= 2:
                                                    crosses_all.append(f"{parts[0]}={parts[1]}")
                                except:
                                    pass
                            
                            # Filter duplikat, kosong, dan header
                            crosses_all = [c for c in crosses_all if c and '=' in c and len(c) > 3]
                            crosses_all = [c for c in crosses_all if 'owner' not in c.lower()]
                            crosses_all = list(dict.fromkeys(crosses_all))
                            
                            if crosses_all:
                                print(f"     → Berhasil mengumpulkan {len(crosses_all)} data Crosses")
                                # Tampilkan 5 contoh pertama
                                for idx, cross in enumerate(crosses_all[:5]):
                                    print(f"       Contoh {idx+1}: {cross}")
                                if len(crosses_all) > 5:
                                    print(f"       ... dan {len(crosses_all)-5} data lainnya")
                            else:
                                print(f"     → Tidak menemukan data Crosses yang valid")
                                
                                # DEBUG: Simpan screenshot untuk analisis
                                try:
                                    debug_name = f"debug_{code}_{int(time.time())}"
                                    driver.save_screenshot(f"{debug_name}.png")
                                    with open(f"{debug_name}.html", "w", encoding="utf-8") as f:
                                        f.write(driver.page_source)
                                    print(f"     ⚠ Debug files saved: {debug_name}.png/.html")
                                except:
                                    pass
                            
                        except TimeoutException:
                            print(f"     ⚠ Timeout saat menunggu halaman detail")
                        except Exception as e:
                            error_msg = str(e)[:100]
                            print(f"     ⚠ Error saat ambil Crosses: {error_msg}")

                        # Tutup tab detail dan kembali ke tab utama
                        driver.close()
                        driver.switch_to.window(driver.window_handles[0])
                        time.sleep(1)

                crosses_str = "; ".join(crosses_all) if crosses_all else "No Crosses Found"
                if crosses_all:
                    print(f"     ↳ Total: {len(crosses_all)} data Crosses dikumpulkan")

            except Exception as e:
                if "NoSuchElementException" in str(type(e)) or "TimeoutException" in str(type(e)):
                    crosses_str = "No Crosses Available"
                    print(f"     ⚠ Crosses tidak tersedia")
                else:
                    crosses_str = f"Crosses Error: {str(e)[:80]}"
                    print(f"     ⚠ {crosses_str}")

            # Simpan data hasil
            details = f"{item_type}"
            jikiu_code = ""  # (opsional ambil dari halaman detail jika perlu)

        else:
            status = "CHECK MANUAL"
            details = "Response tidak jelas"
            jikiu_code = ""
            crosses_str = ""
            print(f"{progress} {code:15} ? Perlu dicek manual")

        # ==============================
        # SIMPAN DATA KE RESULT
        # ==============================
        results.append({
            'No': i,
            'Item Code': original_code,
            'Cleaned Code': code,
            'Status': status,
            'JIKIU Code': jikiu_code,
            'Details': details,
            'Crosses': crosses_str,
            'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })

    except Exception as e:
        error_count += 1
        error_msg = str(e)[:120]
        print(f"{progress} {code:15} ❌ Error: {error_msg}")

        results.append({
            'No': i,
            'Item Code': original_code,
            'Cleaned Code': code,
            'Status': 'ERROR',
            'JIKIU Code': '',
            'Details': error_msg,
            'Crosses': '',
            'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })

    if i < len(item_codes):
        time.sleep(1)

# ==============================================
# 4️⃣ SIMPAN HASIL KE EXCEL
# ==============================================
end_time = time.time()
execution_time = end_time - start_time

print("\n" + "="*60)
print("HASIL VALIDASI 100 ITEM")
print("="*60)
print(f"✓ DITEMUKAN      : {found_count} item")
print(f"✗ TIDAK DITEMUKAN: {not_found_count} item")
print(f"❌ ERROR          : {error_count} item")
print(f"⏱ WAKTU EKSEKUSI : {execution_time:.1f} detik")

# Buat DataFrame hasil scraping
result_df = pd.DataFrame(results)

# Gabungkan dengan data asli dari Excel
merged_df = pd.merge(
    result_df,
    excel_df,
    left_on='Item Code',
    right_on='ItemCode',
    how='left'
)

# Urutkan kolom biar rapi
ordered_cols = [
    'No', 'Item Code', 'Cleaned Code', 'Status', 'JIKIU Code', 'Details', 'Crosses', 'Timestamp',
    'Brand', 'Car Maker Name', 'Car Model Name', 'Car Chassis Name', 'Car EngineDesc Name',
    'Car Vehicle Name', 'Year From', 'Year To', 'OEM No.', 'Part Description',
    'Alias Name', 'Print Description'
]
for col in merged_df.columns:
    if col not in ordered_cols:
        ordered_cols.append(col)

merged_df = merged_df[ordered_cols]

# Simpan hasil
output_file = f'validation_100_items_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
merged_df.to_excel(output_file, index=False)

print(f"\n📁 Hasil disimpan di: {output_file}")
driver.quit()
print("\n✅ Validasi selesai!")