import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
from urllib.parse import quote
import json

class JikiuCrawler:
    def __init__(self):
        self.base_url = "https://www.jikiu.com/catalogue"
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Accept-Encoding': 'gzip, deflate',
            'Connection': 'keep-alive',
        })
        
    def search_part(self, item_code):
        """Search for a part on Jikiu website"""
        try:
            # Search URL
            search_url = f"{self.base_url}/search?part={quote(item_code)}"
            print(f"Searching: {item_code}")
            
            response = self.session.get(search_url, timeout=10)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Check if part is found
            if "no results" in response.text.lower() or "not found" in response.text.lower():
                return {
                    'found': False,
                    'url': search_url,
                    'item_code': item_code
                }
            
            # Extract specifications
            specs = self.extract_specifications(soup)
            
            # Extract crosses/alternative part numbers
            crosses = self.extract_crosses(soup)
            
            return {
                'found': True,
                'url': search_url,
                'item_code': item_code,
                'specifications': specs,
                'crosses': crosses
            }
            
        except requests.exceptions.RequestException as e:
            print(f"Error fetching {item_code}: {e}")
            return {
                'found': False,
                'error': str(e),
                'url': search_url if 'search_url' in locals() else '',
                'item_code': item_code
            }
    
    def extract_specifications(self, soup):
        """Extract specifications from the page"""
        specs = {}
        
        # Method 1: Look for specification table/section
        spec_section = soup.find('div', class_='specification') or soup.find('section', class_='specification')
        
        if spec_section:
            # Try to find all specification rows
            spec_rows = spec_section.find_all(['tr', 'div'], class_=['spec-row', 'specification-item'])
            
            for row in spec_rows:
                label = row.find(['th', 'dt', 'span', 'div'], class_=['label', 'spec-label', 'key'])
                value = row.find(['td', 'dd', 'span', 'div'], class_=['value', 'spec-value', 'val'])
                
                if label and value:
                    key = label.get_text(strip=True).replace(':', '')
                    val = value.get_text(strip=True)
                    specs[key] = val
        
        # Method 2: Look for specific specification fields
        spec_fields = [
            'Cone Pitch', 'Cone Size', 'Thread Size', 
            'Overall Height', 'Diameter', 'Mounting Height',
            'Location', 'Position'
        ]
        
        for field in spec_fields:
            if field not in specs:
                # Try to find by text content
                element = soup.find(text=lambda t: t and field.lower() in t.lower())
                if element:
                    parent = element.parent
                    # Try to get the value from next sibling or parent
                    value_elem = parent.find_next_sibling() or parent.find('span', class_='value')
                    if value_elem:
                        specs[field] = value_elem.get_text(strip=True)
        
        return specs
    
    def extract_crosses(self, soup):
        """Extract crosses/alternative part numbers"""
        crosses = []
        
        # Look for crosses section
        crosses_section = soup.find(['div', 'section'], class_=['crosses', 'cross-references', 'alternatives'])
        
        if crosses_section:
            # Look for table
            table = crosses_section.find('table')
            if table:
                rows = table.find_all('tr')[1:]  # Skip header
                for row in rows:
                    cols = row.find_all('td')
                    if len(cols) >= 2:
                        crosses.append({
                            'owner': cols[0].get_text(strip=True),
                            'number': cols[1].get_text(strip=True)
                        })
            else:
                # Look for list items
                items = crosses_section.find_all(['li', 'div'], class_=['cross-item', 'alternative'])
                for item in items:
                    owner_elem = item.find(['span', 'strong'], class_=['owner', 'brand'])
                    number_elem = item.find(['span', 'code'], class_=['number', 'part-number'])
                    
                    if owner_elem and number_elem:
                        crosses.append({
                            'owner': owner_elem.get_text(strip=True),
                            'number': number_elem.get_text(strip=True)
                        })
        
        return crosses
    
    def process_excel(self, input_file, output_file='Jikiu_Crawl_Results.xlsx'):
        """Process Excel file and crawl data"""
        print(f"Reading Excel file: {input_file}")
        
        # Read Excel file
        df = pd.read_excel(input_file)
        
        print(f"Found {len(df)} rows to process")
        
        # Initialize result columns
        df['Found_in_Jikiu'] = False
        df['Jikiu_URL'] = ''
        df['Jikiu_Cone_Pitch'] = ''
        df['Jikiu_Cone_Size_mm'] = ''
        df['Jikiu_Thread_Size'] = ''
        df['Jikiu_Overall_Height_mm'] = ''
        df['Jikiu_Diameter_mm'] = ''
        df['Jikiu_Mounting_Height_mm'] = ''
        df['Jikiu_Location'] = ''
        df['Jikiu_Position'] = ''
        df['Jikiu_Crosses'] = ''
        df['Crawl_Error'] = ''
        
        # Process each row
        for idx, row in df.iterrows():
            # Get ItemCode (handle different column name variations)
            item_code = row.get('ItemCode') or row.get('Item Code') or row.get('ITEM CODE') or ''
            
            if not item_code or pd.isna(item_code):
                print(f"Row {idx + 1}: No ItemCode found, skipping...")
                continue
            
            # Crawl the part
            result = self.search_part(str(item_code))
            
            # Update DataFrame
            df.at[idx, 'Found_in_Jikiu'] = result.get('found', False)
            df.at[idx, 'Jikiu_URL'] = result.get('url', '')
            
            if result.get('found'):
                specs = result.get('specifications', {})
                
                # Map specifications to columns
                df.at[idx, 'Jikiu_Cone_Pitch'] = specs.get('Cone Pitch', '')
                df.at[idx, 'Jikiu_Cone_Size_mm'] = specs.get('Cone Size Ø (mm)', specs.get('Cone Size', ''))
                df.at[idx, 'Jikiu_Thread_Size'] = specs.get('Thread Size', '')
                df.at[idx, 'Jikiu_Overall_Height_mm'] = specs.get('Overall Height (mm)', specs.get('Overall Height', ''))
                df.at[idx, 'Jikiu_Diameter_mm'] = specs.get('Ø (mm)', specs.get('Diameter', ''))
                df.at[idx, 'Jikiu_Mounting_Height_mm'] = specs.get('Mounting Height (mm)', specs.get('Mounting Height', ''))
                df.at[idx, 'Jikiu_Location'] = specs.get('Location', '')
                df.at[idx, 'Jikiu_Position'] = specs.get('Position', '')
                
                # Format crosses
                crosses = result.get('crosses', [])
                if crosses:
                    crosses_str = '; '.join([f"{c['owner']}: {c['number']}" for c in crosses])
                    df.at[idx, 'Jikiu_Crosses'] = crosses_str
            
            if 'error' in result:
                df.at[idx, 'Crawl_Error'] = result['error']
            
            # Progress update
            print(f"Progress: {idx + 1}/{len(df)} - {item_code} - {'FOUND' if result.get('found') else 'NOT FOUND'}")
            
            # Delay to avoid overwhelming the server
            time.sleep(1)
        
        # Save results
        print(f"\nSaving results to {output_file}")
        df.to_excel(output_file, index=False)
        
        # Print summary
        found_count = df['Found_in_Jikiu'].sum()
        not_found_count = len(df) - found_count
        
        print("\n" + "="*50)
        print("CRAWLING COMPLETED!")
        print("="*50)
        print(f"Total Items: {len(df)}")
        print(f"Found in Jikiu: {found_count}")
        print(f"Not Found: {not_found_count}")
        print(f"Success Rate: {(found_count/len(df)*100):.1f}%")
        print(f"\nResults saved to: {output_file}")
        
        return df


def main():
    """Main function to run the crawler"""
    print("="*50)
    print("JIKIU WEB CRAWLER")
    print("="*50)
    print()
    
    # Initialize crawler
    crawler = JikiuCrawler()
    
    # Input file name
    input_file = 'List spare parts-Anugerah Auto.xlsx'
    output_file = 'Jikiu_Crawl_Results.xlsx'
    
    # Check if file exists
    try:
        crawler.process_excel(input_file, output_file)
    except FileNotFoundError:
        print(f"\nError: File '{input_file}' not found!")
        print("Please make sure the Excel file is in the same directory as this script.")
    except Exception as e:
        print(f"\nError occurred: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()