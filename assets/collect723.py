AGENTC_SERPER_API_KEY = "c213ee9acbb70eb39a32c3da91ff8acffb8778e7"
APS7140_SERPER_API_KEY = "7758de7844f957cba53fd28049598694b60a64cf"

import os
import json
import requests
import pandas as pd
from tqdm import tqdm

# --- CONFIGURATION ---
SERPER_API_KEY = "b57ded568e853310d963f81985d90c508a86d987"
INPUT_EXCEL = "assets/exportUsers_2026-7-13.xlsx"
NAME_COLUMN = "displayName"
OUTPUT_OSINT_JSON = "septa_social_links.json"   
OUTPUT_METADATA_CSV = "document_metadata.csv" 
DOWNLOAD_DIR = "./downloaded_docs"           

TEST_LIMIT = 1500

os.makedirs(DOWNLOAD_DIR, exist_ok=True)


def search_employee(name):
    """Queries Serper using the exact playground POST method structure."""
    url = "https://google.serper.dev/search"
    query = f'"{name}" SEPTA'
    
    payload = {
        "q": query,
    }

    headers = {
        'X-API-KEY': SERPER_API_KEY,
        'Content-Type': 'application/json'
    }
    
    # This dictionary structures all links directly under the employee's name
    employee_record = {
        "employee_name": name,
        "total_links_uncovered": 0,
        "discovered_assets": []
    }

    file_urls = []
    
    try:
        response = requests.request("POST", url, headers=headers, json=payload)
        
        if response.status_code == 200:
            data = response.json()
            organic_results = data.get("organic", [])
            
            for item in organic_results:
                link = item.get("link", "")
                is_doc = False

                if link.lower().endswith(('.pdf', '.docx', '.xlsx', '.pptx', '.doc', '.xls')):
                    file_urls.append(link)
                    is_doc = True
                
                employee_record["discovered_assets"].append({
                    "result_title": item.get("title", ""),
                    "url": link,
                    "snippet": item.get("snippet", ""),
                    "is_doc": is_doc
                })
            
            employee_record["total_links_uncovered"] = len(employee_record["discovered_assets"])
        else:
            print(f"\n [!] Error: API responded with status {response.status_code} for {name}")
    
    except Exception as e:
        print(f"\n [!] Network connection exception for {name}: {e}")
        
    return employee_record, file_urls


def download_file(url):
    """Downloads target files discovered during the search phase."""
    try:
        filename = os.path.join(DOWNLOAD_DIR, url.split("/")[-1].split("?")[0])
        response = requests.request("GET", url)
        if response.status_code == 200:
            with open(filename, 'wb') as f:
                f.write(response.content)
    except Exception:
        pass


def main():
    print(f"Accessing spreadsheet file input: {INPUT_EXCEL}...")
    try:
        df = pd.read_excel(INPUT_EXCEL, engine='openpyxl')
    except Exception as e:
        print(f"[!] Error loading Excel file: {e}")
        return

    employee_names = df[NAME_COLUMN].dropna().unique().tolist()
    
    if TEST_LIMIT is not None:
        print(f"Total file records: {len(employee_names)} names.")
        employee_names = employee_names[:TEST_LIMIT]
        print(f"[!] TEST MODE ACTIVE: Slicing list down to {len(employee_names)} test targets.")
    else:
        print(f"Processing run across all {len(employee_names)} targets.")

    master_json_output = []
    all_file_urls = []
    
    # Process targets sequentially
    for name in tqdm(employee_names, desc="Searching Google"):
        employee_record, file_data = search_employee(name)
        
        if employee_record["total_links_uncovered"] > 0:
            master_json_output.append(employee_record)
            
        all_file_urls.extend(file_data)
    
    print(f"\nTotal profiled employee profiles with records: {len(master_json_output)}")
    print(f"Total direct binary document tracks isolated: {len(all_file_urls)}")

    # Handle file downloads sequentially for ExifTool processing
    if all_file_urls:
        print("Downloading discovered documents for ExifTool processing...")
        for url in tqdm(set(all_file_urls), desc="Downloading Assets"):
            download_file(url)
            
        print("Running ExifTool forensic metadata extraction mapping...")
        os.system(f"exiftool -csv -r {DOWNLOAD_DIR} > {OUTPUT_METADATA_CSV}")
        print(f"ExifTool analytical tracking report compiled: {OUTPUT_METADATA_CSV}")
    else:
        print("No target document files found. Skipping ExifTool phase.")

    if master_json_output:
        with open(OUTPUT_OSINT_JSON, "w", encoding="utf-8") as f:
            json.dump(master_json_output, f, indent=2, ensure_ascii=False)
        print(f"Master JSON intelligence file successfully saved to: {OUTPUT_OSINT_JSON}")

if __name__ == "__main__":
    main()