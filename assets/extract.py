import os
import json
import re
import time
from urllib.parse import urlparse
import requests
import pandas as pd
from openpyxl.utils import get_column_letter

# --- CONFIGURATION ---
INPUT_EXCEL_USERS = "assets/exportUsers_2026-7-13.xlsx"
INPUT_RAW_JSON = "septa_social_links.json"               
OUTPUT_EXCEL_REPORT = "septa_employee_digital_footprint.xlsx"

PHILLY_CARTO_API_URL = "https://phl.carto.com/api/v2/sql"

from requests.packages.urllib3.exceptions import InsecureRequestWarning
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

def fetch_residential_address(employee_name):
    """
    Queries Open Data Philly OPA Property registers via Carto SQL API.
    Uses optimized wildcard truncation to increase match accuracy.
    """
    if not employee_name:
        return ""

    clean_name = re.sub(r'[^\w\s]', '', employee_name).strip().upper()
    name_parts = clean_name.split()

    if len(name_parts) < 2:
        return "No Matches"

    # Format query to adapt to OPA naming standard: "LASTNAME FIRSTINITIAL%"
    first_name = name_parts[0]
    last_name = name_parts[-1]
    
    # Use Last Name plus the first character of the First Name to fix formatting mismatches
    primary_search = f"{last_name} {first_name[0]}"
    sql_query = f"SELECT location FROM opa_properties_public WHERE owner_1 LIKE '{primary_search}%' LIMIT 3"
    
    try:
        # A tiny delay prevents rate limits from dropping packets
        time.sleep(0.3)
        payload = {'q': sql_query}
        
        session = requests.Session()
        session.trust_env = False  
        response = session.get(PHILLY_CARTO_API_URL, params=payload, timeout=7, verify=False)

        if response.status_code == 200:
            data = response.json()
            rows = data.get("rows", [])
            if rows:
                addresses = [row.get("location", "").strip() for row in rows if row.get("location")]
                return " | ".join(list(set(addresses))) if addresses else "No Matches"
            else:
                return "No Matches"
    except Exception:
        return "API Timeout / Offline"

    return "No Matches"


def clean_job_title(raw_segment):
    """
    Strips fluff, filler words, and trailing text 
    to extract a clean, concise job title.
    """
    if not raw_segment or raw_segment == "":
        return ""
        
    # Strip common leading filler text variations
    cleaned = re.sub(r'^(is currently an?|working as an?|a hired|a professional|an?)\s+', '', raw_segment, flags=re.IGNORECASE)
    
    # Truncate anything following corporate employer transitions
    cleaned = re.split(r'\s+(at|for|with|brings|from|in)\s+', cleaned, flags=re.IGNORECASE)[0]
    
    # Strip trailing noisy conjunction text blocks
    cleaned = re.split(r'\s+(and his|and her|and their|,)\s+', cleaned, flags=re.IGNORECASE)[0]
    
    # Capitalize title blocks neatly
    return cleaned.strip().title()


def dynamic_intel_extraction(url, title, snippet, employee_name):
    """Targeted extraction core cleaning text variables into strict formatting schemas."""
    if not url:
        return {}

    parsed_url = urlparse(url)
    domain = parsed_url.netloc.lower().replace("www.", "")
    path = parsed_url.path
    combined_text = f"{title} | {snippet}"

    linkedin_profile = ""
    other_socials = []
    social_urls_only = []
    deduced_title = ""
    public_contact = []

    # LinkedIn Profile Extraction
    if "linkedin.com" in domain and "/in/" in path:
        handle_match = re.search(r"/in/([^/?#]+)", path)
        if handle_match:
            linkedin_profile = f"https://linkedin.com{handle_match.group(1)}"

    # Unified Platform Registry Maps
    social_platforms = {
        "facebook.com": "Facebook", "instagram.com": "Instagram",
        "tiktok.com": "TikTok", "twitter.com": "X/Twitter", "x.com": "X/Twitter",
        "snapchat.com": "Snapchat", "reddit.com": "Reddit"
    }
    
    for s_domain, s_name in social_platforms.items():
        if s_domain in domain:
            handle_match = re.search(r"^/(?:@|people/|user/)?([^/?#]+)", path)
            if handle_match and handle_match.group(1) not in ["pages", "groups", "share", "watch"]:
                user_handle = handle_match.group(1)
                other_socials.append(f"@{user_handle} ({s_name})")
                social_urls_only.append(url)

    clean_text = combined_text
    if employee_name.lower() in clean_text.lower():
        clean_text = re.sub(re.escape(employee_name), '', clean_text, flags=re.IGNORECASE)

    headline_segments = re.split(r'\s*[|\-—••,]\s*', clean_text)
    corporate_titles = ["manager", "analyst", "director", "officer", "lead", "specialist", "engineer", "coordinator", "supervisor", "operator"]
    
    for segment in headline_segments:
        if len(segment.strip()) > 3 and any(kw in segment.lower() for kw in corporate_titles):
            deduced_title = clean_job_title(segment)
            break

    emails = re.findall(r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b", combined_text)
    phones = re.findall(r"\b(?:\+?1[-. ]?)?\(?[0-9]{3}\)?[-. ]?[0-9]{3}[-. ]?[0-9]{4}\b", combined_text)
    
    if emails:
        public_contact.extend(list(set(emails)))
    if phones:
        for p in list(set(phones)):
            digits = re.sub(r"\D", "", p)
            if len(digits) == 10:
                public_contact.append(f"({digits[:3]}) {digits[3:6]}-{digits[6:]}")
            else:
                public_contact.append(p)

    return {
        "LinkedIn Profile": linkedin_profile,
        "Other Social Media Profiles": ", ".join(other_socials) if other_socials else "",
        "Social Media Links Only": " | ".join(social_urls_only) if social_urls_only else "",
        "Deduced Job Title": deduced_title,
        "Public Contact Information": ", ".join(public_contact) if public_contact else ""
    }


def main():
    print(f"[*] Extracting primary identity metadata from: {INPUT_EXCEL_USERS}...")
    if not os.path.exists(INPUT_EXCEL_USERS):
        print(f"[!] Error: The base template database file '{INPUT_EXCEL_USERS}' was not found.")
        return

    try:
        df_base = pd.read_excel(INPUT_EXCEL_USERS, engine="openpyxl")
        dept_col = df_base.columns[3]
        title_col = df_base.columns[4]
        name_col = df_base.columns[1]
    except Exception as e:
        print(f"[!] Failed parsing internal spreadsheet schema: {e}")
        return

    user_registry = {}
    for _, row in df_base.iterrows():
        emp_name = str(row.get(name_col, "")).strip()
        if emp_name and emp_name != "nan":
            user_registry[emp_name] = {
                "Internal Department": str(row.get(dept_col, "")),
                "Internal Job Title": str(row.get(title_col, ""))
            }

    print(f"[*] Accessing raw JSON OSINT dataset input: {INPUT_RAW_JSON}...")
    if not os.path.exists(INPUT_RAW_JSON):
        print(f"[!] Error: Target input collection file '{INPUT_RAW_JSON}' not found.")
        return

    try:
        with open(INPUT_RAW_JSON, "r", encoding="utf-8") as f:
            master_data = json.load(f)
    except Exception as e:
        print(f"[!] Error loading JSON file: {e}")
        return

    dossier_rows = []
    reference_links_rows = []

    print("[*] Launching structured key field extraction pipeline...")
    for employee in master_data:
        name = employee.get("employee_name", "Unknown")
        assets = employee.get("discovered_assets", [])

        internal_profile = user_registry.get(name, {"Internal Department": "", "Internal Job Title": ""})
        residential_address = fetch_residential_address(name)

        deduced_titles = []
        primary_linkedin = ""
        all_socials = []
        all_social_links = []
        all_contacts = []

        for asset in assets:
            url = asset.get("url", "")
            title = asset.get("result_title", "")
            snippet = asset.get("snippet", "")

            intel = dynamic_intel_extraction(url, title, snippet, name)

            if intel.get("LinkedIn Profile") and not primary_linkedin:
                primary_linkedin = intel.get("LinkedIn Profile")
            if intel.get("Other Social Media Profiles"):
                all_socials.extend(intel.get("Other Social Media Profiles").split(", "))
            if intel.get("Social Media Links Only"):
                all_social_links.extend(intel.get("Social Media Links Only").split(" | "))
            if intel.get("Deduced Job Title") != "":
                deduced_titles.append(intel.get("Deduced Job Title"))
            if intel.get("Public Contact Information"):
                all_contacts.extend(intel.get("Public Contact Information").split(", "))

            # Populate a separate sheet list mapping every raw platform asset back to the individual
            if url:
                reference_links_rows.append({
                    "Employee Name": name,
                    "Target Resource Link": url,
                    "Discovery Text Snippet": snippet
                })

        final_deduced_title = deduced_titles[0] if deduced_titles else ""
        final_socials = ", ".join(list(set(all_socials))) if all_socials else ""
        final_social_links = " | ".join(list(set(all_social_links))) if all_social_links else ""
        final_contacts = ", ".join(list(set(all_contacts))) if all_contacts else ""

        dossier_rows.append({
            "Full Name": name,                                                 
            "Internal Department": internal_profile["Internal Department"], 
            "Internal Job Title": internal_profile["Internal Job Title"],   
            "Deduced Job Title (Online)": final_deduced_title,
            "LinkedIn Profile": primary_linkedin if primary_linkedin else "",
            "Social Media Handles": final_socials,
            "Social Media Links Asset": final_social_links,
            "Phone Number / Contact Info": final_contacts,
            "Residential Address": residential_address
        })
        
    df_dossier = pd.DataFrame(dossier_rows)
    df_links = pd.DataFrame(reference_links_rows)
    print(f"[*] Writing multi-sheet workbook report: {OUTPUT_EXCEL_REPORT}")

    try:
        with pd.ExcelWriter(OUTPUT_EXCEL_REPORT, engine="openpyxl") as writer:
            df_dossier.to_excel(writer, index=False, sheet_name="Employee Overview")
            df_links.to_excel(writer, index=False, sheet_name="Social Media Reference Log")
            
            # Apply column widths automatically across tabsfor sheet_name in ["Employee Dossier Overview", "Social Media Reference Log"]:
            worksheet = writer.sheets[sheet_name]
            for col_idx, col in enumerate(worksheet.columns, start=1):
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col_idx)
                worksheet.column_dimensions[col_letter].width = min(max(max_len + 3, 14), 55)
            
            print(f"[++] Targeted corporate dossier report successfully saved to: {OUTPUT_EXCEL_REPORT}")
    
    except Exception as e:
        print(f"[!] File saving exception hit: {e}")

if __name__ == "__main__":
    main()                 
