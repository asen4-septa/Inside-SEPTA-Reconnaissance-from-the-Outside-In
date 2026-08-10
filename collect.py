AGENTC_SERPER_API_KEY = "c213ee9acbb70eb39a32c3da91ff8acffb8778e7"
ASEN_SEPTA_SERPER_API_KEY = "b57ded568e853310d963f81985d90c508a86d987"
APS7140_SERPER_API_LEY = "7758de7844f957cba53fd28049598694b60a64cf"

"""
collect.py — SEPTA OSINT Collection Engine (v3)
=================================================
Changes from v2:
  (1) LLM prompt expanded — now extracts social_platform + social_handle
      per asset in addition to job_title, employer, location, education,
      email_addresses, phone_numbers.
  (2) URL passed to llm_enrich_asset() so the model can identify the
      platform from the URL itself (e.g. instagram.com → "instagram").
  (3) social_handles field removed from v2 (it was returning raw strings
      with no platform attribution and was never used by extract.py).
      Replaced by social_platform + social_handle as a pair per asset.

Architecture:
  collect.py  →  ALL fetching + scraping + LLM extraction  →  JSON
  extract.py  →  aggregate JSON fields + OPA lookup + Excel  →  XLSX
"""

import os
import json
import time
import requests
import argparse
import pandas as pd
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urlparse

try:
    import trafilatura
    HAS_TRAFILATURA = True
except ImportError:
    HAS_TRAFILATURA = False
    print("[!] trafilatura not found — install with: pip install trafilatura")

try:
    from bs4 import BeautifulSoup
    HAS_BS4 = True
except ImportError:
    HAS_BS4 = False

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------

SERPER_API_KEY      = "c213ee9acbb70eb39a32c3da91ff8acffb8778e7"
INPUT_EXCEL         = "assets/exportUsers_2026-7-13.xlsx"
NAME_COLUMN         = "displayName"
OUTPUT_OSINT_JSON   = "septa_social_links.json"
OUTPUT_METADATA_CSV = "document_metadata.csv"
DOWNLOAD_DIR        = "./downloaded_docs"

# Ollama settings
OLLAMA_MODEL = "llama3.2"
OLLAMA_BASE  = "http://localhost:11434"

# Scraping settings
SCRAPE_MAX_CHARS = 8000   # Hard cap on stored full_content length
SCRAPE_WORKERS   = 12     # Parallel threads for URL fetching per employee
SCRAPE_TIMEOUT   = 12     # Seconds before giving up on a URL fetch

# Batch processing settings (1-indexed based on the list of unique employees)
# This allows you to process the spreadsheet in smaller chunks.
# Examples: 1 to 500, 501 to 1000. Set END_INDEX = None to process to the end.
# You can also override these via command line: python collect.py --start 501 --end 1000
START_INDEX = 1
END_INDEX = 500

# Set False to skip LLM entirely (scraping still runs)
ENABLE_LLM = True

# ---------------------------------------------------------------------------
# GATED / BLOCKED DOMAINS  (skip fetching — saves time, avoids bans)
# ---------------------------------------------------------------------------

GATED_DOMAINS = {
    "linkedin.com",
    "rocketreach.co",
    "contactout.com",
    "zoominfo.com",
    "theorg.com",
    "radaris.com",
    "whitepages.com",
    "spokeo.com",
    "beenverified.com",
    "intelius.com",
    "peoplefinder.com",
    "peoplesmart.com",
    "instantcheckmate.com",
    "truthfinder.com",
    "familytreenow.com",
    "fastpeoplesearch.com",
    "anywho.com",
    "411.com",
    "checkr.com",
    "pipl.com",
    "peekyou.com",
    "thatsthem.com",
    "voterrecords.com",
}

os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# ---------------------------------------------------------------------------
# GATED DOMAIN CHECK
# ---------------------------------------------------------------------------

def _is_gated(url):
    """Returns True if the URL belongs to a known gated/blocked domain."""
    try:
        domain = urlparse(url).netloc.lower().replace("www.", "")
        return any(gated in domain for gated in GATED_DOMAINS)
    except Exception:
        return False

# ---------------------------------------------------------------------------
# FULL PAGE SCRAPING
# ---------------------------------------------------------------------------

_SCRAPE_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
}


def scrape_full_content(url):
    """
    Fetches and extracts the main body text from a URL.
    Returns (content_text: str, was_scraped: bool).
    """
    if not url:
        return ("", False)
    if _is_gated(url):
        return ("", False)
    if url.lower().split("?")[0].endswith((".pdf", ".docx", ".xlsx", ".pptx", ".doc", ".xls")):
        return ("", False)

    # trafilatura (primary)
    if HAS_TRAFILATURA:
        try:
            downloaded = trafilatura.fetch_url(url)
            if downloaded:
                text = trafilatura.extract(
                    downloaded,
                    include_comments=False,
                    include_tables=True,
                    no_fallback=False,
                    favor_precision=False,
                )
                if text and len(text.strip()) > 120:
                    return (text.strip()[:SCRAPE_MAX_CHARS], True)
        except Exception:
            pass

    # BeautifulSoup fallback
    if HAS_BS4:
        try:
            resp = requests.get(
                url,
                timeout=SCRAPE_TIMEOUT,
                headers=_SCRAPE_HEADERS,
                allow_redirects=True,
            )
            if resp.status_code == 200 and "text/html" in resp.headers.get("Content-Type", ""):
                soup = BeautifulSoup(resp.text, "html.parser")
                for tag in soup(["nav", "footer", "header", "script", "style", "aside", "form"]):
                    tag.decompose()
                paragraphs = soup.find_all(["p", "article", "main", "section", "li"])
                text = " ".join(p.get_text(separator=" ", strip=True) for p in paragraphs)
                text = " ".join(text.split())
                if len(text) > 120:
                    return (text[:SCRAPE_MAX_CHARS], True)
        except Exception:
            pass

    return ("", False)




# ---------------------------------------------------------------------------
# OLLAMA  (local LLM enrichment)
# ---------------------------------------------------------------------------

_OLLAMA_STATUS = None   # None = unchecked, True = up, False = down


def ollama_available():
    """
    Checks if Ollama is reachable at localhost:11434.
    Result is cached after the first call.
    """
    global _OLLAMA_STATUS
    if _OLLAMA_STATUS is not None:
        return _OLLAMA_STATUS
    try:
        r = requests.get(f"{OLLAMA_BASE}/api/tags", timeout=4)
        _OLLAMA_STATUS = (r.status_code == 200)
    except Exception:
        _OLLAMA_STATUS = False
    if not _OLLAMA_STATUS:
        print(
            "[!] Ollama not reachable — LLM features disabled.\n"
            "    To enable: run 'ollama serve' or open the Ollama desktop app.\n"
            "    Scraping will still run; relevance/entities will be left empty."
        )
    return _OLLAMA_STATUS


def llm_enrich_asset(name, title, url, content, snippet):
    """
    Single Ollama call per asset. Returns ALL structured fields:
      - relevance score
      - job title, employer, location, education
      - email addresses, phone numbers
      - social platform + handle (identified from the URL and page content)

    Social platform/handle is identified here because the LLM has the URL,
    title, and page text — it can reliably determine "this is an Instagram
    profile for user X" without separate URL-parsing logic in extract.py.

    Uses format:"json" (JSON mode) — guarantees parseable output.

    Returns:
      {
        "relevance": "relevant" | "uncertain" | "irrelevant",
        "entities": {
          "job_title":       str,
          "employer":        str,
          "location":        str,
          "education":       str,
          "email_addresses": [str, ...],
          "phone_numbers":   [str, ...],
          "social_accounts": [
            {"platform": "linkedin"|"facebook"|"instagram"|"tiktok"|
                         "twitter"|"github"|"youtube"|"reddit"|"other",
             "handle": str},
            ...
          ]
        }
      }
    A single page can reference multiple social platforms (e.g. a personal
    website listing LinkedIn, Instagram, and Twitter handles), so
    social_accounts is always a list — even if only one platform is found.
    """
    if not ENABLE_LLM or not ollama_available():
        return {"relevance": "uncertain", "entities": {}}

    analysis_text = content if content else snippet
    if not analysis_text:
        return {"relevance": "uncertain", "entities": {}}

    text_sample = analysis_text[:3000]

    prompt = (
        f'You are an OSINT analyst. Analyze this search result about "{name}", '
        f"a SEPTA (Southeastern Pennsylvania Transportation Authority) employee.\n\n"
        f"URL: {url}\n"
        f"Page title: {title}\n"
        f"Page text: {text_sample}\n\n"
        f"Return ONLY a JSON object with these exact keys:\n"
        f'  "relevance": Is this page specifically about {name} at SEPTA? '
        f'Use "relevant", "uncertain", or "irrelevant".\n'
        f'  "job_title": Their job title as a string, or "".\n'
        f'  "employer": Their employer name as a string, or "".\n'
        f'  "location": Their city and state as a string, or "".\n'
        f'  "education": Their school or university name as a string, or "".\n'
        f'  "email_addresses": Array of personal email strings found on this page, or [].\'\n'
        f'  "phone_numbers": Array of phone number strings found on this page, or [].\n'
        f'  "social_accounts": Array of social media accounts referenced on this page. '
        f'Each entry is an object with "platform" (one of: linkedin, facebook, instagram, '
        f'tiktok, twitter, github, youtube, reddit, other) and "handle" (the username). '
        f'Include ALL platforms you find — a page may reference several. Use [] if none.\n'
        f"No explanations. Return the JSON object only."
    )

    try:
        resp = requests.post(
            f"{OLLAMA_BASE}/api/generate",
            json={
                "model":  OLLAMA_MODEL,
                "prompt": prompt,
                "format": "json",
                "stream": False,
                "options": {
                    "temperature": 0,
                    "num_predict": 400,
                },
            },
            timeout=90,
        )

        if resp.status_code == 200:
            raw    = resp.json().get("response", "{}")
            parsed = json.loads(raw)

            relevance = str(parsed.get("relevance", "uncertain")).lower()
            if relevance not in ("relevant", "uncertain", "irrelevant"):
                relevance = "uncertain"

            VALID_PLATFORMS = {
                "linkedin", "facebook", "instagram", "tiktok",
                "twitter",  "github",   "youtube",   "reddit", "other"
            }
            raw_accounts = parsed.get("social_accounts", [])
            social_accounts = []
            if isinstance(raw_accounts, list):
                for acct in raw_accounts:
                    if isinstance(acct, dict):
                        p = str(acct.get("platform", "")).lower().strip()
                        h = str(acct.get("handle",   "")).strip()
                        if p in VALID_PLATFORMS and h:
                            social_accounts.append({"platform": p, "handle": h})

            entities = {
                "job_title": str(parsed.get("job_title", "")).strip(),
                "employer":  str(parsed.get("employer",  "")).strip(),
                "location":  str(parsed.get("location",  "")).strip(),
                "education": str(parsed.get("education", "")).strip(),
                "email_addresses": [
                    str(e).strip() for e in parsed.get("email_addresses", [])
                    if e and isinstance(e, str)
                ],
                "phone_numbers": [
                    str(p).strip() for p in parsed.get("phone_numbers", [])
                    if p and isinstance(p, str)
                ],
                "social_accounts": social_accounts,
            }

            return {"relevance": relevance, "entities": entities}

    except Exception:
        pass

    return {"relevance": "uncertain", "entities": {}}

# ---------------------------------------------------------------------------
# SERPER SEARCH + SCRAPE + LLM
# ---------------------------------------------------------------------------

def search_employee(name):
    """
    Full pipeline for one employee:
      1. Fire 5 targeted Serper queries (12 results each = 60 total).
         Queries used:
           Q1  "{name}" SEPTA                  — primary anchor
           Q2  "{name}" "@septa.org"            — official email exposure
           Q3  intitle:"{name}" SEPTA           — name must be in page title
           Q4  "{name}" site:linkedin.com       — LinkedIn profile
           Q5  "{name}" site:facebook.com       — Facebook profile
      2. Deduplicate results by URL across all 5 queries.
      3. Scrape all non-gated URLs in parallel (ThreadPoolExecutor).
      4. Run LLM enrichment on each asset.
      5. Return enriched employee_record + list of document URLs.
    """
    SERPER_URL  = "https://google.serper.dev/search"
    NUM_RESULTS = 12   # per query  (5 × 12 = 60 total)

    QUERIES = [
        (f'"{name}" SEPTA',              "primary"),
        (f'"{name}" "@septa.org"',       "email_exposure"),
        (f'intitle:"{name}" SEPTA',      "intitle"),
        (f'"{name}" site:linkedin.com',  "linkedin"),
        (f'"{name}" site:facebook.com',  "facebook"),
    ]

    headers = {
        "X-API-KEY":    SERPER_API_KEY,
        "Content-Type": "application/json",
    }

    employee_record = {
        "employee_name":         name,
        "total_links_uncovered": 0,
        "discovered_assets":     [],
    }
    file_urls  = []
    raw_assets = []
    seen_urls  = set()   # deduplication tracker

    # --- Step 1: Fire all Serper queries ---
    for query_str, query_label in QUERIES:
        payload = {"q": query_str, "num": NUM_RESULTS}
        try:
            response = requests.post(
                SERPER_URL, headers=headers, json=payload, timeout=15
            )
            if response.status_code != 200:
                print(f"\n [!] Serper error {response.status_code} "
                      f"(query={query_label}) for: {name}")
                continue

            for item in response.json().get("organic", []):
                link = item.get("link", "")
                if not link or link in seen_urls:
                    continue           # skip empty or duplicate URLs
                seen_urls.add(link)

                is_doc = link.lower().split("?")[0].endswith(
                    (".pdf", ".docx", ".xlsx", ".pptx", ".doc", ".xls")
                )
                if is_doc:
                    file_urls.append(link)

                raw_assets.append({
                    "result_title":    item.get("title",   ""),
                    "url":             link,
                    "snippet":         item.get("snippet", ""),
                    "query_source":    query_label,   # tracks which query found this
                    "is_doc":          is_doc,
                    "full_content":    "",
                    "content_scraped": False,
                    "relevance":       "uncertain",
                    "entities":        {},
                })

        except Exception as exc:
            print(f"\n [!] Network exception (query={query_label}) for {name}: {exc}")

    # --- Step 2: Parallel scraping ---
    scrapeable_indices = [
        i for i, a in enumerate(raw_assets)
        if a["url"] and not a["is_doc"]
    ]

    if scrapeable_indices:
        with ThreadPoolExecutor(max_workers=SCRAPE_WORKERS) as executor:
            future_to_idx = {
                executor.submit(scrape_full_content, raw_assets[i]["url"]): i
                for i in scrapeable_indices
            }
            for future in as_completed(future_to_idx):
                idx = future_to_idx[future]
                try:
                    content, scraped = future.result()
                    raw_assets[idx]["full_content"]    = content
                    raw_assets[idx]["content_scraped"] = scraped
                except Exception:
                    pass

    # --- Step 3: LLM enrichment (sequential — Ollama is single-threaded) ---
    for asset in raw_assets:
        result = llm_enrich_asset(
            name    = name,
            title   = asset["result_title"],
            url     = asset["url"],
            content = asset["full_content"],
            snippet = asset["snippet"],
        )
        asset["relevance"] = result["relevance"]
        asset["entities"]  = result["entities"]

    employee_record["discovered_assets"]     = raw_assets
    employee_record["total_links_uncovered"] = len(raw_assets)

    return employee_record, file_urls


# ---------------------------------------------------------------------------
# FILE DOWNLOAD
# ---------------------------------------------------------------------------

def download_file(url):
    """Downloads target files discovered during the search phase."""
    try:
        filename = os.path.join(DOWNLOAD_DIR, url.split("/")[-1].split("?")[0])
        response = requests.get(url, timeout=20)
        if response.status_code == 200:
            with open(filename, "wb") as f:
                f.write(response.content)
    except Exception:
        pass

# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="SEPTA OSINT Collection Engine")
    parser.add_argument("--start", type=int, default=START_INDEX, help="Starting index (1-indexed)")
    parser.add_argument("--end", type=int, default=END_INDEX, help="Ending index (1-indexed, or None)")
    args, _ = parser.parse_known_args()

    start_val = args.start
    end_val = args.end

    print(f"Accessing spreadsheet file input: {INPUT_EXCEL}...")
    try:
        df = pd.read_excel(INPUT_EXCEL, engine="openpyxl")
    except Exception as exc:
        print(f"[!] Error loading Excel file: {exc}")
        return

    employee_names = [str(n).title() for n in df[NAME_COLUMN].dropna().unique().tolist()]

    print(f"Total file records: {len(employee_names)} names.")
    
    start_idx = max(0, start_val - 1) if start_val is not None else 0
    end_idx = end_val if end_val is not None else len(employee_names)
    
    employee_names = employee_names[start_idx:end_idx]
    print(f"[*] BATCH MODE ACTIVE: Processing {len(employee_names)} targets (from index {start_idx + 1} to {end_idx}).")

    if ENABLE_LLM:
        ollama_available()

    master_json_output = []
    processed_names    = set()
    all_file_urls      = []

    # Resume capability: Load existing data if it exists
    if os.path.exists(OUTPUT_OSINT_JSON):
        try:
            with open(OUTPUT_OSINT_JSON, "r", encoding="utf-8") as f:
                master_json_output = json.load(f)
            for record in master_json_output:
                if "employee_name" in record:
                    # Title-case to match the new format so resume works with older uppercase records
                    processed_names.add(str(record["employee_name"]).title())
            print(f"[*] Resuming from existing JSON. Skipping {len(processed_names)} previously processed employees.")
        except Exception as e:
            print(f"[!] Warning: Could not load existing JSON file. Starting fresh. Error: {e}")
            master_json_output = []
            processed_names = set()

    for name in tqdm(employee_names, desc="Collecting OSINT"):
        if name in processed_names:
            continue
            
        employee_record, file_data = search_employee(name)

        # Append ALL employees to memory so the script remembers they were processed, 
        # even if they had 0 search results. This fixes the amnesia bug.
        master_json_output.append(employee_record)

        all_file_urls.extend(file_data)
        processed_names.add(name)

        # Autosave after every employee so no data is lost on a crash or Ctrl+C
        try:
            with open(OUTPUT_OSINT_JSON, "w", encoding="utf-8") as f:
                json.dump(master_json_output, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"\n [!] Warning: Could not autosave JSON. Is the file open in another program? Error: {e}")

    print(f"\nTotal employee profiles with records: {len(master_json_output)}")
    print(f"Total document URLs isolated:         {len(all_file_urls)}")

    if all_file_urls:
        print("Downloading discovered documents for ExifTool processing...")
        for url in tqdm(set(all_file_urls), desc="Downloading Assets"):
            download_file(url)

        print("Running ExifTool forensic metadata extraction...")
        os.system(f"exiftool -csv -r {DOWNLOAD_DIR} > {OUTPUT_METADATA_CSV}")
        print(f"ExifTool report compiled: {OUTPUT_METADATA_CSV}")
    else:
        print("No document files found. Skipping ExifTool phase.")

    if master_json_output:
        print(f"Master JSON fully saved/autosaved to: {OUTPUT_OSINT_JSON}")


if __name__ == "__main__":
    main()