# Inside SEPTA: Reconnaissance from the Outside In

This tool automates the process of gathering publicly available information about SEPTA employees from online. It searches across Google (via the Serper API), scrapes the pages it finds, uses a local AI model (Llama 3.2) to make sense of what it found, and then neatly organizes everything into an Excel spreadsheet and an interactive web dashboard.

---

## Step 1: Collecting the data (`collect.py`)

This is the main script and the slowest one. For each employee, it fires off **5 different Google searches** and collects up to 60 results total:

| Search query | What it's looking for |
|---|---|
| `"Name" SEPTA` | General results linking the person to SEPTA |
| `"Name" "@septa.org"` | Any page where their official SEPTA email was exposed |
| `intitle:"Name" SEPTA` | Pages that are specifically *about* this person |
| `"Name" site:linkedin.com` | Their LinkedIn profile |
| `"Name" site:facebook.com` | Their Facebook profile |

After searching, it visits each page, reads its content, and asks the local AI model running on Ollama (Llama 3.2) to identify anything useful (contact info, social media accounts, job titles, etc.). Expect roughly 10-30 seconds per employee, depending on how fast Ollama is running on your machine.

---

## Step 2: Building the Excel spreadsheet (`extract.py`)

Once `collect.py` has finished (or you've collected enough data for now), run `extract.py` to turn the raw JSON data into a clean, organized Excel spreadsheet.

This produces `inside_septa.xlsx` with two sheets:
- **Inside SEPTA** — one row per employee, with all their discovered info
- **Source Links Log** — every URL that was found, with titles and snippets

---

## Step 3: Refreshing the dashboard (`excel_to_json.py`)

The dashboard can't read Excel spreadsheets directly, so this script converts `inside_septa.xlsx` into a JSON format the browser can seamlessly use. It only takes a few seconds to run. You need to run this every time after you run `extract.py`, otherwise the dashboard will still be showing old, stale data.

---

## Step 4: Viewing the dashboard

Start a local web server so you can open the dashboard in your browser:

```powershell
python -m http.server 8765
```

Then go to: **http://localhost:8765/dashboard.html**

Leave that terminal window open while you're using the dashboard.

### Dashboard UI Structure

- **Overview** - high-level stats and charts: how many employees have social profiles, phone numbers, emails, and residential addresses on record
- **Employees** - the full employee list, searchable and filterable, with a detailed profile panel for each person
- **Social Media** - browse by platform (e.g., LinkedIn, Facebook, Instagram, etc.) to see who has a profile on each one
- **Contacts** - all discovered phone numbers and email addresses in one central place
- **Address Map** - an interactive map showing residential addresses pinned to Philadelphia
- **Source Links** - every URL that was discovered, with the title and snippet from the page

---

## Run Steps

```powershell
# 1. Open your terminal and activate the virtual environment
python -m venv venv
.\venv\Scripts\activate
pip install -r requirements.txt

# 2. Make sure Ollama is running

# 3. Run the collection for your target range
python collect.py --start {start_pos} --end {end_pos}
# "start_pos" and "end_pos" used to specify a range to query the rows in the spreadsheet `inside_septa.xlsx`

# 4. Once it's done, build the Excel spreadsheet
python extract.py

# 5. Refresh the dashboard data
python excel_to_json.py

# 6. View the dashboard
python -m http.server 8765
# Then open http://localhost:8765/dashboard.html in your browser
```

---

## NOTES

**Running the Llama 3.2 is optional but valuable.** If Llama 3.2 isn't running by Ollama, `collect.py` still works; it just skips augumenting the number of results with AI. The pages will still be saved, but fields like emails, phone numbers, and social handles won't be automatically extracted.

**Removing common name ambiguity.** For employees with very common names (e.g., "John Smith"), some results may not be about the right person. The AI is designed to flag these as irrelevant or uncertain (but it's not fully perfected), and `extract.py` filters them accordingly.

---

## Configuration

If you ever need to change the default settings, they're all at the top of `collect.py`:

| Setting | What it controls |
|---|---|
| `SERPER_API_KEY` | Your Serper API key for Google search (currently linked to Agent Cookie's account) |
| `INPUT_EXCEL` | Path to the source employee spreadsheet |
| `START_INDEX` / `END_INDEX` | Default row range when no `--start`/`--end` is passed |
| `OLLAMA_MODEL` | The AI model Ollama uses (default: `llama3.2`) |
| `SCRAPE_WORKERS` | How many pages are fetched at the same time (default: 12) |
| `ENABLE_LLM` | Set to `False` to skip enriching results with AI entirely |