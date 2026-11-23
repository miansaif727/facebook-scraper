# Facebook Page About Scraper 🚀

**A Python-based Facebook Page Information Scraper powered by Selenium and BeautifulSoup.**

Extract publicly available Facebook Page information automatically, including page names, followers, following, business addresses, phone numbers, email addresses, and website links.

This tool is designed to process multiple Facebook pages in parallel using independent Chrome browser sessions, with automatic progress saving and error handling to make large scraping jobs easier to manage.

---

## ✨ Key Features

- **Facebook Page Information Extraction:** Collect important business and page details from Facebook About sections.
- **Multiple URL Formats:** Supports direct Facebook URLs and Markdown-formatted links.
- **Followers & Following:** Extract publicly displayed follower and following counts.
- **Contact Information:** Collect available business addresses, phone numbers, and email addresses.
- **Website Extraction:** Automatically identifies website links and unwraps Facebook redirect URLs to retrieve the original destination.
- **Multi-Threaded Scraping:** Run multiple Chrome browser instances simultaneously to process pages in parallel.
- **Automatic Progress Saving:** Save results to Excel after every 100 completed links (configurable).
- **Error Handling:** Handle page-load timeouts, missing information, invalid URLs, and other unexpected errors.
- **Optional Facebook Login:** Supports an optional login configuration for pages where additional publicly accessible information requires authentication.
- **Headless Mode:** Run Chrome in the background without displaying browser windows.
- **Randomized Delays:** Introduce configurable delays between page requests.
- **Excel Export:** Export all collected information into a structured `.xlsx` spreadsheet.

---

## 📊 Information Extracted

The scraper collects the following information:

| Field | Description |
|---|---|
| Page URL | Original Facebook Page URL |
| Name | Facebook Page name |
| Followers | Number of followers |
| Following | Number of accounts followed |
| Address | Business or page address |
| Phone Number | Available contact number |
| Email | Publicly displayed email address |
| Website Link | External website associated with the page |
| Status | Scraping result or error status |

**Note:** Information availability depends on each Facebook Page's public visibility, page layout, and the information provided by its owner. Empty fields do not necessarily indicate an error.

---

## 🛠️ Requirements

Before running the scraper, make sure you have:

- Python 3.9 or newer
- Google Chrome installed
- An active internet connection
- Selenium 4.6 or newer

### Install Required Libraries

Open your terminal or command prompt and execute:

```bash
pip install selenium beautifulsoup4 openpyxl pandas
```

### ChromeDriver

No manual ChromeDriver download is required.

Selenium 4.6+ includes Selenium Manager, which can automatically discover and manage a compatible ChromeDriver installation.

---

## 📁 Project Structure

Organize your project files as follows:

```text
Facebook-Page-About-Scraper/
│
├── fb_scraper.py
├── links.txt
├── output.xlsx
├── README.md
└── requirements.txt
```

The `output.xlsx` file will be generated automatically when scraping starts.

### requirements.txt

```text
selenium
beautifulsoup4
openpyxl
pandas
```

Install everything using:

```bash
pip install -r requirements.txt
```

---

## 🔗 Input File Configuration

Create a file named `links.txt` in the same directory as the Python script.

Add Facebook Page URLs, with one URL per line.

Example:

```text
https://www.facebook.com/yourpage1
https://www.facebook.com/yourpage2
https://www.facebook.com/yourpage3
https://www.facebook.com/yourpage4
```

### Markdown Link Support

The scraper also supports Markdown-style links.

Example:

```text
[Facebook Page](https://www.facebook.com/yourpage/about)
```

The script automatically extracts the actual URL and navigates to the appropriate About page.

---

## ⚙️ Configuration

Open `fb_scraper.py` and adjust the configuration variables according to your requirements.

### 1. Input and Output Files

```python
INPUT_FILE = "links.txt"
OUTPUT_FILE = "output.xlsx"
```

- `INPUT_FILE`: File containing Facebook Page URLs.
- `OUTPUT_FILE`: Excel file where extracted information will be stored.

### 2. Facebook Login (Optional)

```python
LOGIN = False

FB_EMAIL = os.environ.get("FB_EMAIL", "")
FB_PASSWORD = os.environ.get("FB_PASSWORD", "")
```

By default, Facebook login is disabled.

If authentication is required for your permitted use case, set:

```python
LOGIN = True
```

Set your credentials through environment variables rather than hardcoding them into the script.

**Windows PowerShell:**

```powershell
$env:FB_EMAIL="your_email@example.com"
$env:FB_PASSWORD="your_password"
```

Only use an account you own or are authorized to use. Login does not guarantee access to restricted information.

### 3. Browser Configuration

```python
HEADLESS = True
```

- `True`: Run Chrome in the background.
- `False`: Open a visible Chrome browser window.

### 4. Page Loading Settings

```python
PAGE_LOAD_WAIT = 5
PAGE_LOAD_TIMEOUT = 25
```

- `PAGE_LOAD_WAIT`: Additional waiting time for JavaScript-rendered content.
- `PAGE_LOAD_TIMEOUT`: Maximum page navigation time before marking a page as timed out.

### 5. Request Delays

```python
MIN_DELAY, MAX_DELAY = 4, 9
```

The scraper introduces a randomized delay between pages handled by each worker.

### 6. Parallel Browser Sessions

```python
MAX_PARALLEL_ALLOWED = 10
```

This defines the maximum number of browser workers offered when the script asks how many links should run simultaneously.

You can choose the number of workers when starting the scraper.

**Tip:** More browsers require more RAM and CPU resources. Start with 1–3 workers and increase gradually if your system can handle it.

### 7. Automatic Saving

```python
AUTOSAVE_EVERY = 100
```

The scraper automatically saves progress after every 100 completed links.

You can change this value according to your requirements.

---

## ▶️ How to Run

Once everything is configured, open your project directory in the terminal.

Run:

```bash
python fb_scraper.py
```

The program will:

1. Check whether the input file exists.
2. Load all Facebook Page URLs.
3. Ask how many browser sessions you want to run simultaneously.
4. Start the Chrome browser workers.
5. Open each Facebook Page's About section.
6. Extract available page and contact information.
7. Display scraping progress in the terminal.
8. Automatically save results to Excel.
9. Generate the final Excel spreadsheet.

### Example Terminal Interaction

```text
======================================================================
 FACEBOOK PAGE ABOUT SCRAPER
======================================================================
 Input file  : links.txt
 Output file : output.xlsx
 Headless    : True
 Login       : False
 Autosave    : every 100 links
======================================================================

Loaded 500 link(s) from links.txt

How many links do you want to run in parallel at a time?
(1-10, default 1): 3

Running with 3 browser(s) in parallel.
Starting browser(s)...
```

### Example Progress Output

```text
[W1|1/500] Opening : https://www.facebook.com/page1/about
[W2|2/500] Opening : https://www.facebook.com/page2/about
[W3|3/500] Opening : https://www.facebook.com/page3/about

[W1|1/500] Done : Example Business [6.2s]
[W1|1/500]   Followers : 12.5K   Following: 120
[W1|1/500]   Address   : New York, USA
[W1|1/500]   Phone     : +1 123 456 7890
[W1|1/500]   Email     : contact@example.com
[W1|1/500]   Website   : https://example.com

[Autosave] 100/500 links done -- progress saved to output.xlsx
```

---

## 📂 Output Excel File

All extracted information is stored in:

```text
output.xlsx
```

The spreadsheet contains nine columns:

| Page URL | Name | Followers | Following | Address | Phone Number | Email | Website Link | Status |
|---|---|---|---|---|---|---|---|---|
| Facebook URL | Example Business | 12.5K | 120 | New York | +11234567890 | contact@example.com | example.com | OK |

### Status Values

- `OK`: Page name successfully extracted.
- `No data found`: Page loaded, but no page name was extracted.
- `Invalid URL`: Unable to identify a valid URL.
- `Timed out (likely deleted/unavailable)`: Page navigation exceeded the configured timeout.
- `Error`: An unexpected scraping error occurred.

---

## 🧠 How It Works

The project uses several Python libraries to handle browser automation, HTML parsing, and data management.

| Library | Purpose |
|---|---|
| Selenium | Opens Facebook pages and handles browser automation |
| BeautifulSoup | Parses HTML and extracts page information |
| Pandas | Organizes extracted information into structured data |
| OpenPyXL | Writes results into Excel spreadsheets |
| ThreadPoolExecutor | Manages multiple browser workers |
| Threading | Protects shared progress and file-saving operations |
| Regular Expressions | Identifies emails, phone numbers, and follower counts |

### Scraping Workflow

```text
            links.txt
                |
                v
       Load Facebook URLs
                |
                v
       Ask Parallel Workers
                |
                v
       Launch Chrome Sessions
                |
                v
      Open Facebook About Pages
                |
                v
        Extract Page Details
                |
                v
        Parse HTML with BS4
                |
                v
       Store Results in Excel
                |
                v
        Automatic Progress Save
                |
                v
          Final output.xlsx
```

---

## ⚡ Performance & Optimization

The scraper supports multiple browser sessions to improve processing efficiency.

Some recommended configurations:

| System | Suggested Workers |
|---|---:|
| 4 GB RAM | 1 |
| 8 GB RAM | 1–2 |
| 16 GB RAM | 2–4 |
| 32 GB RAM | 4–6 |

These are starting suggestions, not guaranteed limits. Actual performance depends on CPU, network speed, Chrome memory usage, and page complexity.

**Important:** Running more browser sessions does not always mean faster scraping. Excessive concurrency may cause resource issues or trigger platform restrictions.

---

## 🔧 Troubleshooting

### 1. ChromeDriver Error

Make sure Google Chrome is installed and updated.

Try updating Selenium:

```bash
pip install --upgrade selenium
```

### 2. Input File Not Found

Make sure `links.txt` is in the same directory as `fb_scraper.py`, or update `INPUT_FILE` with its full path.

### 3. No Data Found

Possible reasons:

- The Facebook Page is unavailable.
- The page has restricted visibility.
- Facebook changed its HTML structure.
- The information is not publicly displayed.
- The page has not finished rendering.

### 4. Pages Timing Out

Increase the timeout:

```python
PAGE_LOAD_TIMEOUT = 40
PAGE_LOAD_WAIT = 8
```

### 5. High RAM Usage

Reduce the number of parallel workers.

For example:

```text
How many links do you want to run in parallel?
(1-10, default 1): 2
```

### 6. Excel File Not Saving

Make sure `output.xlsx` is not open in Microsoft Excel while the scraper is running and that the project directory is writable.

---

## ⚠️ Important Notes

- This tool is intended for extracting information that is publicly available or that you are otherwise authorized to access.
- It does not bypass Facebook authentication, privacy settings, CAPTCHAs, or access restrictions.
- Facebook may change its page structure, which can require updates to the extraction logic.
- Extracted information may be incomplete or outdated.
- Use reasonable request rates and comply with Facebook's applicable terms, policies, and privacy laws.
- Do not use this tool for spam, unauthorized profiling, or other abusive activities.

---

## 👨‍💻 Author

**Mian Saif**

Python Developer | Web Scraping Specialist | Lead Generation & Automation

Specializing in:

- Web Scraping & Data Extraction
- Python Automation
- Selenium & BeautifulSoup
- Lead Generation
- Data Processing & Excel Automation
- Custom Scraping Solutions

---

## ⭐ Support the Project

If you find this project useful:

- Give the repository a ⭐ Star.
- Fork the repository.
- Report bugs or suggest improvements through Issues.
- Share the project with other developers.

**Built with Python, Selenium, and a passion for automation.**
