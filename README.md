# 🚀 Social Media Lead Generation & Comment Scraper

Automated Python suite built with **Playwright (Async)** to scrape, filter, and structure public comments and lead data from **Facebook Pages** and **Instagram Profiles**.

Designed specifically for market research, audience analysis, and lead generation from high-intent buyers (pricing, buying, and inquiry comments).

---

## ✨ Features

### 📸 Instagram Comment Scraper
- **Automated Profile Navigation:** Reads company URLs from Excel and navigates through target profiles.
- **Dynamic Post Discovery:** Scrolls through pages to locate recent posts and reels automatically.
- **Smart DOM Parsing:** Extracts comment texts alongside original user profile links while filtering out UI/system noise (likes, time labels, UI texts).
- **Auto-Save & Resume:** Saves progress incrementally after each post to avoid data loss.

### 📘 Facebook Lead & Phone Scraper
- **Language Detection:** Specifically filters and extracts Arabic comments and potential buyer inquiries.
- **Dynamic Comment Expansion:** Expands "View More Comments" and nested replies using human-like interaction.
- **Public Profile & Contact Scrape:** Optionally inspects commenter profiles to extract publicly disclosed phone numbers (supports Egyptian formats e.g., `01x`, `+20`).
- **Target-Driven Search:** Smartly stops per-page scan once a target threshold of qualified Arabic-comment posts is reached.

---

## 🛠️ Tech Stack & Requirements

- **Python:** 3.9+
- **Browser Automation:** [Playwright (Async API)](https://playwright.dev/python/)
- **Data Processing:** [Pandas](https://pandas.pydata.org/), [OpenPyXL](https://openpyxl.readthedocs.io/)
- **Regex & Asyncio:** Native Python asynchronous handling and regular expression filtering.

---

## 🚀 Quick Start

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/YOUR_USERNAME/YOUR_REPOSITORY_NAME.git](https://github.com/YOUR_USERNAME/YOUR_REPOSITORY_NAME.git)
   cd YOUR_REPOSITORY_NAME
   Install dependencies:

Bash
pip install playwright pandas openpyxl
playwright install chromium
Prepare Input Data:

Create an input_data.xlsx file.

For Instagram: Provide columns Company Name and Instagram.

For Facebook: Provide columns Brand and Facebook URL.

Run the Scrapers:

Instagram:

Bash
python instagram_scraper.py
Facebook:

Bash
python facebook_scraper.py
🔒 Security & Privacy Notice
These scripts use Playwright Persistent Contexts allowing manual initial login and bypass of anti-bot checks. No passwords or authentication cookies are stored in the public source code. All output datasets strip irrelevant metadata to strictly respect target platform terms and public data boundaries.
