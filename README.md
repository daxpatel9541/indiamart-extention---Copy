# IndiaMART Lead Extractor

A powerful Chrome extension designed to automate the extraction of lead details from the IndiaMART Seller Panel (Lead Manager). It captures comprehensive business information and exports it directly to Excel or Google Sheets.

## 🚀 Key Features
- **Automated Extraction**: Scans the Lead Manager and clicks through messages automatically.
- **Detailed Data Extraction**: Captures company name, mobile number, address, GST, business type, owner name, ratings, reviews, and more.
- **Excel Export**: Download all extracted leads as a `.xlsx` file with one click.
- **Google Sheets Sync**: Integration provided via Google Apps Script for real-time data logging.
- **Intelligent Selectors**: Uses advanced DOM traversal strategies (Logo match, Header match, etc.) to ensure reliability even if the page layout changes.

## 🔄 Process Flow
The extension follows a systematic flow to ensure high-quality data capture:

1.  **Initiation**: Open the [IndiaMART Lead Manager](https://seller.indiamart.com/messagecentre/) page.
2.  **Activation**: Click the extension icon and press **"Start Extraction"**.
3.  **Lead Discovery**: The `content.js` script identifies all lead entries in the left-hand message panel.
4.  **Sequential Processing**: For each lead entry:
    - **Step A: Selection**: Clicks the lead to load the conversation.
    - **Step B: Mobile Extraction**: Scans the chat header and "tel:" links for the most accurate mobile number.
    - **Step C: Detail Search**: Clicks the **"View Details «"** button to trigger the company profile panel.
    - **Step D: Data Parsing**: Wait for the "Fact Sheet" and "Contact Details" to appear, then parses ~20 different fields.
    - **Step E: Storage**: Saves the lead data to `chrome.storage.local` and marks the item as processed.
    - **Step F: Cleanup**: Closes the detail panel to prepare for the next lead.
5.  **Completion**: Once finished, the user can use the **"Download Excel"** button in the popup to retrieve the consolidated list.

## 🛠️ Installation
1.  Download this project folder.
2.  Open Google Chrome and navigate to `chrome://extensions/`.
3.  Enable **"Developer mode"** in the top right corner.
4.  Click **"Load unpacked"** and select the folder containing these files.

## 📂 File Structure
- `manifest.json`: Extension configuration, permissions, and service worker setup.
- `content.js`: The "engine" that interacts with the IndiaMART page and extracts data.
- `popup.html/js/css`: The user interface and Excel export logic.
- `background.js`: Orchestrates communication between the popup and content scripts.
- `GOOGLE_APPS_SCRIPT.js`: A script template for your Google Sheet to receive lead data via POST requests.
- `libs/xlsx.full.min.js`: Supporting library for client-side Excel generation.

## 📊 Extracted Fields
- **Company Info**: Name, Address, Website, Year Established.
- **Contact Details**: Mobile Number, Contact Person, Owner Name.
- **Business Insights**: Business Type, Annual Turnover, GST No., PAN Number.
- **Trust Metrics**: IndiaMART Member Since, Ratings, Review Count, Response Rate.
- **Products**: List of items the company sells.

---
*Disclaimer: This tool is intended for personal productivity. Please respect IndiaMART's terms of service and use responsibly.*
