# 🤖 ChatGPT Excel Automation (Agent Mode)

A fully automated **Python + Playwright** workflow that connects to a running Chrome session and controls **ChatGPT’s Agent Mode** to upload Excel files, issue prompts, and automatically download the AI-generated results — all without any manual clicks.

---

## 🎥 Live Demo
*(Insert your demo video link or thumbnail here)*

---

## 🚀 Overview

This script enables **programmatic control of ChatGPT through Chrome’s DevTools Protocol (CDP)**.  
It simulates a complete human interaction flow with ChatGPT’s **Agent Mode**, including:

1. Connecting to an existing Chrome session  
2. Finding an open ChatGPT tab  
3. Uploading an Excel file  
4. Sending a structured prompt for analysis  
5. Waiting for AI to process  
6. Detecting and downloading the generated spreadsheet  

It is designed for **reliability, transparency, and async performance**, with fallback selectors and retry logic to handle dynamic UI updates.

---

## 🧰 Features

- ✅ **Asynchronous automation** using `asyncio` + Playwright  
- 🔄 **Reuses existing Chrome session** via CDP (no relogin needed)  
- 📎 **Uploads Excel files** into ChatGPT Agent Mode  
- 💬 **Sends structured prompts automatically**  
- ⏳ **Waits and verifies output Excel files** before downloading  
- 🧩 **Robust multi-selector fallback system**  
- 🪶 **Step-by-step logging for debugging and transparency**  
- ⚡ **Extensible for multi-file or API-driven workflows**

---

## 🧑‍💻 Requirements

### 1. Install Dependencies
```bash
pip install playwright asyncio
playwright install chromium
```

### 2. Start Chrome with Remote Debugging
```bash
/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --remote-debugging-port=9222
```

Keep this Chrome window open and make sure ChatGPT (chat.openai.com or chatgpt.com) is loaded and logged in.

### 3. Run the Script

Place your Excel file in the same directory as the script (e.g., `test.xlsx`), then run:
```bash
python automate_chatgpt_excel.py
```

You'll see step-by-step console logs as the bot:
* Connects to Chrome
* Finds your ChatGPT tab
* Uploads the file
* Sends your prompt
* Waits for processing
* Clicks the download button automatically

---

## 📁 Output

The processed Excel file will appear in your system Downloads folder once ChatGPT completes the task.

---

## ⚙️ Customization

You can modify:
* `excel_file` → filename of your input Excel
* `prompt` → the query or instruction sent to ChatGPT

Example:

```python
excel_file = "sales_data.xlsx"
prompt = "Summarize each sheet and create a dashboard summary tab."
```

---

## 🧩 Future Roadmap

* Multi-file batch processing
* Integration with Google Sheets API
* Optional headless mode
* Error recovery and alert system
* CLI version with YAML task definitions
