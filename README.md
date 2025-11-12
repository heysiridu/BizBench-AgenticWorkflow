# 🤖 ChatGPT Excel Automation (Agent Mode)

A fully automated **Python + Playwright** workflow that connects to a running Chrome session and controls **ChatGPT’s Agent Mode** to upload Excel files, issue prompts, and automatically download the AI-generated results — all without any manual clicks.

---

## 🚀 Overview

This script enables **programmatic control of ChatGPT through Chrome’s DevTools Protocol (CDP)**.  
It simulates a complete human interaction flow with ChatGPT’s **Agent Mode**, including:

1. Connecting to an existing Chrome session  
2. Finding an open ChatGPT tab  
3. Uploading an Excel file  
4. Sending a prompt for analysis  
5. Waiting for AI processing  
6. Detecting and downloading the generated spreadsheet  

It is designed for **reliability, transparency, and async performance**, using robust DOM-query fallbacks and retry logic for dynamic UI states.

---

## 🧰 Features

- ✅ **Asynchronous browser automation** using `asyncio` + Playwright  
- 🔄 **Reconnects to existing Chrome sessions** via CDP (no new browser windows)  
- 📎 **Uploads Excel files** directly to ChatGPT  
- 💬 **Sends structured prompts** automatically  
- ⏳ **Polls and verifies results** until downloadable Excel files are ready  
- 🧩 **Resilient selector system** supporting multiple fallback locators  
- 🪶 **Detailed logging** for every automation step  
- 🧠 **Ready for Agent-based workflow extensions**

---

## 🧑‍💻 Requirements

### 1. Install Dependencies

```bash
pip install playwright asyncio
playwright install chromium

### 2. Launch Chrome with Debugging Port
```bash
/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --remote-debugging-port=9222



