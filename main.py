import asyncio
from playwright.async_api import async_playwright
import os

async def automate_chatgpt_excel(excel_path, prompt_text):
    async with async_playwright() as p:
        # Connect to existing Chrome browser
        print("   Connecting to Chrome browser...")
        print("   Make sure Chrome was started with: --remote-debugging-port=9222")
        
        try:
            browser = await p.chromium.connect_over_cdp("http://localhost:9222")
        except Exception as e:
            print(f"Could not connect to Chrome: {e}")
            print("\nTo fix this:")
            print("1. Close all Chrome windows")
            print("2. Start Chrome with debugging port:")
            print('   Mac: /Applications/Google\\ Chrome.app/Contents/MacOS/Google\\ Chrome --remote-debugging-port=9222')
            return
        
        # Get the first context and page
        contexts = browser.contexts
        if not contexts:
            print("No browser context found")
            return
        
        context = contexts[0]
        pages = context.pages
        
        # Find ChatGPT tab
        chatgpt_page = None
        for page in pages:
            if "chat.openai.com" in page.url or "chatgpt.com" in page.url:
                chatgpt_page = page
                break
        
        if not chatgpt_page:
            print("❌ ChatGPT tab not found. Please open ChatGPT in Chrome first")
            return
        
        print("✅ Found ChatGPT tab")
        await chatgpt_page.bring_to_front()
        
        # Step 1: Click the plus button to open menu
        print("📎 Step 1: Clicking plus button [Add files and more]...")
        
        plus_button_selectors = [
            'button[data-testid="composer-plus-btn"]',
            'button#composer-plus-btn',
            'button[aria-label="Add files and more"]',
            'button.composer-btn'
        ]
        
        clicked = False
        for selector in plus_button_selectors:
            try:
                await chatgpt_page.click(selector, timeout=3000)
                print(f"✅ Clicked plus button")
                clicked = True
                break
            except:
                continue
        
        if not clicked:
            print("❌ Could not find plus button")
            return
        
        await asyncio.sleep(1)
        
        # Step 2: Click "Agent mode" in dropdown
        print("🤖 Step 2: Selecting Agent mode...")
        
        agent_mode_selectors = [
            'div.truncate:has-text("Agent mode")',
            'div:text("Agent mode")',
            'text="Agent mode"',
            '[role="menuitem"]:has-text("Agent mode")',
            '*:has-text("Agent mode")'
        ]
        
        agent_selected = False
        for selector in agent_mode_selectors:
            try:
                await chatgpt_page.click(selector, timeout=3000)
                print("✅ Selected Agent mode")
                agent_selected = True
                break
            except:
                continue

        if not agent_selected:
            print("⚠️ Warning: Could not confirm Agent mode selection")

        await asyncio.sleep(3)

        # Verify Agent mode is active
        try:
            agent_indicator = await chatgpt_page.query_selector('text="Agent"')
            if agent_indicator:
                print("✅ Agent mode confirmed active")
            else:
                print("⚠️ Warning: Agent mode may not be active")
        except:
            print("⚠️ Could not verify Agent mode status")

        # Step 3: Click attachment icon to upload files
        print("📎 Step 3: Clicking attachment icon to upload files...")

        attachment_selectors = [
            'button[aria-label="Attach files"]',
            'button[data-testid="composer-attach-btn"]',
            'button:has(svg[class*="paperclip"])',
            'button[data-testid="composer-plus-btn"]',
        ]

        clicked_attachment = False
        for selector in attachment_selectors:
            try:
                await chatgpt_page.click(selector, timeout=3000)
                print(f"✅ Clicked attachment button")
                clicked_attachment = True
                break
            except:
                continue

        if not clicked_attachment:
            print("⚠️ Warning: Could not find attachment button, trying to upload directly")

        await asyncio.sleep(1)

        # Step 4: Upload the Excel file
        print(f"📤 Step 4: Uploading {excel_path}...")
        
        abs_excel_path = os.path.abspath(excel_path)
        if not os.path.exists(abs_excel_path):
            print(f"❌ File not found: {abs_excel_path}")
            print(f"   Make sure {excel_path} is in the same folder as this script")
            return
        
        file_input = await chatgpt_page.query_selector('input[type="file"]')
        if file_input:
            await file_input.set_input_files(abs_excel_path)
            print(f"✅ Uploaded file: {os.path.basename(abs_excel_path)}")
            print(f"   Full path: {abs_excel_path}")
        else:
            print("❌ Could not find file upload input")
            return
        
        print("⏳ Waiting for file to process...")
        await asyncio.sleep(2)

        # Step 5: Type the prompt
        print(f"💬 Step 5: Typing prompt: {prompt_text}")
        
        input_selectors = [
            'div[class*="grid-area:leading"]',
            'textarea[placeholder*="Message"]',
            'textarea[id="prompt-textarea"]',
            'div[contenteditable="true"]',
            'textarea[data-id="root"]',
            'textarea'
        ]
        
        typed = False
        for selector in input_selectors:
            try:
                element = await chatgpt_page.query_selector(selector)
                if element:
                    await element.click()
                    await chatgpt_page.keyboard.type(prompt_text)
                    print(f"✅ Typed prompt")
                    typed = True
                    break
            except:
                continue
        
        if not typed:
            print("⚠️ Couldn't find text input, typing directly...")
            await chatgpt_page.keyboard.type(prompt_text)
        
        # Step 6: Submit the prompt
        print("📨 Step 6: Submitting prompt...")
        
        send_selectors = [
            'button[data-testid="send-button"]',
            'button[aria-label*="Send"]',
            'button:has(svg[class*="arrow"])',
            'button.absolute.rounded-md:has(svg)'
        ]
        
        sent = False
        for selector in send_selectors:
            try:
                await chatgpt_page.click(selector, timeout=2000)
                print(f"✅ Clicked send button")
                sent = True
                break
            except:
                continue
        
        if not sent:
            await chatgpt_page.keyboard.press('Enter')
            print("✅ Pressed Enter to submit")
        
        # Step 7: IMPROVED Download Section
        print("\n⏳ Step 7: Waiting for ChatGPT Agent to process...")
        print("⏰ This typically takes 2-6 minutes for Excel analysis...")
        print("📊 Will start checking for results after 1 minute...\n")

        # Configuration
        initial_wait = 15  # Wait 1 minute before starting to check
        check_interval = 10  # Check every 20 seconds
        max_wait_time = 360  # Maximum 6 minutes
        elapsed_time = 0
        
        # Initial wait
        print(f"⏳ Initial wait: {initial_wait} seconds...")
        await asyncio.sleep(initial_wait)
        elapsed_time = initial_wait
        print(f"✅ Starting to check for results\n")

        # Start scrolling and checking for download buttons
        download_found = False
        original_filename = os.path.basename(excel_path).lower()
        
        while elapsed_time < max_wait_time and not download_found:
            minutes = elapsed_time // 60
            seconds = elapsed_time % 60
            print(f"[{minutes}m {seconds}s] Checking for download button...")
            
            # Scroll to bottom to see latest messages
            await chatgpt_page.evaluate('window.scrollTo(0, document.body.scrollHeight)')
            await asyncio.sleep(2)  # Give time for content to load
            
            try:
                assistant_download = None
                
                # Get the last assistant message
                assistant_messages = await chatgpt_page.query_selector_all('[data-message-author-role="assistant"]')
                
                if assistant_messages:
                    print(f"   Found {len(assistant_messages)} assistant message(s)")
                    last_assistant_msg = assistant_messages[-1]
                    
                    # Use JavaScript to find downloadable elements efficiently
                    download_info = await last_assistant_msg.evaluate('''
                        (element) => {
                            const uuidPattern = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i;

                            // Find all buttons with SVG icons (download buttons)
                            const buttons = element.querySelectorAll('button');

                            for (let i = 0; i < buttons.length; i++) {
                                const button = buttons[i];
                                const svg = button.querySelector('svg');

                                // Check if button contains an SVG (potential download button)
                                if (svg) {
                                    // Get the parent container to check nearby text
                                    let parent = button.parentElement;
                                    let searchDepth = 0;
                                    let nearbyText = '';

                                    // Search up to 3 levels up to find text context
                                    while (parent && searchDepth < 3) {
                                        nearbyText += parent.textContent || '';
                                        parent = parent.parentElement;
                                        searchDepth++;
                                    }

                                    nearbyText = nearbyText.toLowerCase();

                                    // CRITICAL: Only accept if we see "spreadsheet" or ".xlsx" or ".xls" near the button
                                    const hasSpreadsheetIndicator = nearbyText.includes('spreadsheet') ||
                                                                     nearbyText.includes('.xlsx') ||
                                                                     nearbyText.includes('.xls') ||
                                                                     nearbyText.includes('excel');

                                    // Must NOT contain UUID
                                    const hasUUID = uuidPattern.test(nearbyText);

                                    if (hasSpreadsheetIndicator && !hasUUID) {
                                        console.log('✓ Found valid download button with spreadsheet indicator');
                                        return {found: true, type: 'button', index: i, filename: 'spreadsheet-download'};
                                    }
                                }
                            }

                            return {found: false};
                        }
                    ''')
                    
                    if download_info and download_info.get('found'):
                        # Get the actual element based on type and index
                        filename = download_info.get('filename', 'unknown')
                        print(f"   ✅ Found valid Excel file: {filename}")

                        if download_info['type'] == 'link':
                            links = await last_assistant_msg.query_selector_all('a')
                            if download_info['index'] < len(links):
                                assistant_download = links[download_info['index']]
                                print(f"   📎 Download link ready")
                        elif download_info['type'] == 'button':
                            buttons = await last_assistant_msg.query_selector_all('button')
                            if download_info['index'] < len(buttons):
                                assistant_download = buttons[download_info['index']]
                                print(f"   📎 Download button ready")
                    else:
                        print(f"   ⏳ No valid Excel file found yet (skipping UUID/temp files)")
                
                # If found, click the download
                if assistant_download:
                    try:
                        # ADDED: Wait 5 seconds after finding download element
                        # This prevents clicking on temporary/incomplete files (UUID-named files)
                        # Gives ChatGPT time to fully generate the processed Excel file
                        print(f"   ⏳ Download element found! Waiting 5 seconds for file generation...")
                        await asyncio.sleep(5)  # ADDED: 5-second delay for file generation
                        
                        # Original code: Scroll element into view and click
                        await assistant_download.scroll_into_view_if_needed()
                        await asyncio.sleep(1)
                        await assistant_download.click()
                        print(f"✅ Successfully clicked assistant's download button!")
                        download_found = True
                        await asyncio.sleep(3)
                    except Exception as e:
                        print(f"⚠️ Failed to click: {e}")
                else:
                    print(f"   No download found yet, assistant still processing...")
                    
            except Exception as e:
                print(f"   Error: {e}")
            
            if not download_found:
                print(f"   ⏳ Waiting {check_interval} seconds...\n")
                await asyncio.sleep(check_interval)
                elapsed_time += check_interval
        
        # Final status
        if download_found:
            print(f"\n🎉 Successfully downloaded result after {elapsed_time} seconds (~{elapsed_time//60} minutes)!")
            print("📁 Check your Downloads folder for the Excel file")
        else:
            print(f"\n⏱️ Timeout after {max_wait_time} seconds ({max_wait_time//60} minutes)")
            print("   The agent might still be processing.")
            print("   You can:")
            print("   1. Check the browser window manually")
            print("   2. Run the script again with a longer timeout")
            print("   3. Wait and manually download when ready")
        
        print("\n✅ Automation complete!")
        print("🌐 Browser staying open for you to review...")

def main():
    """Main entry point with hardcoded Excel file"""
    
    print("="*50)
    print("ChatGPT Excel Automation (Agent Mode)")
    print("="*50)
    print("\n⚠️ Prerequisites:")
    print("1. Chrome must be running with debugging port")
    print("2. ChatGPT must be open and logged in")
    print("3. test.xlsx must be in current directory\n")
    
    # CHANGED: Hardcoded Excel filename - no user input needed
    excel_file = "test.xlsx"
    print(f"📄 Using Excel file: {excel_file}")
    
    # Use the specific prompt requested
    prompt = "create a summary sheet for this file as a new sheet and give me back an excel to download"
    print(f"💬 Using prompt: {prompt}\n")
    
    # Run automation
    asyncio.run(automate_chatgpt_excel(excel_file, prompt))

if __name__ == "__main__":
    main()