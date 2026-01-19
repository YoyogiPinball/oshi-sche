# Oshi-Sche

> Just upload your favorite VTuber's schedule images to Google Drive, and they'll be automatically added to your calendar!

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)

## 🎯 Overview

An automated system that analyzes VTuber stream schedule images using **Gemini AI** and automatically registers them to Google Sheets & Google Calendar.

Perfect for fans who follow multiple VTubers and find it challenging to keep track of all streaming schedules!

---

## ✨ Key Features

- 📸 **Automatic Image Analysis**: Analyze schedule images using Gemini 2.0 Flash API
- 📊 **Spreadsheet Management**: Automatic writing to Google Sheets
- 📅 **Calendar Integration**: Automatic event registration to Google Calendar (1-hour time slot)
- 🤖 **Full Automation**: Periodic execution with Google Apps Script
- 🔄 **Duplicate Prevention**: Automatically skip already processed images
- 👥 **Multi-VTuber Support**: Manage multiple VTubers with folder organization
- 🧪 **Dry Run Mode**: Test safely before actual execution

---

## 🛠️ Setup (Step-by-Step)

### Step 1: Obtain Gemini API Key

1. Visit [Google AI Studio](https://aistudio.google.com/)
2. Click "Get API Key"
3. Copy and save the API key (you'll need it later)

> **Note**: Free tier is approximately 1500 requests per day. Exceeding this may result in charges.

---

### Step 2: Create Google Drive Folders

1. Create two folders in Google Drive:
   - `Oshi-Sche_Input` (folder for uploading schedule images)
   - `Oshi-Sche_Done` (folder where processed images are moved)

2. Create subfolders for each VTuber within these folders:
   - Folder name format: `Affiliation：VTuber Name` (using full-width colon "：")
   - Examples: `Individual：Shirakami Mishiro`, `Mixist：Yukishiro Cal`

3. Note the folder IDs:
   - Open the folder and check the URL
   - `https://drive.google.com/drive/folders/【This is the folder ID】`

---

### Step 3: Create Google Spreadsheet

1. Create a new Google Spreadsheet
2. Note the Spreadsheet ID:
   - The `/d/【This is the Spreadsheet ID】/edit` part of the URL

---

### Step 4: Create Google Calendar (Optional)

1. Create a new calendar in Google Calendar (recommended)
2. Note the Calendar ID:
   - Calendar Settings → "Integrate calendar"
   - Copy the Calendar ID (e.g., `xxxx@group.calendar.google.com`)

---

### Step 5: Create GAS Project

1. Visit [Google Apps Script](https://script.google.com/)
2. Click "New project"
3. Rename the project to "Oshi-Sche"
4. Delete all content in `code.gs`
5. Copy and paste the content from [src/code.gs](src/code.gs)
6. Save (click 💾 icon)

---

### Step 6: Configure Script Properties

1. Click "Project Settings (⚙ icon)" in the GAS editor
2. Scroll to the "Script Properties" section
3. Add the following properties by clicking "Add property":

| Property Name | Description | How to Obtain | Example |
|--------------|-------------|---------------|---------|
| `GEMINI_API_KEY` | Gemini API Key | Google AI Studio | `AIza...` |
| `INPUT_FOLDER_ID` | Input Folder ID | Drive URL | `173ibp...` |
| `DONE_FOLDER_ID` | Done Folder ID | Drive URL | `18I6b...` |
| `SPREADSHEET_ID` | Spreadsheet ID | Sheet URL | `1rWFI...` |
| `SHEET_NAME` | Sheet Name | Arbitrary | `Schedule` |
| `CALENDAR_ID` | Calendar ID | Calendar Settings | `xxxx@group.calendar.google.com` |
| `DRY_RUN` | Dry Run Mode | `true` or `false` | `true` |

> **Important**: Always set `DRY_RUN` to `true` for initial testing!

---

### Step 7: Initial Test Execution

1. In the GAS editor, select "Run" → "testProcessImages"
2. You'll be asked to authorize permissions on first run:
   - Click "Review permissions"
   - Select your Google account
   - Click "Advanced" → "Go to Oshi-Sche (unsafe)"
   - Click "Allow"
3. Check the execution log ("Executions" → "Latest execution")

In **Dry Run Mode**, only image analysis is performed, without actual writing or moving.

---

### Step 8: Set Up Trigger (Periodic Execution)

Once testing is successful, set up periodic execution.

1. Click "Triggers (⏰ icon)" in the GAS editor
2. Click "Add Trigger"
3. Configure as follows:
   - Function to run: `processNewImages`
   - Deployment to run: `Head`
   - Event source: `Time-driven`
   - Type of time-based trigger: `Hour timer`
   - Time interval: `Every hour` (adjust as needed)
4. Click "Save"

---

### Step 9: Production Execution

After confirming proper operation in dry run mode, switch to production mode.

1. Change the `DRY_RUN` script property to `false`
2. Test again to confirm actual writing occurs

---

## 📋 Usage

1. **Prepare Schedule Image**
   - Save schedule images published by VTubers

2. **Upload to Google Drive**
   - Upload images to `Oshi-Sche_Input/Affiliation：VTuber Name/` folder

3. **Wait for Automatic Processing**
   - Automatically processed at the set trigger interval
   - Upon completion:
     - Schedule added to Spreadsheet
     - Events registered in Calendar
     - Image moved to `Oshi-Sche_Done/Affiliation：VTuber Name/`

4. **Verify**
   - Check schedule in Spreadsheet or Calendar
   - With Claude (MCP integration), you can ask "What streams are scheduled for today?"

---

## 🎨 Recommended Image Format

For better Gemini AI analysis accuracy, the following image characteristics are ideal:

### ✅ Recommended
- Horizontally-oriented schedule images
- Simple layout
- Reasonably large text size
- High contrast between background and text

### ⚠️ May Reduce Accuracy
- Vertical text orientation
- Overly decorative design
- Handwritten-style fonts
- Text too small
- Blurry images

**Track Record**: Tested with 10+ VTubers' schedule images, no discrepancies!

---

## 📊 Workflow

```
[Manual] Upload images to Drive "Input Folder/Affiliation：VTuber Name/"
    ↓
[Auto] GAS periodic execution (trigger setting)
    ↓
[Auto] Gemini 2.0 Flash analyzes image → Extract schedule JSON
    ↓
[Auto] Write to Google Spreadsheet
    ↓
[Auto] Register events to Google Calendar
    ↓
[Auto] Move processed images to "Done Folder/Affiliation：VTuber Name/"
    ↓
[Manual] Ask Claude via MCP integration "What streams are scheduled for today?"
```

---

## 💰 Gemini API Free Tier

**Free Tier**:
- 15 requests per minute
- Approximately 1500 requests per day

**Important Notes**:
- Exceeding the free tier may result in charges
- Be cautious when processing many images at once
- For details, check [Google AI Studio Pricing](https://ai.google.dev/pricing)

---

## 🔒 Permissions and Security

This tool requires the following permissions and explains why.

### Required Permissions

| Permission | Reason |
|-----------|--------|
| Google Drive Access | To read and move schedule images |
| Google Sheets Access | To write schedule data |
| Google Calendar Access | To register stream events |
| External Service Connection | To send requests to Gemini API |

### Security Notes

- ⚠️ **Never share your API key publicly**
- ⚠️ Storing in GAS Script Properties prevents hardcoding API keys in code
- ⚠️ If sharing with others, remove your API key first

---

## ❓ FAQ / Known Limitations

### Q: Are year-end/new-year schedules handled correctly?

**A**: Generally supported. The logic works as follows:

- Default to current year
- If current month is December and schedule contains January/February, use next year
- If current month is January and schedule contains December, use previous year

However, schedules several months ahead may be misjudged.

### Q: What if there are multiple streams on the same day?

**A**: They are registered as separate entries.

### Q: How to update a schedule image?

**A**:
1. Delete the original image
2. Upload the new image
3. Manually delete old events from Spreadsheet/Calendar
4. Wait for automatic processing

### Q: Processing doesn't execute

**A**: Please check:
- Is the trigger properly configured?
- Are all script properties set?
- Are there errors in the execution log? (Check "Executions" in GAS editor)

### Q: Getting errors

**A**: Check the execution log. Error messages include troubleshooting steps.

---

## 🛡️ Disclaimer

- ⚠️ This tool uses Gemini API. Use at your own risk.
- ⚠️ This is an unofficial tool.
- ⚠️ Malicious use is prohibited.
- ⚠️ Do not use in ways that harm others.
- ⚠️ For personal use only.
- ⚠️ Please also comply with [Gemini API Terms of Service](https://ai.google.dev/gemini-api/terms).
- ⚠️ Copyrights of VTuber/talent schedule images belong to their respective agencies/individuals.

---

## 🤝 Feature Requests & Bug Reports

Issues and Pull Requests are welcome!

- [Create an Issue](https://github.com/YoyogiPinball/oshi-sche/issues)
- [Create a Pull Request](https://github.com/YoyogiPinball/oshi-sche/pulls)

---

## 📄 License

[MIT License](LICENSE)

Copyright (c) 2025 YoyogiPinball

---

## 🚀 Future Plans

- [ ] More detailed image analysis (URL, hashtag extraction, etc.)
- [ ] LINE notification feature
- [ ] Multiple calendar distribution feature
- [ ] Web UI provision

---

## 🎉 Acknowledgments

We hope this tool helps you never miss your favorite VTuber's streams!

---

**Never miss your oshi's streams. Accelerate your oshi-katsu with Oshi-Sche!** 🚀
