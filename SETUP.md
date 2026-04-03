# MatchaBoy POS — Setup Guide

Complete setup guide for the MatchaBoy cashier and inventory management system.

---

## Step 1: Create a Google Spreadsheet

1. Go to [Google Sheets](https://sheets.google.com) and create a new blank spreadsheet
2. Name it something like **"MatchaBoy POS - Main Branch"**
3. Copy the **Spreadsheet ID** from the URL:
   ```
   https://docs.google.com/spreadsheets/d/SPREADSHEET_ID_HERE/edit
   ```
   The ID is the long string between `/d/` and `/edit`

> **Tip:** Create one spreadsheet per branch. Each branch has its own isolated data.

---

## Step 2: Set Up Google Apps Script

1. Go to [script.google.com](https://script.google.com)
2. Click **New Project**
3. Delete the default `myFunction()` code
4. Copy the entire contents of `Code.gs` and paste it in
5. **Edit the BRANCHES config** at the top of the file:

```javascript
const BRANCHES = {
  "main": {
    password: "1234",                          // Choose a secure password
    spreadsheetId: "YOUR_SPREADSHEET_ID_HERE"  // Paste the ID from Step 1
  }
};
```

6. Click **Save** (Ctrl+S) and name the project **"MatchaBoy POS Backend"**

---

## Step 3: Deploy as Web App

This is the most critical step. Follow carefully:

1. In the Apps Script editor, click **Deploy** → **New deployment**
2. Click the gear icon ⚙ next to "Select type" and choose **Web app**
3. Set these options:
   - **Description:** `POS Backend v1`
   - **Execute as:** `Me`
   - **Who has access:** `Anyone`
4. Click **Deploy**
5. **Authorize** when prompted:
   - Click "Review permissions"
   - Select your Google account
   - Click "Advanced" → "Go to MatchaBoy POS Backend (unsafe)"
   - Click "Allow"
6. **Copy the Web App URL** — it looks like:
   ```
   https://script.google.com/macros/s/AKfycb.../exec
   ```

> ⚠ **IMPORTANT:** Every time you change `Code.gs`, you must create a **New Deployment** (not just save). The old URL keeps running the old code. Go to **Deploy → New deployment** each time.

---

## Step 4: Configure the App

1. Open `index.html` in a web browser (double-click the file or open from a local/web server)
2. Log in with:
   - **Branch Name:** `main` (or whatever you named it in BRANCHES)
   - **Password:** `1234` (or whatever you set)
3. Go to **Settings** (⚙ tab at the bottom)
4. Paste the **Web App URL** from Step 3
5. Click **Save & Test**
6. You should see ✓ **Connected** — the sync dot in the top bar turns green

> If you see an error, see Troubleshooting below.

---

## Step 5: First-Time Data Setup

The system auto-creates all required sheet tabs (Menu, Ingredients, Recipes, Orders, etc.) on the first request. You can:

1. Go to the **Stock** tab → Add your ingredients (coffee beans, milk, sugar, etc.)
2. Go to the **Menu** tab → Add your menu items with prices and ingredient links
3. Go to the **Cashier** tab → Start taking orders!

---

## Sharing Across Multiple Devices

The system works on any device with a modern web browser:

### Option A: Local File Sharing
1. Share the `index.html` file via USB, email, Google Drive, etc.
2. Each device opens the file in their browser
3. Enter the same Web App URL + branch credentials
4. All devices see the same live data from Google Sheets

### Option B: Host on a Web Server
1. Upload `index.html` to any static hosting (GitHub Pages, Netlify, Vercel, etc.)
2. Share the URL with your team
3. Everyone logs in with their branch credentials

### Option C: Share via Google Drive
1. Upload `index.html` to Google Drive
2. Download on each device and open in browser

> **Key point:** The HTML file is just the interface. All data lives in Google Sheets, so any device with the correct URL + credentials sees the same data.

---

## Adding More Branches

Each branch gets its own Google Spreadsheet for data isolation.

1. Create a new Google Spreadsheet for the branch
2. Copy the Spreadsheet ID
3. Edit `Code.gs` and add the branch to the BRANCHES config:

```javascript
const BRANCHES = {
  "main": {
    password: "1234",
    spreadsheetId: "SPREADSHEET_ID_FOR_MAIN"
  },
  "branch2": {
    password: "5678",
    spreadsheetId: "SPREADSHEET_ID_FOR_BRANCH2"
  },
  "branch3": {
    password: "abcd",
    spreadsheetId: "SPREADSHEET_ID_FOR_BRANCH3"
  }
};
```

4. **Deploy as a New Deployment** (Deploy → New deployment)
5. Share the **new** Web App URL with all devices
6. Each branch logs in with their own name + password

---

## Offline / Demo Mode

- Click **"Offline / Demo Mode"** on the login page to use the app without Google Sheets
- All data is stored in the browser's localStorage only
- Great for testing or when internet is unavailable
- Data does NOT sync to Google Sheets in this mode

---

## Troubleshooting

### "Connection failed" or sync dot stays red

1. **Check the URL** — it must end with `/exec`, not `/dev`
2. **Create a New Deployment** — just saving the script is NOT enough. You must go to Deploy → New deployment
3. **Check "Who has access"** is set to "Anyone" (not "Only myself")
4. **Check the branch name and password** match exactly what's in `Code.gs` (case-sensitive)
5. **Try in an incognito window** to rule out browser extensions blocking requests

### "Unknown branch" error

- Branch names are case-sensitive. If the config says `"main"`, you must type `main`, not `Main`

### Data not showing on another device

- Make sure both devices use the **same** Web App URL
- Click the **Sync** button in the top bar to force a fresh load
- Check that both devices are logged into the same branch

### Spreadsheet tabs not created

- The tabs are created automatically on the first API request
- Make sure the Apps Script has permission to access the spreadsheet
- The Spreadsheet ID must be correct

### CORS errors in browser console

- This should NOT happen if you're using the system correctly
- The app sends all requests as GET with `?payload=` to avoid CORS
- If you see CORS errors, make sure you deployed as a Web App (not just running in the editor)

### Slow responses

- Google Apps Script has a cold start of ~2-5 seconds on the first request
- Subsequent requests are faster
- Large datasets (1000+ orders) may slow down — consider archiving old data periodically

---

## Technical Notes

### How the CORS fix works

Standard `fetch()` POST with `Content-Type: application/json` triggers a CORS preflight (OPTIONS request) that Google Apps Script cannot respond to. The solution:

- All data is sent as a **GET request** with a URL parameter called `payload`
- The payload contains JSON-encoded data: `{action, branch, password, data}`
- Google Apps Script's `doGet(e)` reads `e.parameter.payload` and parses it
- GET requests don't trigger CORS preflight, so it works from any origin

### Data flow

```
Browser → GET ?payload={action,branch,password,data} → Google Apps Script → Google Sheets
Browser ← JSON response ← Google Apps Script ← Google Sheets
```

### localStorage backup

All data is also saved to `localStorage` as a backup. If the network fails, the app continues working with local data and syncs when connectivity returns.

---

## File Structure

```
cashier/
├── index.html   — Complete web app (HTML + CSS + JS in one file)
├── Code.gs      — Google Apps Script backend (paste into script.google.com)
└── SETUP.md     — This setup guide
```
