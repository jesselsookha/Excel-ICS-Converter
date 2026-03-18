# Teaching Timetable Converter

**Excel → ICS (Outlook / Apple / Google)**

■ **What it does**  
This tool converts a lecturer’s **Excel timetable** (as exported from the campus system) into an **ICS calendar file** that can be imported into **Outlook**, **Apple Calendar**, or **Google Calendar**. All parsing happens **locally in the browser**—no uploads to any server.

■ **Who it is for**  
Lecturers who receive a calendar‑style Excel timetable and want an up‑to‑date personal calendar with classroom sessions, while optionally including **Public Holidays** and **NOT AVAIL** days.

***

## Contents

*   \#features
*   \#quick-start-github-pages-or-local
*   \#file--folder-structure
*   \#usage-walkthrough
    *   \#1-export-the-excel-timetable
    *   \#2-convert-to-ics
    *   \#3-import-into-outlook--apple--google
    *   \#4-recommended-update-workflow
*   \#public-holidays-south-africa
*   \#configuration-update-holidays-for-a-new-year
*   \#data-handling--privacy
*   \#troubleshooting
*   \#technical-notes
*   \#license

***

## Features

● **Excel → ICS** in one page  
● **Local parsing** using a local copy of SheetJS (no network calls)  
● **Public Holiday recognition** for South Africa (by date)  
● Optional inclusion of **Holidays** and **NOT AVAIL** as **single all‑day blocks** per day (no duplicates per period)  
● **Stable UIDs** for events to reduce accidental duplicates if a user re‑imports into the same calendar  
● **Calendar metadata** in ICS (default calendar name: **Teaching Timetable**)  
● **Compact “Next 7 Days”** preview and a quick list of **Subjects & Groups**  
● Clean **light theme**, accessible styling, and **drag‑and‑drop** upload

***

## Quick Start (GitHub Pages or local)

1.  **Clone** or download the repository.
2.  Folder layout must be:
        /timetable-converter/
          index.html
          /styles/
            styles.css
          /js/
            app.js
            xlsx.full.min.js   ← local SheetJS (no CDN required)
3.  Open `index.html` directly in a browser, or host via GitHub Pages:
    *   In the repo settings → **Pages** → Source: **main** → **/(root)** → Save.
    *   URL will be `https://<your-username>.github.io/timetable-converter/`.

> Note: If a browser/extension blocks local scripts on `file://`, use a simple local server (e.g., VS Code Live Server). No backend is required.

***

## File & Folder Structure

    index.html              → UI layout and tutorial content
    styles/styles.css       → Light theme, SA palette, responsive layout
    js/app.js               → Excel parsing, holiday logic, ICS generation
    js/xlsx.full.min.js     → Local SheetJS build (Excel reader)

***

## Usage Walkthrough

### 1) Export the Excel Timetable

*   Go to the **live timetable** site: `masterscheduler.org` (sign in with Microsoft).
*   Click the **Excel** button to download your timetable.

### 2) Convert to ICS

*   Open this tool in your browser.
*   Drag & drop the Excel file (or click **Browse**).
*   Set **Year** (e.g., `2026`).
*   (Optional) Set **Calendar name** (default: `Teaching Timetable`).
*   Tick **Include Holidays & NOT AVAIL** if you want those in your calendar:
    *   The tool collapses multiple “NOT AVAIL” / Holiday entries for the same day into **one all‑day** entry.
*   Click **Parse Excel**.
*   Review **Stats**, **Subjects & Groups**, and **Next 7 Days**.
*   Click **Download ICS**.

### 3) Import into Outlook / Apple / Google

**Outlook (desktop/web)**

*   Open **Outlook Calendar** → Import the ICS.
*   A separate calendar named **Teaching Timetable** (or your chosen name) will be created and shown alongside your main calendar.

**Apple Calendar (macOS / iOS)**

*   macOS Calendar → **File → Import…** → choose the ICS → **New Calendar**.
*   On iOS, share the ICS file to the Calendar app and add to a new calendar.

**Google Calendar (web)**

*   Go to **calendar.google.com** → Left “Other calendars” → **+** → **Import**.
*   Choose ICS and **create a new calendar** or import into your existing timetable calendar.

### 4) Recommended Update Workflow

When the timetable changes:

1.  Export the **new Excel** from the system.
2.  Convert to a **new ICS** with this tool.
3.  In your calendar app, **delete the old “Teaching Timetable” calendar**.
4.  **Import** the new ICS into a **fresh** “Teaching Timetable” calendar.
5.  Done.

This avoids duplicates and ensures a clean, current timetable.

***

## Public Holidays (South Africa)

The converter recognises South African **Public Holidays** by **date** for the chosen year (e.g., **2026**, including observed Mondays where applicable) based on the Public Holidays Act (Act 36 of 1994) and official listings.  [\[timetiki.com\]](https://timetiki.com/holidays/south-africa/good-friday/)

*   **Good Friday** and **Family Day** are **movable** (Easter‑related) and are set for the specific year (e.g., **3 Apr 2026** and **6 Apr 2026**). [\[timetiki.com\]](https://timetiki.com/holidays/south-africa/good-friday/), [\[labourguide.co.za\]](https://labourguide.co.za/general/public-holidays-that-falls-on-sundays)

If you tick **Include Holidays & NOT AVAIL**, the ICS will add **one all‑day event** per holiday day.

***

## Configuration: Update Holidays for a New Year

You prefer to maintain holidays centrally in code (no “admin panel”). Update this block in **`js/app.js`**:

```js
const HOLIDAYS_BY_YEAR = {
  2026: [
    { date: '2026-01-01', name: "New Year's Day" },
    // …
    { date: '2026-12-26', name: 'Day of Goodwill' }
  ],
  2027: [
    // Add official dates for 2027 here
  ]
};
```

Notes:

*   Keep **ISO dates** (`YYYY-MM-DD`).
*   Include **observed Mondays** when a holiday falls on a **Sunday**, per the Public Holidays Act. 
*   For **Good Friday / Family Day**, use the official calendar for the relevant year.
  
***

## Data Handling & Privacy

*   Files are parsed **locally** in your browser using `xlsx.full.min.js`.
*   The tool does **not** send data to a server.
*   The ICS file is generated in memory and downloaded directly.

***

## Troubleshooting

● **“XLSX is not defined”**  
Ensure `./js/xlsx.full.min.js` exists and is referenced **above** `app.js` in `index.html`.  
Some browsers/extensions block scripts on `file://`. If so, serve locally (e.g., Live Server) or use GitHub Pages.

● **Days off by one (e.g., Tuesday shows Monday’s data)**  
The app formats dates as pure `YYYY-MM-DD` strings to avoid timezone shifts. If you changed that, restore the “manual YYYY-MM-DD” logic in `parseDateLabel()`.

● **Multiple NOT AVAIL periods appear as many events**  
Tick **Include Holidays & NOT AVAIL** to have them **collapsed** into a single **all‑day** block per day.

● **Duplicates after re-import**  
Use the recommended **delete‑and‑import** workflow (delete the old “Teaching Timetable” calendar first).  
Stable `UID`s minimize duplicates if someone re‑imports without deleting, but the cleanest approach is still **replace**.

***

## Technical Notes

*   **Parsing model** (Excel):
    *   Finds week markers like `w8` in **Column A**.
    *   Reads **day+date headers** from the **same row** (Columns **B–G**).
    *   Iterates **time rows** below (e.g., `08H00 - 08H50`) and collects cell contents for **B–G**.
    *   Extracts **module\_code**, **group**, **course\_code** (before “Group”), and **location** (e.g., `LR35 - CR`) from cell text.
    *   Converts day+month to `YYYY-MM-DD` using the chosen **Year** (no timezone conversions).

*   **ICS generation**:
    *   Adds `X-WR-CALNAME` and `X-WR-CALDESC`.
    *   Generates stable `UID`s (based on date, time, module/course, group, location).
    *   **Classes** are time‑bounded events.
    *   **Holidays** / **NOT AVAIL** (if included) are **all‑day** events, **one per day**.

*   **Libraries**:
    *   The project uses a **local** copy of **SheetJS** (`xlsx.full.min.js`) to parse Excel in the browser.

***

## License

This project is provided for campus/lecturer use.  

***

### Maintainer Notes (for your future self)

*   When a new year approaches, update **`HOLIDAYS_BY_YEAR`** in `app.js`.
*   Verify Easter dates (Good Friday, Family Day) and any observed Mondays from the official government listing.
***
