# Teaching Timetable Converter

**Excel → ICS (Outlook / Apple / Google)**

![GitHub Pages](https://img.shields.io/badge/Live%20Site-GitHub%20Pages-blue?logo=github)
![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)
![Built with JavaScript](https://img.shields.io/badge/Built%20with-JavaScript-yellow?logo=javascript)
![Privacy](https://img.shields.io/badge/Data-Local%20Only-success)

---

## ◆ What it does

This tool converts a lecturer’s **Excel timetable** (as exported from the campus system) into an **ICS calendar file** that can be imported into **Outlook**, **Apple Calendar**, or **Google Calendar**.

All parsing happens **locally in the browser**—no uploads to any server.

---

## ◇ Who it is for

Lecturers who receive a calendar-style Excel timetable and want an up-to-date personal calendar with classroom sessions, while optionally including:

* **Public Holidays**
* **NOT AVAIL** days

---

🔗 [Live Site](https://jesselsookha.github.io/Excel-ICS-Converter/)

---

## ◈ Contents

* [Features](#features)
* [Quick Start (GitHub Pages or local)](#quick-start-github-pages-or-local)

  * [File & Folder Structure](#file--folder-structure)
  * [Usage Walkthrough](#usage-walkthrough)

    * [1) Export the Excel Timetable](#1-export-the-excel-timetable)
    * [2) Convert to ICS](#2-convert-to-ics)
    * [3) Import into Outlook / Apple / Google](#3-import-into-outlook--apple--google)
    * [4) Recommended Update Workflow](#4-recommended-update-workflow)
* [Public Holidays (South Africa)](#public-holidays-south-africa)
* [Configuration: Update Holidays for a New Year](#configuration-update-holidays-for-a-new-year)
* [Data Handling & Privacy](#data-handling--privacy)
* [Troubleshooting](#troubleshooting)
* [Technical Notes](#technical-notes)
* [Disclaimer](#disclaimer)
* [Author](#author)
* [License](#license)

---

## ◉ Features

* **Excel → ICS** in one page
* **Local parsing** using a local copy of SheetJS (no network calls)
* **Public Holiday recognition** for South Africa (by date)
* Optional inclusion of **Holidays** and **NOT AVAIL** as **single all-day blocks** per day
* **Stable UIDs** to reduce duplicate imports
* **Calendar metadata** (default: *Teaching Timetable*)
* **Next 7 Days preview** and **Subjects & Groups summary**
* Clean **light theme** with **drag-and-drop upload**

---

## ◉ Quick Start (GitHub Pages or local)

1. **Clone** or download the repository.

2. Ensure folder layout:

```
/timetable-converter/
  index.html
  /styles/
    styles.css
  /js/
    app.js
    xlsx.full.min.js
```

3. Open locally or host via GitHub Pages:

* Go to **Settings → Pages**
* Source: **main branch / (root)**
* Save

Your site will be available at:

```
https://<your-username>.github.io/timetable-converter/
```

> ◊ Note: If scripts are blocked on `file://`, use a local server (e.g., VS Code Live Server).

---

## ◉ File & Folder Structure

```
index.html              → UI layout and guidance
styles/styles.css       → Styling and layout
js/app.js               → Core logic (parsing + ICS generation)
js/xlsx.full.min.js     → Local SheetJS library
```

---

## ◉ Usage Walkthrough

### ◇ 1) Export the Excel Timetable

* Go to the timetable system (`masterscheduler.org`)
* Sign in with Microsoft
* Click **Excel** to download your timetable

---

### ◇ 2) Convert to ICS

* Open the tool in your browser

* Drag & drop the Excel file (or click **Browse**)

* Set:

  * **Year** (e.g., `2026`)
  * **Calendar name** (optional)

* Optional:

  * Tick **Include Holidays & NOT AVAIL**

  ◊ The tool collapses multiple entries into **one all-day event per day**

* Click **Parse Excel**

* Review preview sections

* Click **Download ICS**

---

### ◇ 3) Import into Outlook / Apple / Google

**Outlook**

* Open Calendar → Import ICS
* A new calendar will be created

**Apple Calendar**

* macOS: File → Import
* iOS: Share file → Add to Calendar

**Google Calendar**

* Go to calendar.google.com
* Other calendars → **+ → Import**

---

### ◇ 4) Recommended Update Workflow

When changes occur:

1. Export updated Excel
2. Convert to new ICS
3. Delete old **Teaching Timetable** calendar
4. Import new ICS into a fresh calendar

✔ This avoids duplicates and keeps things clean

---

## ◉ Public Holidays (South Africa)

The tool recognises South African public holidays **by date** for the selected year.

* Includes observed Mondays where applicable
* Based on official holiday listings

◊ Movable holidays:

* **Good Friday**
* **Family Day**

If enabled, holidays are added as **single all-day events**

---

## ◉ Configuration: Update Holidays for a New Year

Update in `js/app.js`:

```js
const HOLIDAYS_BY_YEAR = {
  2026: [
    { date: '2026-01-01', name: "New Year's Day" },
    { date: '2026-12-26', name: 'Day of Goodwill' }
  ],
  2027: [
    // Add new dates here
  ]
};
```

### ◊ Notes

* Use `YYYY-MM-DD` format
* Include observed Mondays
* Verify Easter-related dates annually

---

## ◉ Data Handling & Privacy

* All processing happens **locally in your browser**
* No uploads or server communication
* ICS file is generated and downloaded directly

---

## ◉ Troubleshooting

**XLSX is not defined**

* Ensure `xlsx.full.min.js` loads before `app.js`

**Dates are incorrect**

* Use `YYYY-MM-DD` formatting logic (avoid timezone shifts)

**Too many NOT AVAIL entries**

* Enable collapse option via checkbox

**Duplicate events after import**

* Follow **delete + re-import workflow**

---

## ◉ Technical Notes

### ◇ Parsing Logic

* Detects week markers (`w8`) in Column A
* Reads headers (Columns B–G)
* Iterates time rows
* Extracts:

  * Module code
  * Group
  * Course code
  * Location

---

### ◇ ICS Generation

* Adds calendar metadata
* Generates stable UIDs
* Creates:

  * Timed class events
  * All-day holiday/NOT AVAIL events

---

### ◇ Libraries

* Uses local **SheetJS** (`xlsx.full.min.js`)
* No CDN or external dependency

---

## ◉ Disclaimer

This tool was developed specifically for the internal timetable format used at this institution.

The imported Excel file must follow the expected structure defined by the campus scheduling system.

* It may not work with other institutional formats
* Structural differences may result in parsing errors

This tool is intended for **internal lecturer use only**.

---

## ◉ Author

**Jessel Sookha**  
📧 [jsookha@emeris.ac.za](mailto:jsookha@emeris.ac.za)

Developed for internal academic use.

---

## ◉ License

This project is provided for campus/lecturer use.

---

### ◊ Maintainer Notes

* Update `HOLIDAYS_BY_YEAR` annually
* Verify Easter dates and observed holidays
* Test with latest exported timetable format

---
![Version](https://img.shields.io/badge/version-1.0-blue)
![Maintained](https://img.shields.io/badge/Maintained-yes-brightgreen)

