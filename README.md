# AAF Gorpara — Paperless Production Planning Portal

A robust, browser-based Production Planning web application built exclusively for **Ayesha Abed Foundation (AAF), Gorpara**. It is designed to completely eliminate the dependency on physical paper, solving the critical pain points of lost documents, manual filing errors, and physical storage limitations.

---

## 📌 Purpose: The Shift to Paperless

For a long time, managing production planning was a physical struggle. Maintaining stacks of planning sheets was not only tedious but highly risky—crucial papers frequently went missing, got damaged, or were misplaced during busy production cycles.

The primary goal of this application is to **eradicate the physical paper trail entirely**. By transitioning to a digital-first workflow, this system ensures that:
- **No planning sheet is ever lost again.** Every submitted plan is instantly archived in Google Drive with a direct URL.
- **Physical filing cabinets become obsolete.** Everything is stored securely in the cloud.
- **Retrieval takes seconds, not minutes.** Instead of digging through dusty folders, planners can search and pull up any historical plan instantly using the built-in search modal.

Beyond replacing paper, it automates the complex calculations of overhead costing and standardizes the output format across the organization.

---

## ✨ Key Features

- **Dynamic Table Management**: Add/remove multiple item tables and process rows on the fly without page reloads.
- **Smart Autocomplete**: Fetches Item Names and Process Names directly from a Google Sheets `settings` tab for fast, standardized data entry.
- **Live Calculations**: Real-time computation of Grand Totals (Actual Costing, Wages, Overhead).
- **Debounced Search**: The history modal search uses a 300ms debounce to ensure smooth filtering even with large datasets.
- **Pagination**: The history modal displays 15 records per page with Prev/Next navigation, handling thousands of records smoothly.
- **Status Filter**: Filter history records by All, Pending, or Complete with a single click.
- **Templated PDF Generation**: PDF layout is separated into its own `pdf_template.html` file, rendered via `HtmlService.createTemplateFromFile()`. Generates clean A4-sized PDFs using only web-safe CSS with a "Generated on" footer pinned to the absolute bottom.
- **Card-Based History View**: The "Planning Sheets Archive" modal uses a modern card-based layout for browsing submitted plans, with inline status toggling and PDF links.
- **Delete with Drive Cleanup**: Delete any record directly from the history modal. The system trashes the associated PDF from Google Drive and removes the row from the sheet.
- **Instant Status Tracking**: Update order statuses (Pending/Complete) with instant visual feedback (flash animations) and background server synchronization using optimized `createTextFinder` lookups.
- **Duplicate Prevention**: Blocks submission if the Order No already exists in the history sheet.
- **Transaction Safety**: If the PDF generates successfully but the Google Sheet write fails, the system automatically trashes the orphaned PDF to prevent digital clutter.

---

## ✅ Advantages

- **100% Paperless Workflow**: Eliminates printing, signing, filing, and storing physical sheets.
- **Zero Lost Documents**: Plans are archived to the cloud the second they are submitted.
- **Instant Retrieval**: Finding a specific order takes a 2-second text search.
- **Disaster-Proof Records**: Digital records in Google Drive are securely backed up.
- **Zero Infrastructure Cost**: Runs entirely on Google's free tier (Apps Script, Sheets, Drive).
- **Maintainable PDF Template**: The PDF layout is separated into its own HTML file, making it easy to modify without touching backend logic.
- **Optimized Lookups**: Status updates use `createTextFinder` for instant row targeting instead of iterating through all rows.

---

## ⚠️ Limitations

- **No User Authentication**: The app relies on Google Apps Script's default execution context. Anyone with the web app link can currently access, submit, and change statuses.
- **PDF Engine Constraints**: Google's native PDF renderer does not support custom fonts (like Inter) or complex CSS. The PDF uses `Arial` and basic HTML tables as a reliable workaround.
- **Race Conditions**: While duplicate checking is in place, two users submitting the exact same Order No at the exact same millisecond could theoretically bypass the check before either write completes.

---

## 🚀 Future Improvements

- **Role-Based Access**: Integrate basic email checks to restrict submissions/status updates to authorized MIS personnel only.
- **Edit Functionality**: Allow users to edit previously submitted plan details directly from the UI.
- **Dashboard Analytics**: Add a summary dashboard showing total plans per month, average overhead percentages, and pending vs. completed ratios.

---

## 📄 Credits

- **Developed & Designed By**: [Md Bashir Ahmed](https://www.linkedin.com/in/bashir-6923131b0/), MIS Department
- **Organization**: Ayesha Abed Foundation (AAF) — Gorpara, Manikganj
- **Year**: 2025

---
*© 2025 Ayesha Abed Foundation — Gorpara, Manikganj. All rights reserved.*