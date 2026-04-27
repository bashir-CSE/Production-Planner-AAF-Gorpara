# Production Planner - Ayesha Abed Foundation (Gorpara)

A web-based production planning and costing application built with Google Apps Script, Tailwind CSS, and Google Drive integration. This tool allows the Ayesha Abed Foundation (AAF) Gorpara branch to generate, archive, and manage official production plans.

## 🚀 Features

-   **Dynamic Form Generation:** Add multiple items and processes dynamically to build comprehensive plans.
-   **Automated Costing Calculations:**
    -   Real-time calculation of Total Actual Costing and Section Paid Wages.
    -   Automated Overhead Costing distribution based on user-defined overhead input.
    -   Grand total summaries and percentage tracking.
-   **PDF Generation:** Automatically generates a professional PDF planning sheet with official branding.
-   **Data Archiving:**
    -   Saves records to a "History Sheets" Google Sheet.
    -   Stores generated PDFs in a dedicated Google Drive folder.
    -   Duplicate Order Number prevention.
-   **History Management:**
    -   View previously submitted planning sheets.
    -   Search through records by Date, Order No, or Notes.
    -   Update the status of orders (Pending/Complete) directly from the dashboard.
-   **Smart Autocomplete:** Fetches Item and Process names from a `settings` sheet for faster data entry.

## 🛠️ Technical Stack

-   **Backend:** Google Apps Script (`Code.gs`)
-   **Frontend:** HTML5, CSS3, JavaScript (`index.html`)
-   **UI Framework:** Tailwind CSS
-   **Icons:** Font Awesome 6.4.0
-   **Storage:** Google Sheets (Database) & Google Drive (File Storage)


## 📄 License

© 2026 Ayesha Abed Foundation — Gorpara, Manikganj.
Designed and Developed by **Md Bashir Ahmed**, MIS, AAF Gorpara.
