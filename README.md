# Production Planner - Ayesha Abed Foundation (Gorpara)

A sophisticated, web-based production planning and costing application tailored for the Ayesha Abed Foundation (AAF) Gorpara branch. Built on Google Apps Script, it integrates seamlessly with Google Sheets and Google Drive to automate the lifecycle of production planning.

## 🚀 Key Features

-   **Dynamic Item Management:** Add multiple items per order, each with its own set of production processes.
-   **Automated Costing Engine:**
    -   **Real-time Totals:** Instant calculation of "Actual Production Costing" and "Section Payable Wages".
    -   **Intelligent Overhead Distribution:** Automatically distributes a user-defined overhead cost across all processes based on the payable wages ratio.
    -   **Grand Summary:** Displays a comprehensive financial breakdown including overhead percentages and grand totals.
-   **Professional PDF Generation:**
    -   Creates system-generated, branded PDF planning sheets.
    -   Includes formatted tables, meta-data, and summary sections.
    -   Automatically archives PDFs to a specific Google Drive folder.
-   **Centralized History & Tracking:**
    -   **Archive:** Every submission is logged into a "History Sheets" Google Sheet.
    -   **Status Management:** Track orders through "Pending" and "Complete" statuses directly from the web interface.
    -   **Search & Filter:** Quickly locate past plans by Order No, Date, or Notes using a real-time search interface.
-   **Data Consistency:**
    -   Prevents duplicate Order Numbers.
    -   Uses autocomplete for Item and Process names fetched from a central "settings" sheet.
    -   Synchronized server-side date handling.

## 🛠️ Technical Stack

-   **Backend:** Google Apps Script (`Code.gs`)
-   **Frontend:** HTML5, JavaScript (ES6+), Tailwind CSS for modern styling.
-   **Icons & Typography:** Font Awesome 6.4, Google Fonts (Inter).
-   **Database:** Google Sheets (Relational structure via multiple sheets).
-   **Storage:** Google Drive API for PDF management.

## 📁 Project Structure

-   `Code.gs`: Server-side logic handling Spreadsheet operations, PDF generation, and HTML rendering.
-   `index.html`: Modern, responsive single-page application (SPA) frontend.
-   `README.md`: Project documentation.

## 📄 License & Credits

© 2026 Ayesha Abed Foundation — Gorpara, Manikganj.

**Designed and Developed by:**
**Md Bashir Ahmed**, MIS, AAF Gorpara.
[LinkedIn Profile](https://www.linkedin.com/in/bashir-6923131b0/)
