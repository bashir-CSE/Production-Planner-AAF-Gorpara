// --- CONFIGURATION ---
const SPREADSHEET_ID = "1AgRUT6d-XBhWj3V6w7Xu6sumJxlEwZ12m7oGHt8hvGU";
const DRIVE_FOLDER_ID = "1tu92dH38glARqlogXqftAeV8HG-8p9P4";

function doGet() {
	return HtmlService.createTemplateFromFile("index")
		.evaluate()
		.setTitle("Planning Sheet - Gorpara")
		.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Returns the current server date formatted for the planning date field.
 */
function getServerDate() {
	return Utilities.formatDate(
		new Date(),
		Session.getScriptTimeZone(),
		"MMMM d, yyyy",
	);
}

/**
 * Fetches Item and Process names from the Settings sheet for autocomplete.
 */
function getSettingsData() {
	const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
	const sheet = ss.getSheetByName("settings");
	if (!sheet) return { items: [], processes: [] };

	const lastRow = sheet.getLastRow();
	if (lastRow < 2) return { items: [], processes: [] };

	// Col A = Items, Col B = Processes
	const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

	const items = [...new Set(data.map((r) => r[0]).filter(String))];
	const processes = [...new Set(data.map((r) => r[1]).filter(String))];

	return { items, processes };
}

/**
 * Returns all records from the History Sheets for the Planning Sheets popup.
 * Columns: A=Planning Date, B=Order No, C=Planning Sheet Link, D=Status, E=Notes
 * Returns array of [dateStr, orderNo, status, notes, pdfUrl]
 */
function getPlanningSheets() {
	const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
	let sheet = ss.getSheetByName("History Sheets");

	if (!sheet) return [];

	const lastRow = sheet.getLastRow();
	if (lastRow < 2) return [];

	// Columns: A=Planning Date, B=Order No, C=PDF Link, D=Status, E=Notes
	const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

	return data.map((row) => {
		// Format date properly
		let dateStr = "";
		if (row[0] instanceof Date) {
			dateStr = Utilities.formatDate(
				row[0],
				Session.getScriptTimeZone(),
				"d/MMM/yyyy",
			);
		} else {
			dateStr = String(row[0] || "");
		}

		let orderNo = String(row[1] || "");

		return [
			dateStr, // [0] Planning Date (formatted string)
			orderNo, // [1] Order No
			String(row[3] || "Pending"), // [2] Status
			String(row[4] || ""), // [3] Notes
			String(row[2] || ""), // [4] PDF URL
		];
	});
}

/**
 * Updates the Status column (D) for a given Order No.
 */
function updateSheetStatus(orderNo, newStatus) {
	try {
		const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
		const sheet = ss.getSheetByName("History Sheets");
		if (!sheet) return { success: false, message: "Sheet not found" };

		const lastRow = sheet.getLastRow();
		if (lastRow < 2) return { success: false, message: "No records found" };

		const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
		for (let i = 0; i < data.length; i++) {
			if (String(data[i][1]) === String(orderNo)) {
				sheet.getRange(i + 2, 4).setValue(newStatus);
				return { success: true };
			}
		}
		return { success: false, message: "Order not found" };
	} catch (e) {
		return { success: false, message: e.toString() };
	}
}

/**
 * Handles the form submission: Generates PDF and saves to History Sheets.
 * Columns: A=Planning Date, B=Order No, C=Planning Sheet Link, D=Status, E=Notes
 */
function processFormSubmission(formData) {
	try {
		const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
		let sheet = ss.getSheetByName("History Sheets");

		// Create sheet if it doesn't exist and set up headers
		if (!sheet) {
			sheet = ss.insertSheet("History Sheets");
			sheet.appendRow([
				"Planning Date",
				"Order No",
				"Planning Sheet Link",
				"Status",
				"Notes",
			]);
			const headerRange = sheet.getRange(1, 1, 1, 5);
			headerRange.setFontWeight("bold");
			headerRange.setBackground("#1e3a8a");
			headerRange.setFontColor("#ffffff");
		}

		// 1. Check for duplicate Order No
		const finder = sheet
			.createTextFinder(formData.orderNo)
			.matchEntireCell(true);
		if (finder.findNext()) {
			return {
				success: false,
				message:
					'Order No "' +
					formData.orderNo +
					'" already exists. Please use a unique Order Number.',
			};
		}

		// 2. Generate PDF
		const pdfFile = createPDF(formData);

		// 3. Append the new data row
		const newRow = sheet.getLastRow() + 1;
		sheet.getRange(newRow, 1).setValue(formData.planningDate);
		sheet.getRange(newRow, 2).setValue(formData.orderNo).setNumberFormat("@");
		sheet.getRange(newRow, 3).setValue(pdfFile.getUrl());
		sheet.getRange(newRow, 4).setValue("Pending");
		sheet.getRange(newRow, 5).setValue(formData.remarks || "");

		return {
			success: true,
			message: "Saved successfully!",
			url: pdfFile.getUrl(),
		};
	} catch (e) {
		return { success: false, message: e.toString() };
	}
}

/**
 * Generates the PDF file in the specific Drive folder.
 * @param {Object} data - The form data object.
 */
function createPDF(data) {
	let tablesHtml = "";

	data.tables.forEach((table, index) => {
		let rowsHtml = "";
		table.rows.forEach((row, rowIndex) => {
			rowsHtml += `
        <tr>
          <td style="border: 1px solid #999; font-size: 12px; padding: 4px; text-align: center;">${rowIndex + 1}</td>
          <td style="border: 1px solid #999; font-size: 12px; padding: 4px; text-align: left;">${row.processName}</td>
          <td style="border: 1px solid #999; font-size: 12px; padding: 4px; text-align: center;">${row.actualCost}</td>
          <td style="border: 1px solid #999; font-size: 12px; padding: 4px; text-align: center;">${row.payableWages}</td>
          <td style="border: 1px solid #999; font-size: 12px; padding: 4px; text-align: center;">${row.overheadCosting}</td>
        </tr>
      `;
		});

		tablesHtml += `
      <div class="item-block">
        <h3 class="item-header">Item: ${table.itemName}</h3>
        <table class="data-table">
          <thead>
            <tr>
              <th style="width: 40px; text-align: center;">Sl. No</th>
              <th style="text-align: left;">Process Name</th>
              <th style="width: 100px; text-align: center;">Actual Costing</th>
              <th style="width: 100px; text-align: center;">Payable Wages</th>
              <th style="width: 100px; text-align: center;">Overhead Costing</th>
            </tr>
          </thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </div>
    `;
	});

	const generatedDate = Utilities.formatDate(
		new Date(),
		Session.getScriptTimeZone(),
		"dd MMM yyyy hh:mm:ss a",
	);

	const htmlTemplate = `
    <html>
      <head>
        <style>
            @page { size: A4; margin: 0.5cm; }
            body { 
              font-family: 'Inter', -apple-system, sans-serif; 
              padding: 0; 
              font-size: 11px; 
              color: #1e293b;
              line-height: 1.4;
            }
            h1 { font-size: 24px; margin: 0; color: #1e3a8a; letter-spacing: -0.02em; }
            h2 { font-size: 16px; margin: 4px 0 20px 0; color: #64748b; font-weight: 500; }
            
            .meta-table { width: 100%; margin-bottom: 20px; border-bottom: 2px solid #e2e8f0; padding-bottom: 15px; }
            .meta-table td { padding: 4px 0; }
            
            .item-block { margin-bottom: 25px; }
            .item-header { 
              background-color: #f8fafc; 
              padding: 8px 12px; 
              font-size: 15px; 
              margin: 0 0 -1px 0; 
              color: black;
            }
            
            .data-table { 
              width: 100%; 
              border-collapse: collapse; 
              font-size: 11px;
              table-layout: fixed;
            }
            .data-table thead { display: table-header-group; }
            .data-table th { 
              background-color: #f1f5f9; 
              border: 1px solid #cbd5e1; 
              padding: 8px; 
              color: black;
              font-weight: bold;
              text-transform: capitalize;
              letter-spacing: 0.04em;
              font-size: 12px;
            }
            .data-table td { 
              border: 1px solid #cbd5e1; 
              padding: 6px 8px; 
              word-wrap: break-word;
            }
            .data-table tr { page-break-inside: avoid; }

            .summary-section { 
              margin-top: 30px; 
              page-break-inside: avoid; 
              background: #f8fafc;
              border: 1px solid #e2e8f0;
              border-radius: 8px;
              padding: 15px;
            }
            .summary-title { margin: 0 0 12px 0; font-size: 14px; color: #1e3a8a; border-bottom: 1px solid #cbd5e1; padding-bottom: 5px; }
            .summary-table { width: 100%; border-collapse: collapse; }
            .summary-table td { padding: 5px 0; font-size: 11px; }

            .remarks-box { 
              margin-top: 20px; 
              page-break-inside: avoid; 
              padding: 12px; 
              border-left: 4px solid #indigo-500; 
              background: #fdfdfd; 
              border: 1px solid #e2e8f0;
              border-radius: 6px;
            }

            .footer-date { 
              position: fixed; 
              bottom: 5px; 
              left: 0; 
              width: 100%; 
              text-align: center; 
              color: black; 
              font-size: 10px; 
              padding-top: 2px;
            }
        </style>
      </head>
      <body>
        <div style="text-align: center; margin-bottom: 20px;">
          <h1>Ayesha Abed Foundation - Gorpara</h1>
          <h2>Planning Sheet</h2>
        </div>
        
        <table class="meta-table">
          <tr>
            <td style="font-size: 13px;"><strong>Order Number:</strong> ${data.orderNo}</td>
            <td style="text-align: right; color: #64748b;"><strong>Date:</strong> ${data.planningDate}</td>
          </tr>
        </table>

        ${tablesHtml}

        ${
					data.remarks
						? `
        <div class="remarks-box">
          <h3 style="font-size: 12px; margin: 0 0 8px 0; color: #1e293b;">Notes / Special Instructions</h3>
          <p style="margin: 0; color: #475569; white-space: pre-wrap;">${data.remarks}</p>
        </div>`
						: ""
				}

        <div class="summary-section">
          <h3 class="summary-title">Planning Summary Totals</h3>
          <table class="summary-table">
             <tr>
               <td><strong>Total Actual Production Costing</strong></td>
               <td style="text-align: right; font-weight: bold; font-size: 12px;">${data.grandTotalActual}</td>
             </tr>
             <tr>
               <td><strong>Total Section Payable Wages</strong></td>
               <td style="text-align: right; font-weight: bold; font-size: 12px;">${data.grandTotalWages}</td>
             </tr>
             <tr>
               <td style="padding-bottom: 8px;"><strong>Calculated Overhead Costing</strong></td>
               <td style="text-align: right; font-weight: bold;">${data.grandTotalOverhead}</td>
             </tr>
             <tr>
               <td><strong>Total Overhead Cost %</strong></td>
               <td style="text-align: right; font-weight: bold; font-size: 12px;">${data.overheadPercent}%</td>
             </tr>
             </tr>
          </table>
        </div>

        <div class="footer-date">
            This document was system generated on ${generatedDate} • MIS, AAF - Gorpara
        </div>
      </body>
    </html>
  `;

	const blob = Utilities.newBlob(htmlTemplate, "text/html", "temp.html");
	const pdfBlob = blob.getAs("application/pdf").setName(`${data.orderNo}.pdf`);
	const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
	return folder.createFile(pdfBlob);
}
