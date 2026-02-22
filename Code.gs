// --- CONFIGURATION ---
const SPREADSHEET_ID = '1AgRUT6d-XBhWj3V6w7Xu6sumJxlEwZ12m7oGHt8hvGU'; 
const DRIVE_FOLDER_ID = '1tu92dH38glARqlogXqftAeV8HG-8p9P4'; 

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Production Planning Sheet')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Fetches Item and Process names from the Settings sheet for autocomplete.
 */
function getSettingsData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('settings');
  if (!sheet) return { items: [], processes: [] };
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { items: [], processes: [] };

  // Col A = Items, Col B = Processes
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  
  const items = [...new Set(data.map(r => r[0]).filter(String))];
  const processes = [...new Set(data.map(r => r[1]).filter(String))];
  
  return { items, processes };
}

/**
 * Returns all records from the Documents sheet for the Planning Sheets popup.
 * Returns array of [planningDate, orderNo, done(boolean), pdfUrl]
 */
function getPlanningSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let docSheet = ss.getSheetByName('Documents');

  if (!docSheet) return [];

  const lastRow = docSheet.getLastRow();
  if (lastRow < 2) return [];

  // Columns: A=Planning Date, B=Order No, C=PDF Link, D=Done (checkbox)
  const data = docSheet.getRange(2, 1, lastRow - 1, 4).getValues();

  return data.map(row => {
    // Format date properly — getValues() returns JS Date objects for date cells
    let dateStr = '';
    if (row[0] instanceof Date) {
      dateStr = Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'MMMM d, yyyy');
    } else {
      dateStr = String(row[0] || '');
    }

    // Clean order number — remove leading apostrophe if present
    let orderNo = String(row[1] || '').replace(/^'/, '');

    return [
      dateStr,            // Planning Date (formatted string)
      orderNo,            // Order No
      row[3] === true,    // Done checkbox (strict boolean)
      String(row[2] || '') // PDF URL
    ];
  });
}

/**
 * Handles the form submission: Generates PDF and saves to Documents Sheet.
 * Adds a checkbox (FALSE by default) in column D for each new row.
 */
function processFormSubmission(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let docSheet = ss.getSheetByName('Documents');
    
    // Create sheet if it doesn't exist and set up headers
    if (!docSheet) {
      docSheet = ss.insertSheet('Documents');
      docSheet.appendRow(['Planning Date', 'Order No', 'Planning Sheet Link', 'Done']);
      const headerRange = docSheet.getRange(1, 1, 1, 4);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#1e3a8a');
      headerRange.setFontColor('#ffffff');
    }

    // 1. Check for duplicate Order No
    const finder = docSheet.createTextFinder(formData.orderNo).matchEntireCell(true);
    if (finder.findNext()) {
      return { success: false, message: 'Order No "' + formData.orderNo + '" already exists. Please use a unique Order Number.' };
    }

    // 2. Generate PDF
    const pdfFile = createPDF(formData);

    // 3. Append the new data row (A: Date, B: OrderNo, C: PDF URL, D: Checkbox FALSE)
    const newRow = docSheet.getLastRow() + 1;
    docSheet.getRange(newRow, 1).setValue(formData.planningDate);
    docSheet.getRange(newRow, 2).setValue("'" + formData.orderNo);
    docSheet.getRange(newRow, 3).setValue(pdfFile.getUrl());
    
    // Insert an actual Google Sheets CHECKBOX in column D
    const checkboxCell = docSheet.getRange(newRow, 4);
    checkboxCell.insertCheckboxes();
    checkboxCell.setValue(false);

    return { success: true, message: 'Saved successfully!', url: pdfFile.getUrl() };

  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Generates the PDF file in the specific Drive folder.
 * @param {Object} data - The form data object.
 */
function createPDF(data) {
  let tablesHtml = '';

  data.tables.forEach((table, index) => {
    let rowsHtml = '';
    table.rows.forEach((row, rowIndex) => {
      rowsHtml += `
        <tr>
          <td style="border: 1px solid #999; padding: 4px; text-align: center;">${rowIndex + 1}</td>
          <td style="border: 1px solid #999; padding: 4px;">${row.processName}</td>
          <td style="border: 1px solid #999; padding: 4px; text-align: right;">${row.actualCost}</td>
          <td style="border: 1px solid #999; padding: 4px; text-align: right;">${row.payableWages}</td>
          <td style="border: 1px solid #999; padding: 4px; text-align: right;">${row.overheadCosting}</td>
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
              <th style="width: 80px; text-align: right;">Actual Costing</th>
              <th style="width: 80px; text-align: right;">Payable Wages</th>
              <th style="width: 80px; text-align: right;">Overhead Costing</th>
            </tr>
          </thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </div>
    `;
  });

  const generatedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy hh:mm:ss a");

  const htmlTemplate = `
    <html>
      <head>
        <style>
            @page { size: A4; margin: 1cm; }
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
              border: 1px solid #e2e8f0; 
              font-size: 13px; 
              margin: 0 0 -1px 0; 
              color: #1e293b;
              border-top-left-radius: 4px;
              border-top-right-radius: 4px;
            }
            
            .data-table { 
              width: 100%; 
              border-collapse: collapse; 
              font-size: 10px;
              table-layout: fixed;
            }
            .data-table thead { display: table-header-group; } /* Repeats header on new pages */
            .data-table th { 
              background-color: #f1f5f9; 
              border: 1px solid #cbd5e1; 
              padding: 8px; 
              color: #475569;
              font-weight: 600;
              text-transform: uppercase;
              letter-spacing: 0.05em;
              font-size: 9px;
            }
            .data-table td { 
              border: 1px solid #cbd5e1; 
              padding: 6px 8px; 
              word-wrap: break-word;
            }
            .data-table tr { page-break-inside: avoid; } /* Prevents clipping rows */

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
              bottom: -10px; 
              left: 0; 
              width: 100%; 
              text-align: center; 
              color: #94a3b8; 
              font-size: 9px; 
              padding-top: 10px;
              border-top: 1px solid #f1f5f9;
            }
        </style>
      </head>
      <body>
        <div style="text-align: center; margin-bottom: 25px;">
          <h1>Ayesha Abed Foundation - Gorpara</h1>
          <h2>Production Planning Official Document</h2>
        </div>
        
        <table class="meta-table">
          <tr>
            <td style="font-size: 13px;"><strong>Order Reference:</strong> <span style="color: #1e3a8a;">${data.orderNo}</span></td>
            <td style="text-align: right; color: #64748b;"><strong>Document Issued:</strong> ${data.planningDate}</td>
          </tr>
          <tr>
            <td colspan="2" style="text-align: right; font-size: 10px; color: #94a3b8;">
                Overhead Logic: <strong>৳ ${data.overheadCostInput}</strong> derived at <strong>${data.overheadPercent}%</strong> scaling.
            </td>
          </tr>
        </table>

        ${tablesHtml}

        ${data.remarks ? `
        <div class="remarks-box">
          <h3 style="font-size: 12px; margin: 0 0 8px 0; color: #1e293b;">Notes / Special Instructions</h3>
          <p style="margin: 0; color: #475569; white-space: pre-wrap;">${data.remarks}</p>
        </div>` : ''}

        <div class="summary-section">
          <h3 class="summary-title">Planning Summary Totals</h3>
          <table class="summary-table">
             <tr>
               <td><strong>Total Actual Production Costing</strong></td>
               <td style="text-align: right; font-weight: bold; font-size: 12px;">৳ ${data.grandTotalActual}</td>
             </tr>
             <tr>
               <td><strong>Total Section Payable Wages</strong></td>
               <td style="text-align: right; font-weight: bold; font-size: 12px;">৳ ${data.grandTotalWages}</td>
             </tr>
             <tr>
               <td style="border-bottom: 1px dashed #cbd5e1; padding-bottom: 8px;"><strong>Calculated Overhead Costing</strong></td>
               <td style="text-align: right; font-weight: bold; border-bottom: 1px dashed #cbd5e1; padding-bottom: 8px;">৳ ${data.grandTotalOverhead}</td>
             </tr>
             <tr style="color: #1e3a8a;">
                <td style="padding-top: 10px; font-size: 13px;"><strong>GRAND TOTAL PROJECTED COST</strong></td>
                <td style="padding-top: 10px; text-align: right; font-size: 15px; font-weight: 800;">
                    ৳ ${(parseFloat(data.grandTotalActual) + parseFloat(data.grandTotalWages) + parseFloat(data.grandTotalOverhead)).toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                </td>
             </tr>
          </table>
        </div>

        <div class="footer-date">
            System generated archive document • Generated on ${generatedDate} • AAF MIS Infra
        </div>
      </body>
    </html>
  `;

  const blob = Utilities.newBlob(htmlTemplate, 'text/html', 'temp.html');
  const pdfBlob = blob.getAs('application/pdf').setName(`${data.orderNo}.pdf`);
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  return folder.createFile(pdfBlob);
}