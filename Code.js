/**
 * Generates PDF reports for each brand from a CSV file in Google Drive.
 * The CSV should have columns: ContactName, ..., InvoiceNumber, InvoiceDate, DueDate, ..., Description, Quantity, UnitAmount, ..., TaxAmount, ..., gross
 * Groups data by ContactName (brand), then by InvoiceNumber within brand, creates one PDF per brand with invoice details.
 * Uses an external HTML template file for PDF formatting.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Reports')
    .addItem('Generate Brand Reports', 'generateBrandReports')
    .addToUi();
}

/**
 * Returns Sheet1 and sets up progress tracking area.
 * Creates a dynamic log section in columns A-E starting from row 1.
 */
// Global constants
const LOG_START_ROW = 28;
const TIMESTAMP_FORMAT = 'HH:mm:ss';

function getOrCreateProgressSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Sheet1');
  if (!sheet) {
    // If Sheet1 doesn't exist, create it
    sheet = ss.insertSheet('Sheet1');
  }
  
  // Do not create column headers; leave rows 2..(LOG_START_ROW-1) free for notes or small UI elements.
  // Clear any previous log entries in the log area starting at LOG_START_ROW
  const lastRow = sheet.getLastRow();
  return sheet;
}

/**
 * Return the next writable log row at or after LOG_START_ROW.
 */
function getNextLogRow(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < LOG_START_ROW) return LOG_START_ROW;

  // If the cell at LOG_START_ROW is empty, use that, otherwise find the first empty cell below lastRow
  const firstCheck = sheet.getRange(LOG_START_ROW, 1).getValue();
  if (!firstCheck) return LOG_START_ROW;

  // Otherwise return lastRow + 1
  return lastRow + 1;
}

/**
 * Remove emoji characters from a string so the progress log stays plain text.
 */
function sanitizeText(str) {
  if (!str) return '';
  // Basic emoji removal using Unicode ranges for emoji and symbols
  return str.replace(/[\p{Emoji_Presentation}\p{Emoji}\uFE0F]/gu, '').replace(/[\u2600-\u27BF]/g, '');
}

/**
 * Appends a progress entry to Sheet1 with live updates.
 * Uses SpreadsheetApp.flush() to ensure the UI reflects updates during long runs.
 */
function updateProgressSheet(ss, brand, invoicesCount, itemsCount, status) {
  try {
  const sheet = getOrCreateProgressSheet();

  // Compose a simple single-line message
  const text = sanitizeText(`${status}`);

  // Choose the next writable log row (ensures we start at LOG_START_ROW)
  const writeRow = getNextLogRow(sheet);
  sheet.getRange(writeRow, 1).setValue(text);
    
    // Auto-resize columns for better visibility
    sheet.autoResizeColumns(1, 5);
    
    // Force pending changes to be applied so users see live updates.
    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log('Failed to update progress sheet: ' + e.message);
  }
}

/**
 * Updates the progress with a general status message (not brand-specific)
 */
function updateProgressStatus(ss, message) {
  try {
  const sheet = getOrCreateProgressSheet();

  const text = sanitizeText(`${message}`);
  const writeRow = getNextLogRow(sheet);
  sheet.getRange(writeRow, 1).setValue(text);
    sheet.autoResizeColumns(1, 5);
    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log('Failed to update progress status: ' + e.message);
  }
}

/**
 * Clears the progress log area starting from LOG_START_ROW to the last row.
 */
function clearProgressLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateProgressSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < LOG_START_ROW) {
    // Nothing to clear
    SpreadsheetApp.getUi().alert('No progress log entries to clear.');
    return;
  }

  // Clear the range from LOG_START_ROW to lastRow in column A (single column log)
  sheet.getRange(LOG_START_ROW, 1, lastRow - LOG_START_ROW + 1, 5).clearContent();
  sheet.autoResizeColumns(1, 5);
  SpreadsheetApp.flush();

}

// writeUserInstructions removed

function generateBrandReports() {
  // Get spreadsheet reference early so it can be used throughout the function
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Clear progress log at the start
  clearProgressLog();
  
  // Initialize progress tracking
  updateProgressStatus(ss, 'Starting brand report generation...');
  
  // Prompt user to enter Google Drive URL of the CSV file
  const csvUrl = Browser.inputBox('Enter the Google Drive URL of the CSV file:');
  if (csvUrl === 'cancel') {
    updateProgressStatus(ss, 'Operation cancelled by user');
    return;
  }

  // Extract CSV file ID from URL
  const csvMatch = csvUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!csvMatch) {
    Browser.msgBox('Invalid CSV URL. Please provide a valid Google Drive file URL.');
    return;
  }
  const csvFileId = csvMatch[1];

  // Load HTML template from project file `invoice_template.html` (no prompt)
  // The template should be part of the Apps Script project as a file named `invoice_template.html`.

  try {
  updateProgressStatus(ss, 'Loading HTML template...');
    
    // Load HTML template from project files (preferred) or Drive fallback
    let htmlTemplate = '';
    try {
      // First try to load the template from the Apps Script project (preferred).
      try {
        htmlTemplate = HtmlService.createHtmlOutputFromFile('invoice_template').getContent();
          // Removed verbose log: updateProgressStatus(ss, 'Loaded HTML template from project file');
        Logger.log('Loaded HTML template from project file: invoice_template.html');
      } catch (projErr) {
        // Removed verbose log: updateProgressStatus(ss, 'Template not in project, checking Drive...');
        const fileIter = DriveApp.getFilesByName('invoice_template.html');
        if (fileIter.hasNext()) {
          const f = fileIter.next();
          htmlTemplate = f.getBlob().getDataAsString();
          // Removed verbose log: updateProgressStatus(ss, 'Loaded HTML template from Drive');
          Logger.log('Loaded HTML template from Drive file: ' + f.getName());
        } else {
          throw new Error('Could not find invoice_template.html in project or Drive.');
        }
      }
    } catch (loadErr) {
      throw loadErr;
    }
    // Get the CSV file - handle CSV files, Excel files, and Google Sheets
  updateProgressStatus(ss, 'Loading data file...');
    const file = DriveApp.getFileById(csvFileId);
    const mimeType = file.getMimeType();
    let data;
    
  updateProgressStatus(ss, `Processing ${file.getName()}`);
    Logger.log('File MIME type: ' + mimeType);
    
    if (mimeType === 'application/vnd.google-apps.spreadsheet') {
      // It's a Google Sheet
      const sheet = SpreadsheetApp.openById(csvFileId).getActiveSheet();
      data = sheet.getDataRange().getValues();
    } else if (mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
               mimeType === 'application/vnd.ms-excel' ||
               file.getName().toLowerCase().endsWith('.xlsx') ||
               file.getName().toLowerCase().endsWith('.xls')) {
      // It's an Excel file - convert to Google Sheet first
      Browser.msgBox('Excel files need to be converted to Google Sheets first. Please upload your file to Google Drive and convert it to Google Sheets format, then use that URL.');
      return;
    } else {
      // It's a CSV file
      let csvData = file.getBlob().getDataAsString();
      csvData = csvData.replace(/^\uFEFF/, ''); // Remove BOM if present
      
      // Try different parsing approaches
      try {
        data = Utilities.parseCsv(csvData);
      } catch (parseError) {
        Logger.log('Standard CSV parse failed, trying manual split');
        // Fallback to manual parsing
        const lines = csvData.split(/\r?\n/);
        data = lines.map(line => line.split(','));
      }
    }
    
    const headers = data[0].map(h => h.toString().trim());

    // Validate required headers exist (allow common alternate names)
    const headerSet = new Set(headers.map(h => h.toLowerCase()));
    const required = [
      '*contactname', 'contactname',
      '*invoicenumber', 'invoicenumber',
      'description',
      '*quantity', 'quantity',
      '*unitamount', 'unitamount',
      'taxamount',
      'gross'
    ];

    const hasRequired = required.some(r => headerSet.has(r));
    // A looser check: ensure at least one of contact name, invoice number, description, and gross exist
    const basicOk = (headerSet.has('*contactname') || headerSet.has('contactname'))
      && (headerSet.has('*invoicenumber') || headerSet.has('invoicenumber'))
      && headerSet.has('description')
      && (headerSet.has('*quantity') || headerSet.has('quantity'))
      && (headerSet.has('*unitamount') || headerSet.has('unitamount'))
      && headerSet.has('taxamount');

    if (!basicOk) {
      updateProgressStatus(ss, 'Headers on sheet not as expected');
      return;
    }
    const rows = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = row[i] ? row[i].toString().trim() : '';
      });
      return obj;
    });

    Logger.log('Headers: ' + headers.join(', '));
    Logger.log('First row data: ' + JSON.stringify(rows[0]));

    // Group by ContactName (brand) and then by *InvoiceNumber within brand
    const brands = {};
    rows.forEach(row => {
      const brand = row['*ContactName'] || row[headers[0]] || 'Unknown Brand';
      const invoiceNum = row['*InvoiceNumber'] || row['InvoiceNumber'] || 'Unknown Invoice';
      Logger.log('Processing row with brand: ' + brand + ', invoice: ' + invoiceNum);
      if (!brands[brand]) {
        brands[brand] = {};
      }
      if (!brands[brand][invoiceNum]) {
        brands[brand][invoiceNum] = [];
      }
      brands[brand][invoiceNum].push(row);
    });

  updateProgressStatus(ss, 'Starting PDF generation for each brand...');

    // For each brand, create a PDF
    const pdfResults = Object.keys(brands).map(brand => {
      const brandData = brands[brand];
      
      // NEW LOGIC: Group all items in the brand by order number instead of by invoice
      const allItems = Object.values(brandData).flat();
      
      // Group by order number from Description field
      const itemsByOrderNumber = {};
      allItems.forEach(item => {
        const desc = item['Description'] || '';
        const orderMatch = desc.match(/^#(\d+)/);
        const orderNum = orderMatch ? orderMatch[1] : 'Unknown';
        
        if (!itemsByOrderNumber[orderNum]) {
          itemsByOrderNumber[orderNum] = [];
        }
        itemsByOrderNumber[orderNum].push(item);
      });
      
      // Calculate totals for all items in the brand
      const subtotal = allItems.reduce((sum, row) => {
        const quantity = parseFloat(row['*Quantity'] || row['Quantity'] || '1');
        const unitPrice = parseFloat(row['*UnitAmount'] || row['UnitAmount'] || 0);
        const taxType = row['*TaxType'] || '';
        
        let netAmount;
        if (taxType.toLowerCase().includes('no vat')) {
          netAmount = roundToNearest(unitPrice * quantity);
        } else {
          netAmount = roundToNearest((unitPrice * quantity) / 1.2);
        }
        
        return sum + netAmount;
      }, 0);
      
      const totalVat = allItems.reduce((sum, row) => {
        const quantity = parseFloat(row['*Quantity'] || row['Quantity'] || '1');
        const unitPrice = parseFloat(row['*UnitAmount'] || row['UnitAmount'] || 0);
        const taxType = row['*TaxType'] || '';
        
        let vatAmount = 0;
        if (!taxType.toLowerCase().includes('no vat')) {
          const grossAmount = unitPrice * quantity;
          const netAmount = roundToNearest(grossAmount / 1.2);
          vatAmount = roundToNearest(grossAmount - netAmount);
        }
        
        return sum + vatAmount;
      }, 0);
      
      const totalGbp = allItems.reduce((sum, row) => {
        const quantity = parseFloat(row['*Quantity'] || row['Quantity'] || '1');
        const unitPrice = parseFloat(row['*UnitAmount'] || row['UnitAmount'] || 0);
        return sum + roundToNearest(unitPrice * quantity);
      }, 0);
      
      // Update progress sheet
      const totalItems = allItems.length;
      updateProgressSheet(ss, brand, Object.keys(brandData).length, totalItems, `Processing brand: ${brand}`);
      Logger.log(`Processing brand: ${brand} with ${Object.keys(brandData).length} invoices and ${totalItems} items.`);
      
      const clientName = brand;
      const currentDate = Utilities.formatDate(new Date(), 'GMT', 'dd/MM/yyyy');
      const allInvoiceNumbers = Object.keys(brandData);
      
      // Define styles for consistency
      const thStyle = "background-color: #f7f7f7; color: #333333; padding: 12px; border-bottom: 2px solid #dddddd; text-align: left; font-size: 12px; font-weight: bold; text-transform: uppercase;";
      const tdStyle = "padding: 12px; border-bottom: 1px solid #eeeeee;";
      const tdNumberStyle = `${tdStyle} text-align: right;`;
      
      // Function to render a single item row
      const renderRow = (row) => {
        const quantity = parseFloat(row['*Quantity'] || row['Quantity'] || '1');
        const unitPrice = parseFloat(row['*UnitAmount'] || row['UnitAmount'] || 0);
        const taxType = row['*TaxType'] || '';
        
        let netAmount;
        let vatRate = '20%';
        
        if (taxType.toLowerCase().includes('no vat')) {
          netAmount = roundToNearest(unitPrice * quantity);
          vatRate = 'N/A';
        } else {
          netAmount = roundToNearest((unitPrice * quantity) / 1.2);
        }
        
        const description = row['Description'] || '';
        const formatAmount = (amount) => (amount < 0) ? `(${Math.abs(amount).toFixed(2)})` : `${amount.toFixed(2)}`;
        
        return `
        <tr>
          <td style="${tdStyle}">${description}</td>
          <td style="${tdNumberStyle}">${quantity.toFixed(2)}</td>
          <td style="${tdNumberStyle}">${unitPrice.toFixed(2)}</td>
          <td style="${tdNumberStyle}">${vatRate}</td>
          <td style="${tdNumberStyle}">${formatAmount(netAmount)}</td>
        </tr>`;
      };
      
      // Render all items grouped by order number
      let itemRowsHtml = '';
      const sortedOrderNumbers = Object.keys(itemsByOrderNumber).sort();
      
      sortedOrderNumbers.forEach(orderNum => {
        const orderItems = itemsByOrderNumber[orderNum];
        const shippingRegex = /ship|post|postage|delivery|freight|shipping/i;
        
        const discounts = orderItems.filter(r => {
          const d = (r['Description'] || '').toString().toLowerCase();
          const quantity = parseFloat(r['*Quantity'] || r['Quantity'] || '1');
          return d.includes('discount') || quantity < 0;
        });
        
        const normal = orderItems.filter(r => {
          const d = (r['Description'] || '').toString().toLowerCase();
          const quantity = parseFloat(r['*Quantity'] || r['Quantity'] || '1');
          const isDiscount = d.includes('discount') || quantity < 0;
          const isShipping = shippingRegex.test(r['Description'] || '');
          return !isDiscount && !isShipping;
        });
        
        const shipping = orderItems.filter(r => shippingRegex.test(r['Description'] || ''));
        
        // Render in order: discounts, normal, shipping for this order
        itemRowsHtml += discounts.map(renderRow).join('');
        itemRowsHtml += normal.map(renderRow).join('');
        itemRowsHtml += shipping.map(renderRow).join('');
      });
      
      // Single table for the entire brand
      const invoicesHtml = `
        <!-- Items Table -->
        <table width="100%" border="0" cellspacing="0" cellpadding="0" style="font-size: 13px;">
          <thead>
            <tr>
              <th style="${thStyle}">Description</th>
              <th style="${thStyle} text-align: right;">Quantity</th>
              <th style="${thStyle} text-align: right;">Unit Price</th>
              <th style="${thStyle} text-align: right;">VAT</th>
              <th style="${thStyle} text-align: right;">Net Amount GBP</th>
            </tr>
          </thead>
          <tbody>
            ${itemRowsHtml}
          </tbody>
        </table>

        <!-- Totals Section -->
        <table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top: 20px;">
          <tr>
            <td align="right">
              <table border="0" cellspacing="0" cellpadding="5" style="width: 280px; font-size: 13px;">
                <tr>
                  <td style="padding: 10px; text-align: right;">Subtotal</td>
                  <td style="padding: 10px; text-align: right;">£${subtotal.toFixed(2)}</td>
                </tr>
                <tr>
                  <td style="padding: 10px; text-align: right;">VAT (20%)</td>
                  <td style="padding: 10px; text-align: right;">£${totalVat.toFixed(2)}</td>
                </tr>
                <tr>
                  <td style="padding: 10px; text-align: right; background-color: #f7f7f7; font-weight: bold; border-top: 2px solid #dddddd;">Total GBP</td>
                  <td style="padding: 10px; text-align: right; background-color: #f7f7f7; font-weight: bold; border-top: 2px solid #dddddd;">£${totalGbp.toFixed(2)}</td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      `;
      
      // Replace placeholders in the template
      const htmlContent = htmlTemplate
        .replace(/\{\{BRAND\}\}/g, brand)
        .replace(/\{\{CLIENT_NAME\}\}/g, clientName)
        .replace(/\{\{CURRENT_DATE\}\}/g, currentDate)
        .replace(/\{\{INVOICE_NUMBERS\}\}/g, allInvoiceNumbers.join(', '))
        .replace(/\{\{INVOICES_HTML\}\}/g, invoicesHtml);
    
      // Convert HTML to PDF
      const pdfBlob = Utilities.newBlob(htmlContent, 'text/html').setName(`Invoice_${brand}.pdf`);
      const pdf = pdfBlob.getAs('application/pdf');
      
  // Removed verbose log: updateProgressSheet(ss, brand, Object.keys(brandData).length, totalItems, `PDF created successfully`);
      Logger.log(`PDF created for ${brand}.`);
      
      return {
        brand: brand,
        invoiceNumbers: allInvoiceNumbers,
        pdfBlob: pdf,
        fileName: `Invoice_${brand}.pdf`
      };
    }).filter(result => result !== null);

  // Create timestamp folder and organize PDFs
  // Inform the user that Drive operations (creating files/folders) may take some time
    updateProgressStatus(ss, 'Creating Drive folder and organizing PDFs... (this may take a few minutes for Drive to finish creating files)');
    const now = new Date();
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd-HH:mm');
    const folderName = `CSV_TO_PDF_${timestamp}`;
    const folder = DriveApp.createFolder(folderName);
    
    // Save PDFs to folder and collect file info
    const fileLinks = [];
    pdfResults.forEach(result => {
      const pdfFile = DriveApp.createFile(result.pdfBlob).setName(result.fileName);
      folder.addFile(pdfFile);
      DriveApp.getRootFolder().removeFile(pdfFile); // Remove from root, keep only in folder
      
      fileLinks.push({
        brand: result.brand,
        invoiceNumbers: result.invoiceNumbers.join(', '),
        fileName: result.fileName,
        fileUrl: pdfFile.getUrl()
      });
    });

    const sheetName = timestamp;
    const trackingSheet = ss.insertSheet(sheetName);
    
    // Set up headers
    trackingSheet.getRange(1, 1, 1, 4).setValues([['Brand', 'Invoice Numbers', 'File Name', 'File Link']]);
    trackingSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#4CAF50').setFontColor('white');
    
    // Add data with hyperlinks
    fileLinks.forEach((link, index) => {
      const row = index + 2;
      trackingSheet.getRange(row, 1).setValue(link.brand);
      trackingSheet.getRange(row, 2).setValue(link.invoiceNumbers);
      trackingSheet.getRange(row, 3).setValue(link.fileName);
      trackingSheet.getRange(row, 4).setFormula(`=HYPERLINK("${link.fileUrl}", "Open PDF")`);
    });
    
    // Auto-resize columns
    trackingSheet.autoResizeColumns(1, 4);
    
  updateProgressStatus(ss, `COMPLETED: ${fileLinks.length} PDFs created successfully`);
  Browser.msgBox(`Reports generated successfully! ${fileLinks.length} PDFs created in folder "${folderName}" and tracking sheet "${sheetName}" added.`);
  } catch (e) {
  updateProgressStatus(ss, `ERROR: ${e.message}`);
  Browser.msgBox('Error: ' + e.message);
  }
}

// Add this helper function at the top of the file
function roundToNearest(value, decimals = 2) {
  return Math.round(value * Math.pow(10, decimals)) / Math.pow(10, decimals);
}
