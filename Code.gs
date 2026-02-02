// --- Global Helper Function to get Data ---
function getSheetData(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log("Sheet not found: " + sheetName);
      return [];
    }

    const numRows = sheet.getLastRow();
    
    // If only headers or empty, return an empty array
    if (numRows < 2) return [];

    // ✅ CORRECTION: Read only 5 columns (A, B, C, D, E) to ensure the data aligns 
    // with the calculations (Ref, Date, 8M, 9M, 11M)
    const dataRange = sheet.getRange(2, 1, numRows - 1, 5); 
    return dataRange.getValues();
  } catch(e) {
    Logger.log("Error reading sheet: " + sheetName + " - " + e.message);
    return [];
  }
}

// --- Global Helper Function to Calculate Totals ---
function calculateTotals(data) {
  // Based on the structure [Ref, Date, 8M, 9M, 11M]
  // 8M = index 2 (Column C)
  // 9M = index 3 (Column D)
  // 11M = index 4 (Column E)
  const totals = {'8M': 0, '9M': 0, '11M': 0};

  data.forEach(row => {
    // Use Number() instead of parseInt() for safer conversion
    totals['8M'] += Number(row[2]) || 0; 
    totals['9M'] += Number(row[3]) || 0;
    totals['11M'] += Number(row[4]) || 0;
  });
  return totals;
}


// --------------------------------------------------------------------------
// Function to handle GET requests (used by the Dashboard)
// --------------------------------------------------------------------------
function doGet(e) {
  try {
    // 1. Fetch data from the three sheets
    const bookedData = getSheetData('Booked');
    const receivedData = getSheetData('Received');
    const issuedData = getSheetData('Issued');

    // 2. Calculate totals for each pole size
    const bookedTotal = calculateTotals(bookedData);
    const receivedTotal = calculateTotals(receivedData);
    const issuedTotal = calculateTotals(issuedData);

    // 3. Perform dashboard calculations
    // Store Balance = Received - Booked
    const storeBalance8 = receivedTotal['8M'] - bookedTotal['8M'];
    const storeBalance9 = receivedTotal['9M'] - bookedTotal['9M'];
    const storeBalance11 = receivedTotal['11M'] - bookedTotal['11M'];

    // Ground Balance = Received - Issued
    const groundBalance8 = receivedTotal['8M'] - issuedTotal['8M'];
    const groundBalance9 = receivedTotal['9M'] - issuedTotal['9M'];
    const groundBalance11 = receivedTotal['11M'] - issuedTotal['11M'];

    // 4. Create the JSON response object
    const dashboardData = {
      storeBalance8: storeBalance8,
      storeBalance9: storeBalance9,
      storeBalance11: storeBalance11,
      groundBalance8: groundBalance8,
      groundBalance9: groundBalance9,
      groundBalance11: groundBalance11
    };

    // 5. Return the JSON data to the browser
    return ContentService.createTextOutput(JSON.stringify(dashboardData))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log("doGet Error: " + err.message);
    return ContentService.createTextOutput(JSON.stringify({ error: "Calculation failed: " + err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


// --------------------------------------------------------------------------
// Your Existing doPost (Cleaned up for consistency)
// --------------------------------------------------------------------------
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = JSON.parse(e.postData.contents);

    const type = data.type;   

    const ref = data.ref || "";
    const date = data.date || "";
    const q8 = Number(data.q8 || 0); 
    const q9 = Number(data.q9 || 0);
    const q11 = Number(data.q11 || 0);
    const consumer = data.consumer || "";
    const village = data.village || "";

    const sheet = ss.getSheetByName(type);
    if (!sheet) {
      return ContentService.createTextOutput("❌ Sheet Not Found: " + type);
    }

    if (type === "Booked" || type === "Received") {
      // Appends to A, B, C, D, E
      sheet.appendRow([ref, date, q8, q9, q11]);
    } else if (type === "Issued") {
      // Appends to A, B, C, D, E, F, G
      sheet.appendRow([ref, date, q8, q9, q11, consumer, village]);
    } else {
      return ContentService.createTextOutput("❌ Invalid Transaction Type");
    }

    return ContentService.createTextOutput("✅ Saved Successfully");

  } catch (err) {
    Logger.log("doPost Error: " + err.message);
    return ContentService.createTextOutput("❌ Error: " + err.message);
  }
}
