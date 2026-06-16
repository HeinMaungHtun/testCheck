function doGet() {
  return HtmlService
    .createTemplateFromFile('Index') 
    .evaluate()
    .setTitle('Master Data Checklist');
}

const file_url = "https://docs.google.com/spreadsheets/d/1kIYcn2a43JBvkDgNVEBJVJmg_qm7arr8atgv9u66A8c/edit?usp=sharing";
const excel_file = SpreadsheetApp.openByUrl(file_url)
const tasks_data = excel_file.getSheetByName('Tasks').getDataRange().getValues()

function getCurrentWeek() {
  const today = new Date();
  const firstDay = new Date(today.getFullYear(), 0, 1);
  const dayOfYear = Math.floor((today - firstDay) / 86400000);
  const week = Math.ceil((dayOfYear + firstDay.getDay() + 1) / 7);
  return today.getFullYear() + "-W" + week;
}

function getCompanies() {
  const sheet = excel_file.getSheetByName("Companies");
  const data = sheet.getDataRange().getValues();
  const companies = [...new Set(data.slice(1).map(r => r[0]))];
  console.log(companies);
  return companies;
}

function getCategories() {
  const sheet = excel_file.getSheetByName("Categories");
  const data = sheet.getDataRange().getValues();
  const categories = [...new Set(data.slice(1).map(r => r[1]))];
  console.log(categories);
  return categories;
}

function getTasks(company,category) {
  const sheet = excel_file.getSheetByName("latestdata");;
  const data = sheet.getDataRange().getValues();

  let result = [];

  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    let rowId = row[0];
    let rowCompany = row[1];
    let rowCategory = row[2];    
    let rowtaskId = row[3];
    let rowtaskName = row[4];
    let rowStatus = row[5];
    let rowlastUpdated = row[6];

    if ( rowCompany === company && rowCategory === category) {
      result.push({
        rowId:rowId,
        company : rowCategory,
        category : rowCategory,                
        taskId: rowtaskId,
        task: rowtaskName,
        status : rowStatus,
        lastUpdated: rowlastUpdated
      });
    }
  };

  return result;
}

function findExistingRecord(rowId) {
  const sheet = excel_file.getSheetByName("latestdata");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if ( row[0] === rowId) {
      return i + 1; // Return row number (1-based for Sheets API)
    }
  }
  return null;
}

function saveTasks(payload) {
  const historySheet = excel_file.getSheetByName("data"); // Timeline/History sheet
  const latestSheet = excel_file.getSheetByName("latestdata"); // Latest status sheet
  const dateString = payload.date; // Already formatted as YYYY-MM-DD
  const timestamp = new Date(); // Current timestamp
  
  const results = {
    inserted: 0,
    updated: 0,
    errors: []
  };

  try {
    payload.tasks.forEach(t => {
      try {
        // Step 1: Always INSERT to history sheet (data) for Timeline tracking
        historySheet.appendRow([
          t.rowId,          
          payload.company,
          payload.category,
          'test',
          'testdesc',
          t.status,          
          dateString,
          timestamp,
          new Date().toLocaleString() // Human-readable timestamp
        ]);
        results.inserted++;
        
        // Step 2: UPDATE latestdata sheet for current status display
        const latestRowNumber = findExistingRecord(t.rowId);
        if (latestRowNumber) {
          latestSheet.getRange(latestRowNumber, 5).setValue(t.status); // Update status
          latestSheet.getRange(latestRowNumber, 6).setValue(dateString); // Update date
          latestSheet.getRange(latestRowNumber, 7).setValue(timestamp); // Update timestamp
          results.updated++;
        }
      } catch (e) {
        results.errors.push("Error processing task " + t.rowId + ": " + e.message);
      }
    });

    Logger.log("Save completed - History Inserted: " + results.inserted + ", Latest Updated: " + results.updated);
    return results;
  } catch (e) {
    throw new Error("Failed to save tasks: " + e.message);
  }
}
