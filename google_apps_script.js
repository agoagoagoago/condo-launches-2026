// Google Apps Script - Deploy this as a Web App in Google Sheets
// This receives form submissions and adds them to the spreadsheet

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Parse the form data
    var data = JSON.parse(e.postData.contents);

    var date = new Date().toLocaleDateString('en-SG', {
      year: 'numeric',
      month: 'short',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    });

    var name = data.name || '';
    var email = data.email || '';
    var projects = data.projects || []; // Array of selected projects

    // All project names (must match exactly with form values)
    var allProjects = [
      'Narra Residences',
      'Newport Residences',
      'Duet @ Emily',
      'Sophia Meadows',
      'Pinery Residences',
      'Tengah Garden Avenue',
      'River Modern',
      'Bayshore Road',
      'Media Circle (Parcel A)',
      'Lentor Gardens',
      'Dunearn Road',
      'Holland Link',
      'Lakeside Drive',
      'Chencharu Close',
      'Chuan Grove',
      'Dorset Road',
      'Former Thomson View condo',
      'Upper Thomson Road (Parcel A)',
      'Coastal Cabana (EC)',
      'Rivelle Tampines (EC)',
      'Woodlands Drive (EC)',
      'Sembawang Road (EC)',
      'Senja Close (EC)'
    ];

    // Build the row: Date, Name, Email, then email under each selected project
    var row = [date, name, email];

    allProjects.forEach(function(project) {
      if (projects.includes(project)) {
        row.push(email); // Put email under selected projects
      } else {
        row.push(''); // Empty for non-selected projects
      }
    });

    // Append the row to the sheet
    sheet.appendRow(row);

    // Return success response with CORS headers
    return ContentService
      .createTextOutput(JSON.stringify({success: true, message: 'Data saved successfully'}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({success: false, error: error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Handle GET requests with form data
function doGet(e) {
  try {
    // Check if this is a form submission (has parameters)
    if (e.parameter && e.parameter.name) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

      var date = new Date().toLocaleDateString('en-SG', {
        year: 'numeric',
        month: 'short',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
      });

      var name = e.parameter.name || '';
      var email = e.parameter.email || '';
      var projectsStr = e.parameter.projects || '';
      var projects = projectsStr ? projectsStr.split(',') : [];

      // All project names (must match exactly with form values)
      var allProjects = [
        'Narra Residences',
        'Newport Residences',
        'Duet @ Emily',
        'Sophia Meadows',
        'Pinery Residences',
        'Tengah Garden Avenue',
        'River Modern',
        'Bayshore Road',
        'Media Circle (Parcel A)',
        'Lentor Gardens',
        'Dunearn Road',
        'Holland Link',
        'Lakeside Drive',
        'Chencharu Close',
        'Chuan Grove',
        'Dorset Road',
        'Former Thomson View condo',
        'Upper Thomson Road (Parcel A)',
        'Coastal Cabana (EC)',
        'Rivelle Tampines (EC)',
        'Woodlands Drive (EC)',
        'Sembawang Road (EC)',
        'Senja Close (EC)'
      ];

      // Build the row: Date, Name, Email, then email under each selected project
      var row = [date, name, email];

      allProjects.forEach(function(project) {
        if (projects.includes(project)) {
          row.push(email);
        } else {
          row.push('');
        }
      });

      sheet.appendRow(row);

      return ContentService
        .createTextOutput(JSON.stringify({success: true, message: 'Data saved successfully'}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Default response for status check
    return ContentService
      .createTextOutput(JSON.stringify({status: 'OK', message: 'Prospects Form API is running'}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({success: false, error: error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Run this function ONCE to set up the header row
function setupHeaders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var headers = [
    'Date',
    'Name',
    'Email',
    'Narra Residences',
    'Newport Residences',
    'Duet @ Emily',
    'Sophia Meadows',
    'Pinery Residences',
    'Tengah Garden Avenue',
    'River Modern',
    'Bayshore Road',
    'Media Circle (Parcel A)',
    'Lentor Gardens',
    'Dunearn Road',
    'Holland Link',
    'Lakeside Drive',
    'Chencharu Close',
    'Chuan Grove',
    'Dorset Road',
    'Former Thomson View condo',
    'Upper Thomson Road (Parcel A)',
    'Coastal Cabana (EC)',
    'Rivelle Tampines (EC)',
    'Woodlands Drive (EC)',
    'Sembawang Road (EC)',
    'Senja Close (EC)'
  ];

  // Set headers in first row
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format header row
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1e3a5f');
  headerRange.setFontColor('white');

  // Auto-resize columns
  for (var i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }

  // Freeze header row
  sheet.setFrozenRows(1);
}
