# Auto-Month-lable-in-Google-Sheets
Script created using ChatGBT wich automaticaly creates month lable in Column A using date data in Column B. 

🔹 Step 1: Open Apps Script
	1.	Open your Google Sheet.
	2.	Go to Extensions > Apps Script.
	3.	A new tab will open with the Apps Script editor.

⸻

🔹 Step 2: Add the Script
	1.	Delete any code already there.
	2.	Paste this script:
 
 function mergeAndColorMonths() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var monthCol = 1; // Column A for months
  var dateCol = 2;  // Column B for dates
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  var dates = sheet.getRange(1, dateCol, lastRow).getValues();
  var monthRange = sheet.getRange(1, monthCol, lastRow);
  var monthValues = monthRange.getValues(); // keep existing values

  // Czech month names
  var monthNames = [
    "Leden", "Únor", "Březen", "Duben", "Květen", "Červen",
    "Červenec", "Srpen", "Září", "Říjen", "Listopad", "Prosinec"
  ];

  // Month → Pastel Color map
  var colors = {
    "Leden": "#d0e0e3",
    "Únor": "#c9daf8",
    "Březen": "#cfe2f3",
    "Duben": "#d9d2e9",
    "Květen": "#ead1dc",
    "Červen": "#e6b8af",
    "Červenec": "#f9cb9c",
    "Srpen": "#ffe599",
    "Září": "#f4cccc",
    "Říjen": "#fce5cd",
    "Listopad": "#fff2cc",
    "Prosinec": "#d9ead3"
  };

  // Fill month names only for valid dates, leave other rows untouched
  for (var i = 0; i < lastRow; i++) {
    var d = dates[i][0];
    if (Object.prototype.toString.call(d) === "[object Date]" && !isNaN(d)) {
      monthValues[i][0] = monthNames[d.getMonth()];
    }
  }
  monthRange.setValues(monthValues);

  // Clear old merges safely
  var mergedRanges = monthRange.getMergedRanges();
  mergedRanges.forEach(function(range) {
    range.breakApart();
  });

  // Merge and color month cells (skip blanks)
  var start = 1;
  for (var i = 2; i <= lastRow; i++) {
    if (monthValues[i - 1][0] !== monthValues[i - 2][0]) {
      if (i - start > 1 && monthValues[i - 2][0] !== "") {
        var range = sheet.getRange(start, monthCol, i - start, 1);
        range.mergeVertically()
             .setVerticalAlignment("MIDDLE")
             .setHorizontalAlignment("CENTER");
        var month = monthValues[i - 2][0];
        if (colors[month]) range.setBackground(colors[month]);
      }
      start = i;
    }
  }

  // Handle last group (skip blank)
  if (lastRow - start + 1 > 1 && monthValues[lastRow - 1][0] !== "") {
    var range = sheet.getRange(start, monthCol, lastRow - start + 1, 1);
    range.mergeVertically()
         .setVerticalAlignment("MIDDLE")
         .setHorizontalAlignment("CENTER");
    var month = monthValues[lastRow - 1][0];
    if (colors[month]) range.setBackground(colors[month]);
  }
}

// Only trigger when column B is edited
function onEdit(e) {
  if (!e) return;
  var editedCol = e.range.getColumn();
  if (editedCol === 2) { // Column B
    mergeAndColorMonths();
  }
}

 🔹 Step 3: Save & Authorize
	1.	Click the 💾 Save button (give your project a name, e.g. “MonthMerge”).
	2.	Run mergeAndColorMonths once manually:
	•	Click the dropdown next to the ▶️ button, select mergeAndColorMonths, then press ▶️.
	•	The first time you run it, Google will ask for permissions → Allow.

⸻

🔹 Step 4: Use it
	•	Now, whenever you type or change a date in column B,
→ column A will automatically show the month name,
→ merge identical months,
→ and color the merged cell.
