
function demo2() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Read the data for columns A to F
  var dataRange = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  var data = dataRange.getValues();

  // Extract headers and map column names to their indices

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(header => header.toString().toLowerCase().replace(/_/g, ''));

  console.log(headers)
  var clicksIndex = headers.indexOf("clicks");
  //console.log(clicksIndex)
  var impressionsIndex = headers.indexOf("impressions");
  //console.log(impressionsIndex)
  var currentCTRIndex = headers.indexOf("ctr"); // Adjust this if necessary
  var targetCTRIndex = headers.indexOf("targetctr"); // Adjust this if necessary
  console.log(targetCTRIndex)

  // Check for existing new column headers and store their indices
  var clicksDailyIndex = headers.indexOf("clicksdaily")+1;
  console.log(clicksDailyIndex)
  var clicksMonthlyIndex = headers.indexOf("clicksmonthly")+1;
  //console.log(clicksMonthlyIndex)
  var totalClicksIndex = headers.indexOf("totalclicks")+1;
  var percentageChangeClicksIndex = headers.indexOf("%changeclicks")+1;
  var impressionsDailyIndex = headers.indexOf("impressionsdaily")+1;
  var impressionsMonthlyIndex = headers.indexOf("impressionsmonthly")+1;
  var totalImpressionsIndex = headers.indexOf("totalimpressions")+1;
  var percentageChangeImpressionsIndex = headers.indexOf("%changeimpressions")+1;
  //console.log(percentageChangeImpressionsIndex)

  //console.log(headers.length)

  // Determine the next available column index to add new headers
  var nextColumnIndex = headers.length + 1; // Starting after existing headers

  // Add headers for the new columns dynamically only if they do not exist
  if (clicksDailyIndex === 0) {
    //console.log(nextColumnIndex)
    sheet.getRange(1, nextColumnIndex).setValue("clicks_daily");
    clicksDailyIndex = nextColumnIndex;
    nextColumnIndex++
  }
  if (clicksMonthlyIndex === 0) {
    sheet.getRange(1, nextColumnIndex).setValue("clicks_monthly");
    clicksMonthlyIndex = nextColumnIndex;
    nextColumnIndex++
  }
  if (totalClicksIndex === 0) {
    sheet.getRange(1, nextColumnIndex).setValue("total_clicks");
    totalClicksIndex = nextColumnIndex;
    nextColumnIndex++
  }
  if (percentageChangeClicksIndex === 0) {
    sheet.getRange(1, nextColumnIndex).setValue("%change_clicks");
    percentageChangeClicksIndex = nextColumnIndex;
    nextColumnIndex++
  }
  if (impressionsDailyIndex === 0) {
    sheet.getRange(1, nextColumnIndex).setValue("impressions_daily");
    impressionsDailyIndex = nextColumnIndex;
    nextColumnIndex++
  }
  if (impressionsMonthlyIndex === 0) {
    sheet.getRange(1, nextColumnIndex).setValue("impressions_monthly");
    impressionsMonthlyIndex = nextColumnIndex;
    nextColumnIndex++
  }
  if (totalImpressionsIndex === 0) {
    sheet.getRange(1, nextColumnIndex).setValue("total_impressions");
    totalImpressionsIndex = nextColumnIndex;
    nextColumnIndex++
  }
  if (percentageChangeImpressionsIndex === 0) {
    sheet.getRange(1, nextColumnIndex).setValue("%change_impressions");
    percentageChangeImpressionsIndex = nextColumnIndex;
    nextColumnIndex++
  }

  var outputData = []; // Array to store the output
  var colors = []; // Array to store the background colors

  // Step 1: Calculate average CTR for non-zero click rows
  var totalClicks = 0;
  var totalImpressions = 0;
  var totalRowsWithClicks = 0;

  for (var i = 1; i < data.length; i++) { // Start from 1 to skip headers
    var clicks = data[i][clicksIndex];
    var impressions = data[i][impressionsIndex];
    if (clicks > 0) {
      totalClicks += clicks;
      totalImpressions += impressions;
      totalRowsWithClicks++;
    }
  }


  var averageCTR = totalRowsWithClicks > 0 ? (totalClicks / totalImpressions) : 0;

  for (var i = 1; i < data.length; i++) { // Start from 1 to skip headers
    var clicks = data[i][clicksIndex];
    var impressions = data[i][impressionsIndex];
    var currentCTR = data[i][currentCTRIndex];
    var targetCTR = data[i][targetCTRIndex];

    var totalClicks = clicks;
    var totalImpressions = impressions;
    var percentageChangeClicks = 0;
    var percentageChangeImpressions = 0;

    var colorRow = ['#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff']; // Default white background for each row

    // If clicks are zero, leave all output cells blank and skip further calculations
    if (clicks === 0) {
      outputData.push(["", "", "", "", "", "", "", ""]);
      colors.push(colorRow); // Keep the row's color default
      continue; // Skip to the next row
    }

    if (targetCTR) {
      if (targetCTR > currentCTR) {
        totalClicks = Math.max(1, Math.round(targetCTR * impressions)); // Ensure at least 1 click
        percentageChangeClicks = ((totalClicks - clicks) / Math.max(clicks, 1)) * 100; // Avoid division by zero
        colorRow[2] = '#90EE90'; // Highlight the %change_clicks column in light green
      } else {
        totalImpressions = Math.round(clicks / targetCTR);
        percentageChangeImpressions = impressions === 0 ? 0 : ((totalImpressions - impressions) / impressions) * 100;
        colorRow[6] = '#90EE90'; // Highlight the %change_impressions column in light green
      }
    }

    // Calculate daily and monthly clicks and impressions
    var clicksMonthly = totalClicks - clicks;
    var impressionsMonthly = totalImpressions - impressions;
    var clicksDaily = clicksMonthly / 30;
    var impressionsDaily = impressionsMonthly / 30;

    // Collect data for the row
    outputData.push([
      clicksDaily.toFixed(2), 
      clicksMonthly, 
      totalClicks, 
      percentageChangeClicks.toFixed(2) + "%", 
      impressionsDaily.toFixed(2), 
      impressionsMonthly, 
      totalImpressions, 
      percentageChangeImpressions.toFixed(2) + "%"
    ]);
    colors.push(colorRow); // Collect colors for the row
  }

 // Write the entire output array to the sheet at once (optimize by bulk writing)
  sheet.getRange(2, clicksDailyIndex , outputData.length, outputData[0].length).setValues(outputData);

  // Set the number format for the relevant range
  sheet.getRange(2, clicksDailyIndex , outputData.length).setNumberFormat('0');
  sheet.getRange(2, clicksMonthlyIndex , outputData.length).setNumberFormat('0');
  sheet.getRange(2, totalClicksIndex , outputData.length).setNumberFormat('0');
  sheet.getRange(2, impressionsDailyIndex , outputData.length).setNumberFormat('0');
  sheet.getRange(2, impressionsMonthlyIndex , outputData.length).setNumberFormat('0');
  sheet.getRange(2, totalImpressionsIndex , outputData.length).setNumberFormat('0');


  // Apply the background colors to the relevant cells
  sheet.getRange(2, clicksDailyIndex+1 , colors.length, colors[0].length).setBackgrounds(colors);

  // Add notes for clarity
  sheet.getRange(1, targetCTRIndex + 1).setNote('This is the CTR that we want for a given query/URL. This is what is used to calculate the target impressions and clicks.');
  sheet.getRange(1, clicksDailyIndex ).setNote('This is how many daily clicks are needed to hit the target CTR.');
  sheet.getRange(1, clicksMonthlyIndex ).setNote('This is how many clicks for the month are needed to hit the target CTR.');
  sheet.getRange(1, totalClicksIndex ).setNote('The total amount of clicks that are required to hit the target CTR.');
  sheet.getRange(1, percentageChangeClicksIndex ).setNote('This is the percentage change in the number of clicks needed in order to hit the target CTR (Green is increase).');
  sheet.getRange(1, impressionsDailyIndex ).setNote('This is how many daily impressions are needed to hit the target CTR.');
  sheet.getRange(1, impressionsMonthlyIndex ).setNote('This is how many impressions for the month are needed to hit the target CTR.');
  sheet.getRange(1, totalImpressionsIndex ).setNote('The total amount of impressions that are required to hit the target CTR.');
  sheet.getRange(1, percentageChangeImpressionsIndex ).setNote('This is the percentage change in the number of impressions needed in order to hit the target CTR (Green is increase).');
}
