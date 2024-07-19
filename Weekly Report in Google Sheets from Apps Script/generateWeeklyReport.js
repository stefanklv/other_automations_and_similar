function generateWeeklyReport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('weekly_sales_data'); // Ensure this matches your sheet name
  if (!sheet) {
    Logger.log('Sheet not found!');
    return;
  }

  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  var today = new Date();
  var lastWeek = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 7);

  var totalSales = 0;
  var totalQuantitySold = 0;
  var productSales = {};
  var salespersonSales = {};

  for (var i = 1; i < data.length; i++) {
    var rowDate = new Date(data[i][0]);
    if (rowDate >= lastWeek && rowDate <= today) {
      var product = data[i][2];
      var quantitySold = data[i][3];
      var sales = data[i][4];
      var salesperson = data[i][1];

      totalSales += sales;
      totalQuantitySold += quantitySold;

      if (productSales[product]) {
        productSales[product] += quantitySold;
      } else {
        productSales[product] = quantitySold;
      }

      if (salespersonSales[salesperson]) {
        salespersonSales[salesperson] += sales;
      } else {
        salespersonSales[salesperson] = sales;
      }
    }
  }

  if (totalQuantitySold === 0) {
    Logger.log('No data found for the past week.');
    return;
  }

  var mostSoldProduct = '';
  var mostSoldQuantity = 0;

  for (var product in productSales) {
    if (productSales[product] > mostSoldQuantity) {
      mostSoldProduct = product;
      mostSoldQuantity = productSales[product];
    }
  }

  var reportText = formatReport(totalSales, totalQuantitySold, mostSoldProduct, mostSoldQuantity, salespersonSales);
  sendEmail(reportText);
}

function formatReport(totalSales, totalQuantitySold, mostSoldProduct, mostSoldQuantity, salespersonSales) {
  var report = 'Weekly Sales Summary Report\n\n';
  report += 'Total Sales: $' + totalSales.toFixed(2) + '\n';
  report += 'Total Quantity Sold: ' + totalQuantitySold + '\n';
  report += 'Most Sold Product: ' + mostSoldProduct + ' (Quantity: ' + mostSoldQuantity + ')\n\n';
  report += 'Sales Breakdown by Salesperson:\n';

  for (var salesperson in salespersonSales) {
    report += salesperson + ': $' + salespersonSales[salesperson].toFixed(2) + '\n';
  }

  return report;
}

function sendEmail(reportText) {
  var recipients = 'Enter your email address here'; // Replace with actual email addresses
  var subject = 'Weekly Sales Summary Report';
  var body = reportText;

  MailApp.sendEmail(recipients, subject, body);
}

// Set a time-driven trigger to run the generateWeeklyReport function every Monday at 8 AM
function createTrigger() {
  ScriptApp.newTrigger('generateWeeklyReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();
}
