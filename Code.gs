function onEdit(e) {
  if (!e || !e.range) return;
  
  var sheet = e.range.getSheet();
  
  // Only trigger if B1 is edited on 'Monthly Expences' sheet
  if (sheet.getName() === 'Monthly Expences' && e.range.getA1Notation() === 'B1') {
    showCreditCardSummary();
  }
  if (sheet.getName() === 'Monthly Expences' && e.range.getA1Notation() === 'C1') {
    showPortfolioSummary();
  }

  if (sheet.getName() === 'Monthly Expences' && e.range.getA1Notation() === 'D1') {
    getFriendBalances();
  }
  if (sheet.getName() === 'Monthly Expences' && e.range.getA1Notation() === 'E1') {
    showInvestmentDashboard();
  }
  if (sheet.getName() === 'Monthly Expences' && e.range.getA1Notation() === 'F1') {
    showExp();
  }

}

function showExp() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('moneyFlowDashboard')
      .setWidth(1600)
      .setHeight(1000);
    SpreadsheetApp.getUi().showModalDialog(html, '💸 Investment Flow Dashboard');
  } catch (error) {
     //SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

function getMoneyFlowData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Money Flow and invest');

    if (!sheet) throw new Error('Sheet "Money Flow and invest" not found');

    var lastRow = sheet.getLastRow();
    // Data starts at row 2, columns A–L
    var values = sheet.getRange(2, 1, lastRow - 1, 12).getValues();

    function parseAmt(val) {
      if (!val && val !== 0) return 0;
      var s = val.toString().trim();
      if (s === '' || s === 'NTM' || s === 'N T M' || s.toUpperCase() === 'FAILED') return 0;
      // Strip any NTM suffix mixed in (e.g. "NTM Stocks going down")
      if (s.toUpperCase().startsWith('NTM')) return 0;
      var n = parseFloat(s.replace(/,/g, ''));
      return isNaN(n) ? 0 : n;
    }

    var data = [];
    values.forEach(function(row) {
      var monthRaw = row[0];
      var month;
      if (monthRaw instanceof Date && !isNaN(monthRaw)) {
        month = Utilities.formatDate(monthRaw, Session.getScriptTimeZone(), "MMM/yyyy");
      } else {
        month = (monthRaw || '').toString().trim();
      }
      if (!month) return; // skip blank/totals rows
      data.push({
        month:         month,
        sentHome:      parseAmt(row[1]),
        cal:           parseAmt(row[2]),
        ndb:           parseAmt(row[3]),
        binance:       parseAmt(row[4]),
        stock:         parseAmt(row[5]),
        fd:            parseAmt(row[6]),
        gold:          parseAmt(row[7]),
        totalInvested: parseAmt(row[8]),
        income:        parseAmt(row[9]),
        savingPct:     parseAmt(row[10]),
        notes:         (row[11] || '').toString().trim()
      });
    });

    return data;
  } catch (error) {
    throw new Error('Failed to load money flow data: ' + error.message);
  }
}

function showCreditCardSummary() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('showCreditCardSummary')
      .setWidth(1200)
      .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, '💳 Credit Card Dashboard');
  } catch (error) {
     //SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

function showPortfolioSummary() {
  try {
    // Show password modal first
    var passwordHtml = HtmlService.createHtmlOutputFromFile('passwordPrompt')
      .setWidth(600)
      .setHeight(650);
    SpreadsheetApp.getUi().showModalDialog(passwordHtml, '🔒 Enter Password');
  } catch (error) {
    // Silent fail on mobile
  }
}

function verifyPasswordAndShowPortfolio(enteredPassword) {
  const correctPassword = "9965"; // Change this to your desired password
  
  if (enteredPassword === correctPassword) {
    var html = HtmlService.createHtmlOutputFromFile('showPortfolioSummary')
      .setWidth(1400)
      .setHeight(1200);
    SpreadsheetApp.getUi().showModalDialog(html, '💳 Portfolio Dashboard');
    return true;
  } else {
    return false;
  }
}


// Function to get card data from sheet
// Function to get card data from the sheet
function getCardData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('CC_SW_CL_INST');
    
    if (!sheet) {
      throw new Error('Sheet "CC_SW_CL_INST" not found');
    }
    
    // Get data from range A4:H6 (data starts at row 4)
    var range = sheet.getRange('A4:H6');
    var values = range.getValues();
    
    return values;
  } catch (error) {
    throw new Error('Failed to load data: ' + error.message);
  }
}

// Function to save last statement balance to column H
function saveLastStatementBalance(rowNumber, value) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('CC_SW_CL_INST');
    
    if (!sheet) {
      throw new Error('Sheet "CC_SW_CL_INST" not found');
    }
    
    // Parse the value to ensure it's a number
    var numValue = parseFloat(value);
    if (isNaN(numValue)) {
      throw new Error('Invalid number format');
    }
    
    // Save to column H at the specified row
    var cell = sheet.getRange('H' + rowNumber);
    cell.setValue(numValue);
    
    return {
      success: true,
      message: 'Last statement balance saved successfully',
      row: rowNumber,
      value: numValue
    };
  } catch (error) {
    throw new Error('Failed to save data: ' + error.message);
  }
}

// Function to save monthly settled status to column D
function saveMonthlySettled(rowNumber, value) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('CC_SW_CL_INST');
    
    if (!sheet) {
      throw new Error('Sheet "CC_SW_CL_INST" not found');
    }
    
    // Convert to proper boolean
    var boolValue = (value === true || value === 'true' || value === 'TRUE');
    
    // Log for debugging
    Logger.log('Saving to row: ' + rowNumber + ', Column D, Value: ' + boolValue);
    
    // Save boolean value to column D (Monthly Settled checkbox)
    var cell = sheet.getRange('D' + rowNumber);
    cell.setValue(boolValue);
    
    // Force save
    SpreadsheetApp.flush();
    
    return {
      success: true,
      message: 'Monthly settled status saved successfully',
      row: rowNumber,
      value: boolValue,
      cellAddress: 'D' + rowNumber
    };
  } catch (error) {
    Logger.log('Error in saveMonthlySettled: ' + error.message);
    throw new Error('Failed to save status: ' + error.message);
  }
}

function getPortfolioData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Portfolio');
    
    if (!sheet) {
      throw new Error('Sheet "Portfolio" not found');
    }
    
    // Get data from range A3:E6 (includes header)
    var range = sheet.getRange('A2:B17');
    var values = range.getValues();
    
    return values;
  } catch (error) {
    throw new Error('Failed to load data: ' + error.message);
  }
}
function getGrowthData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Portfolio');
    
    if (!sheet) throw new Error('Sheet "Portfolio" not found');
    
    // Read Date (col R) + Amount (col S) starting at row 2
    var lastRow = sheet.getLastRow();
    var range = sheet.getRange('R2:S' + lastRow);
    var values = range.getValues();
    
    // Return only non-empty rows
    return values.filter(row => row[0] !== '' && row[1] !== '');
  } catch (error) {
    throw new Error('Failed to load growth data: ' + error.message);
  }
}


function updateCSEPrices() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UT_CRYPTO");
  const symbolsRange = sheet.getRange("A20:A29");
  const priceRange = sheet.getRange("B20:B29");

  const symbols = symbolsRange.getValues();
  const existingPrices = priceRange.getValues();
  const output = [];

  const url = "https://www.cse.lk/api/companyInfoSummery";

  symbols.forEach((row, i) => {
    const symbol = row[0];
    const currentValue = existingPrices[i][0];

    if (!symbol) {
      output.push([currentValue]);
      return;
    }

    const options = {
      method: "post",
      payload: { symbol },
      muteHttpExceptions: true
    };

    try {
      const res = UrlFetchApp.fetch(url, options);
      const json = JSON.parse(res.getContentText());
      const price = json?.reqSymbolInfo?.lastTradedPrice;

      // ✅ Only update if valid price exists
      if (price !== null && price !== undefined && price !== "" && price !== "N/A") {
        output.push([price]);
      } else {
        output.push([currentValue]);
      }
    } catch (e) {
      output.push([currentValue]);
    }
  });

  priceRange.setValues(output);
}

function sendMonthlyNotification() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getRange("F58:I83").getValues();

  const recipients = [
    "kamalanathanpirathees@gmail.com"
  ];

  const today = new Date();
  const currentMonthYear = Utilities.formatDate(today, Session.getScriptTimeZone(), "MMM/yyyy");

  // Calculate total pending balance
  let totalPending = 0;
  data.forEach(r => {
    if (r[3] === "Pending" && r[2]) totalPending += Number(r[2]);
  });

  for (let i = 0; i < data.length; i++) {
    const monthDate = data[i][1]; // G
    const amount = data[i][2];    // H
    const status = data[i][3];    // I

    const rowMonthYear = Utilities.formatDate(new Date(monthDate), Session.getScriptTimeZone(), "MMM/yyyy");

    if (rowMonthYear === currentMonthYear && status === "Pending") {
      const subject = `Payment Reminder — MacBook Sajeevan (${rowMonthYear})`;

      const htmlBody = `
      <div style="background:#f4f6f9;padding:20px;font-family:Segoe UI,Arial,sans-serif;">
        <div style="max-width:600px;margin:auto;background:#ffffff;border-radius:12px;overflow:hidden;box-shadow:0 4px 12px rgba(0,0,0,0.1);">
          
          <!-- Header -->
          <div style="background:linear-gradient(135deg,#0f172a,#1e3a8a);color:#ffffff;padding:16px 20px;">
            <h2 style="margin:0;"> MacBook Installment Reminder</h2>
            <p style="margin:4px 0 0;font-size:13px;opacity:0.85;">${rowMonthYear}</p>
          </div>

          <!-- Body -->
          <div style="padding:0 24px 20px;color:#1f2937;">
            <p style="font-size:15px;">Hi</p>
            <p>This is a friendly reminder for your MacBook installment payment:</p>

            <table style="width:100%;border-collapse:collapse;margin-top:12px;">
              <tr>
                <td style="padding:10px;background:#f1f5f9;font-weight:600;">Month</td>
                <td style="padding:10px;">${rowMonthYear}</td>
              </tr>
              <tr>
                <td style="padding:10px;background:#f1f5f9;font-weight:600;">Amount</td>
                <td style="padding:10px;color:#16a34a;font-weight:700;">${amount} LKR</td>
              </tr>
              <tr>
                <td style="padding:10px;background:#f1f5f9;font-weight:600;">Remaining Balance</td>
                <td style="padding:10px;color:#dc2626;font-weight:700;">${totalPending} LKR</td>
              </tr>
            </table>

            <p style="margin-top:18px;">
              Please arrange payment at your convenience
            </p>

            <p style="margin-top:20px;font-size:13px;color:#6b7280;">
              Thanks,<br/>
              Pirathees
            </p>
          </div>

        </div>
      </div>
      `;

      GmailApp.sendEmail(recipients.join(","), subject, "", {
        htmlBody: htmlBody
      });

      // ✅ Update Status column (I)
      sheet.getRange(58 + i, 9).setValue("Notified");
      break;
    }
  }
}

// Splitwise API check
const CONSUMER_KEY = 'Gdg3ICNvkIyre7PAYsj6FRrgxeaOdnHob0mSYtHp';
const CONSUMER_SECRET = 'HGoLJOypYGNA1GA4pChql6uquVqzbVDuZ0LONrNZ';

function getSplitwiseData() {
  
  const apiKey = 'Lo46GCiwIVipgzU3aV64dM5YysVZkLP3nY6vxtWv';
  const url = 'https://secure.splitwise.com/api/v3.0/get_current_user';
  
  const options = {
    'method': 'get',
    'headers': {
      'Authorization': 'Bearer ' + apiKey
    }
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  Logger.log(data);
  
  return data;
}

function getSplitwiseData() {
  
  const apiKey = 'Lo46GCiwIVipgzU3aV64dM5YysVZkLP3nY6vxtWv';
  const url = 'https://secure.splitwise.com/api/v3.0/get_current_user';
  
  const options = {
    'method': 'get',
    'headers': {
      'Authorization': 'Bearer ' + apiKey
    }
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  Logger.log(data);
  
  return data;
}

function getFriendBalances() {
  const apiKey = 'Lo46GCiwIVipgzU3aV64dM5YysVZkLP3nY6vxtWv';
  const url = 'https://secure.splitwise.com/api/v3.0/get_friends';
  
  const options = {
    'method': 'get',
    'headers': {
      'Authorization': 'Bearer ' + apiKey
    }
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    const friends = data.friends;
    
    // Prepare data and track totals by currency
    const balanceData = [];
    const totals = {}; // Track totals by currency
    
    friends.forEach(friend => {
      const fullName = `${friend.first_name} ${friend.last_name || ''}`.trim();
      
      friend.balance.forEach(bal => {
        const amount = parseFloat(bal.amount);
        const currency = bal.currency_code;
        

        const adjustedAmount = amount;
        
        let status = '';
        if (adjustedAmount < 0) {
          status = 'I need to pay';
        } else if (adjustedAmount > 0) {
          status = 'They need to pay';
        } else {
          status = 'Settled up';
        }
        
        balanceData.push({
          name: fullName,
          amount: adjustedAmount,
          currency: currency,
          status: status
        });
        
        // Add to totals
        if (!totals[currency]) {
          totals[currency] = 0;
        }
        totals[currency] += adjustedAmount;
      });
    });
    
    // Get the sheet balance for validation
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CC_SW_CL_INST');
    const sheetBalanceCell = 'C44'; // Change this to your actual cell reference
    const sheetBalance = sheet.getRange(sheetBalanceCell).getValue();
    
    // Validate LKR balance
    const lkrTotal = totals['LKR'] || 0;
    const validationResult = validateBalance(lkrTotal, sheetBalance);
    
    // Show modal with data and validation
    showBalanceModal(balanceData, totals, validationResult, sheetBalance);
    
  } catch (error) {
    //SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

function validateBalance(splitwiseAmount, sheetAmount) {
  // Convert to numbers and remove decimals for comparison
  const splitwiseRounded = Math.round(splitwiseAmount);
  const sheetRounded = Math.round(parseFloat(sheetAmount) || 0);
  
  // Calculate difference
  const difference = Math.abs(splitwiseRounded - sheetRounded);
  
  // Allow ±10 LKR tolerance
  const tolerance = 10;
  const match = difference <= tolerance;
  
  return {
    splitwiseAmount: splitwiseAmount,
    sheetAmount: sheetAmount,
    splitwiseRounded: splitwiseRounded,
    sheetRounded: sheetRounded,
    difference: difference,
    tolerance: tolerance,
    match: match
  };
}

function showBalanceModal(balanceData, totals, validationResult, sheetBalance) {

  try {
      var html = HtmlService.createTemplateFromFile('SplitwiseBalances');
      html.balanceData = balanceData;
      html.totals = totals;
      html.validationResult = validationResult;
      html.sheetBalance = sheetBalance;
      
      var htmlOutput = html.evaluate()
        .setWidth(1400)
        .setHeight(800);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '💰 Splitwise Balances');
    SpreadsheetApp.getUi().showModalDialog(html, '💳 Credit Card Dashboard');
  } catch (error) {
    // Silent fail on mobile
  }
}

function updateGoldPrice() {
  const url = "https://ravijewellers.lk/";
  const html = UrlFetchApp.fetch(url).getContentText();

  const match = html.match(/goldrate-rate[^>]*>.*?LKR\s*([\d,]+)/i);
  if (!match) throw new Error("Gold rate not found");

  const rate = Number(match[1].replace(/,/g, ""));
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UT_CRYPTO");

  sheet.getRange("B1").setValue(new Date());
  sheet.getRange("C1").setValue(rate);
}

function updateUnitTrustPrices() {
  const url = "https://www.utasl.lk/unit-prices/";
  const sheet = SpreadsheetApp.getActive().getSheetByName("UT_CRYPTO");

  const fundNamesRange = sheet.getRange("A8:A12");
  const priceRange = sheet.getRange("D8:D12");

  const fundNames = fundNamesRange.getValues().flat();
  const existingPrices = priceRange.getValues();

  const html = UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText();

  function htmlDecode(str) {
    return str.replace(/&amp;/g, "&")
              .replace(/&nbsp;/g, " ")
              .replace(/&quot;/g, '"')
              .replace(/&#39;/g, "'")
              .replace(/&lt;/g, "<")
              .replace(/&gt;/g, ">");
  }

  const rowRegex = /<tr>\s*<td>(.*?)<\/td>\s*<td>(.*?)<\/td>\s*<td>(.*?)<\/td>\s*<td>(.*?)<\/td>\s*<\/tr>/g;

  let match;
  const fundMap = {};

  while ((match = rowRegex.exec(html)) !== null) {
    const fundName = htmlDecode(match[2].trim());
    const sellingPrice = match[3].trim();
    fundMap[fundName] = sellingPrice;
  }

  const output = fundNames.map((name, i) => {
    const newPrice = fundMap[name];
    const currentValue = existingPrices[i][0];

    if (newPrice && newPrice !== "NOT FOUND") {
      return [newPrice];
    }
    return [currentValue];
  });

  priceRange.setValues(output);
}

function logMonthlyEarnings() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("UT_CRYPTO");

  const month = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM");

  const utEarn = sheet.getRange("F2").getValue();
  const stockEarn = sheet.getRange("I2").getValue();
  const goldEarn = sheet.getRange("C3").getValue();

  const utInvest = sheet.getRange("F1").getValue();
  const stockInvest = sheet.getRange("I1").getValue();
  const goldInvest = sheet.getRange("C2").getValue();

  const startRow = 46;
  const lastRow = sheet.getLastRow();

  const existing = lastRow >= startRow
    ? sheet.getRange(startRow, 1, lastRow - startRow + 1, 7).getValues().filter(r => r.some(c => c !== ""))
    : [];

  existing.push([
    month,
    utEarn,
    stockEarn,
    goldEarn,
    utInvest,
    stockInvest,
    goldInvest
  ]);

  sheet.getRange(startRow, 1, existing.length, 7).setValues(existing);
}


function showInvestmentDashboard() {
  const html = HtmlService.createHtmlOutputFromFile("investmentDashboard")
    .setWidth(1600)
    .setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(html, "📊 Investment Performance Dashboard");
}

function getInvestmentData() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("UT_CRYPTO");
  const startRow = 46;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return [];

  const raw = sheet.getRange(startRow, 1, lastRow - startRow + 1, 7).getValues();

  return raw
    .filter(r => r[0])
    .map(r => [
      Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), "yyyy-MM"),
      Number(r[1]) || 0,
      Number(r[2]) || 0,
      Number(r[3]) || 0,
      Number(r[4]) || 0,
      Number(r[5]) || 0,
      Number(r[6]) || 0
    ]);
}







