/* ── Web App entry point (SPA) ── */
function doGet(e) {
  var appUrl = ScriptApp.getService().getUrl();
  var html = HtmlService.createHtmlOutputFromFile('render');
  var inject = '<script>var _WEB_APP=true; var _APP_URL="' + appUrl + '";<\/script>\n';
  var content = html.getContent().replace('</head>', inject + '</head>');
  return HtmlService.createHtmlOutput(content)
    .setTitle('PTS Finance')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onEdit(e) {
  if (!e || !e.range) return;
  
  var sheet = e.range.getSheet();
  
  // A1 → Home Command Centre
  if (sheet.getName() === 'Monthly Expences' && e.range.getA1Notation() === 'A1') {
    showHomeDashboard();
  }
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
    showInvestFlow();
  }
  if (sheet.getName() === 'Monthly Expences' && e.range.getA1Notation() === 'G1') {
    showExpenseDashboard();
  }
  if (sheet.getName() === 'Monthly Expences' && e.range.getA1Notation() === 'H1') {
    showCSEDashboard();
  }
}

// ═══════════════════════════════════════════════════════════════
// HOME COMMAND CENTRE
// ═══════════════════════════════════════════════════════════════

function showHomeDashboard() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('homeDashboard')
      .setWidth(2000)
      .setHeight(2000);
    SpreadsheetApp.getUi().showModalDialog(html, '🏠 PTS Command Centre');
  } catch (error) { /* silent fail */ }
}

function getHomeData() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var out = {};

  // ── 1. EXPENSES ──────────────────────────────────────────────
  try {
    var expSheet  = ss.getSheetByName('Monthly Expences');
    var headerRow = expSheet.getRange(6, 1, 1, 30).getValues()[0];
    var tableData = expSheet.getRange(7, 1, 11, 30).getValues();

    var monthCols = [];
    for (var c = 1; c < headerRow.length; c++) {
      var h = (headerRow[c] || '').toString().trim();
      if (!h || h.toUpperCase() === 'TOTAL') continue;
      monthCols.push({ label: h, colIndex: c });
    }

    var totalRow = null, categories = [];
    tableData.forEach(function(row) {
      var name = (row[0] || '').toString().trim();
      if (!name) return;
      if (name.toLowerCase().indexOf('total') >= 0) { totalRow = row; }
      else { categories.push({ name: name, values: row }); }
    });

    if (monthCols.length > 0 && totalRow) {
      var last = monthCols[monthCols.length - 1];
      var prev = monthCols.length > 1 ? monthCols[monthCols.length - 2] : null;
      var thisTotal = parseFloat(totalRow[last.colIndex]) || 0;
      var prevTotal = prev ? (parseFloat(totalRow[prev.colIndex]) || 0) : 0;

      var cats = categories.map(function(c) {
        return { name: c.name, val: parseFloat(c.values[last.colIndex]) || 0 };
      }).sort(function(a, b) { return b.val - a.val; });

      out.expenses = {
        latestMonth: last.label,
        prevMonth:   prev ? prev.label : '',
        thisTotal:   thisTotal,
        prevTotal:   prevTotal,
        topCats:     cats.slice(0, 3),
        allMonths:   monthCols.map(function(mc) {
          return { label: mc.label, total: parseFloat(totalRow[mc.colIndex]) || 0 };
        }).slice(-8)  // last 8 months for trend chart
      };
    }
  } catch(e) { out.expenses = null; }

  // ── 2. CREDIT CARDS + SW SHEET BALANCE ───────────────────────
  try {
    var ccSheet  = ss.getSheetByName('CC_SW_CL_INST');
    var ccValues = ccSheet.getRange('A4:H6').getValues();
    var swBal    = parseFloat(ccSheet.getRange('C44').getValue()) || 0;

    var cards = ccValues.map(function(row) {
      return {
        name:         (row[0] || '').toString().trim(),
        balance:      parseFloat(row[2]) || 0,
        settled:      (row[3] === true || row[3] === 'TRUE'),
        lastStatement: parseFloat(row[7]) || 0
      };
    }).filter(function(c) { return c.name; });

    out.creditCards = { cards: cards, swBalance: swBal };
  } catch(e) { out.creditCards = null; }

  // ── 3. INVESTMENTS (latest row from UT_CRYPTO log) ───────────
  try {
    var utSheet  = ss.getSheetByName('UT_CRYPTO');
    var startRow = 46;
    var lastRow  = utSheet.getLastRow();
    if (lastRow >= startRow) {
      var rows = utSheet.getRange(startRow, 1, lastRow - startRow + 1, 7).getValues()
        .filter(function(r) { return r[0]; });
      if (rows.length > 0) {
        var lr = rows[rows.length - 1];
        var monthLabel = '';
        try { monthLabel = Utilities.formatDate(new Date(lr[0]), Session.getScriptTimeZone(), 'MMM yyyy'); }
        catch(e2) { monthLabel = lr[0].toString(); }
        out.investments = {
          month:         monthLabel,
          utEarnings:    Number(lr[1]) || 0,
          stockEarnings: Number(lr[2]) || 0,
          goldEarnings:  Number(lr[3]) || 0,
          utInvested:    Number(lr[4]) || 0,
          stockInvested: Number(lr[5]) || 0,
          goldInvested:  Number(lr[6]) || 0
        };
      }
    }
  } catch(e) { out.investments = null; }

  // ── 4. MONEY FLOW ─────────────────────────────────────────────
  try {
    var mfSheet  = ss.getSheetByName('Money Flow and invest');
    var mfLastRow = mfSheet.getLastRow();
    var mfRows = mfSheet.getRange(2, 1, mfLastRow - 1, 11).getValues()
      .filter(function(r) { return r[0]; });
    if (mfRows.length > 0) {
      var ml = mfRows[mfRows.length - 1];
      var mLabel = ml[0] instanceof Date
        ? Utilities.formatDate(ml[0], Session.getScriptTimeZone(), 'MMM/yyyy')
        : (ml[0] || '').toString().trim();
      out.moneyFlow = {
        month:         mLabel,
        sentHome:      Number(ml[1]) || 0,
        totalInvested: Number(ml[8]) || 0,
        income:        Number(ml[9]) || 0,
        savingPct:     Number(ml[10]) || 0,
        history: mfRows.slice(-7).map(function(r) {
          var lbl = r[0] instanceof Date
            ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'MMM/yy')
            : (r[0] || '').toString().slice(0, 8);
          return { month: lbl, savingPct: Number(r[10]) || 0, income: Number(r[9]) || 0 };
        })
      };
    }
  } catch(e) { out.moneyFlow = null; }

  // ── 5. PORTFOLIO TOTAL + GROWTH ──────────────────────────────
  // B16 = Total Portfolio · LKR  |  P3:R = growth log (date, amount, note)
  try {
    var portSheet = ss.getSheetByName('Portfolio');
    var portTotal = parseFloat(portSheet.getRange('B16').getValue()) || 0;

    // Growth log: col P(16)=date, Q(17)=amount, R(18)=note, starting row 3
    var pLastRow = portSheet.getLastRow();
    var growth = [];
    if (pLastRow >= 3) {
      var gVals = portSheet.getRange(3, 16, pLastRow - 2, 3).getValues();
      gVals.forEach(function(row) {
        if (!row[0] && !row[1]) return; // skip blank rows
        var lbl = '';
        try {
          lbl = (row[0] instanceof Date)
            ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'M/yyyy')
            : (row[0] || '').toString().trim();
        } catch(e2) { lbl = (row[0] || '').toString().trim(); }
        if (!lbl) return;
        growth.push({
          date:   lbl,
          amount: Number(row[1]) || 0,
          note:   (row[2] || '').toString().trim()
        });
      });
    }

    out.portfolio = { total: portTotal, growth: growth };
  } catch(e) { out.portfolio = null; }

  return out;
}

function showInvestFlow() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('moneyFlowDashboard')
      .setWidth(2000)
      .setHeight(2000);
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
      var raw = row[0];
      var month;
      if (raw instanceof Date && !isNaN(raw.getTime())) {
        month = Utilities.formatDate(raw, Session.getScriptTimeZone(), "MMM/yyyy");
      } else {
        month = (raw || '').toString().trim();
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

function saveMoneyFlowItemStatus(monthStr, jsonString) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Money Flow and invest');
    if (!sheet) throw new Error('Sheet "Money Flow and invest" not found');

    var lastRow = sheet.getLastRow();
    var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

    for (var i = 0; i < values.length; i++) {
      var raw = values[i][0];
      var month;
      if (raw instanceof Date && !isNaN(raw.getTime())) {
        month = Utilities.formatDate(raw, Session.getScriptTimeZone(), "MMM/yyyy");
      } else {
        month = (raw || '').toString().trim();
      }
      if (month === monthStr) {
        sheet.getRange(i + 2, 12).setValue(jsonString);
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    throw new Error('Month "' + monthStr + '" not found in sheet');
  } catch (error) {
    throw new Error('Failed to save status: ' + error.message);
  }
}

function showCreditCardSummary() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('showCreditCardSummary')
      .setWidth(2000)
      .setHeight(2000);
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
      .setWidth(2000)
      .setHeight(2000);
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

    if (!sheet) throw new Error('Sheet "Portfolio" not found');

    // ── Asset data A2:B17 (always loaded) ──────────────────────────
    var assets = sheet.getRange('A2:B17').getValues();

    // ── Growth log: P(col16)=date, Q(col17)=amount, R(col18)=note ──
    // Wrapped in its own try-catch — if this fails, assets still load
    var growth = [];
    try {
      var pLastRow = sheet.getLastRow();
      var pMaxCol  = sheet.getLastColumn();
      if (pLastRow >= 3 && pMaxCol >= 16) {
        var numCols = Math.min(3, pMaxCol - 15); // how many of P/Q/R actually exist
        var gVals = sheet.getRange(3, 16, pLastRow - 2, numCols).getValues();
        gVals.forEach(function(row) {
          if (!row[0] && !row[1]) return;
          var lbl = '';
          try {
            lbl = (row[0] instanceof Date)
              ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'M/yyyy')
              : (row[0] || '').toString().trim();
          } catch(e3) { lbl = (row[0] || '').toString().trim(); }
          if (!lbl) return;
          growth.push({
            date:   lbl,
            amount: Number(row[1]) || 0,
            note:   numCols >= 3 ? (row[2] || '').toString().trim() : ''
          });
        });
      }
    } catch(eg) {
      // Growth data unavailable — assets still returned below
      Logger.log('Growth data read failed: ' + eg.message);
    }

    return { assets: assets, growth: growth };
  } catch (error) {
    throw new Error('Failed to load data: ' + error.message);
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
  const headers = { 'Authorization': 'Bearer ' + apiKey };
  const opts = { method: 'get', headers: headers };

  try {
    // ── 1. Friend balances ──
    const friendsResp = UrlFetchApp.fetch('https://secure.splitwise.com/api/v3.0/get_friends', opts);
    const friends = JSON.parse(friendsResp.getContentText()).friends;

    const balanceData = [];
    const totals = {};

    friends.forEach(friend => {
      const fullName = (friend.first_name + ' ' + (friend.last_name || '')).trim();
      friend.balance.forEach(bal => {
        const amount = parseFloat(bal.amount);
        const currency = bal.currency_code;
        balanceData.push({
          name: fullName,
          friendId: friend.id,
          amount: amount,
          currency: currency,
          status: amount < 0 ? 'I need to pay' : amount > 0 ? 'They need to pay' : 'Settled up'
        });
        totals[currency] = (totals[currency] || 0) + amount;
      });
    });

    // ── 2. Current user ID (needed to compute per-expense net balance) ──
    var currentUserId = null;
    try {
      const meResp = UrlFetchApp.fetch('https://secure.splitwise.com/api/v3.0/get_current_user', opts);
      currentUserId = JSON.parse(meResp.getContentText()).user.id;
    } catch(e) {}

    // ── 3. Recent 10 expenses ──
    var recentTransactions = [];
    try {
      const expResp = UrlFetchApp.fetch(
        'https://secure.splitwise.com/api/v3.0/get_expenses?limit=15&offset=0', opts);
      const expenses = JSON.parse(expResp.getContentText()).expenses;

      recentTransactions = expenses
        .filter(function(e) { return !e.deleted_at; })
        .slice(0, 10)
        .map(function(e) {
          var myNet = 0;
          var paidByMe = false;
          if (currentUserId) {
            var myUser = null;
            for (var i = 0; i < e.users.length; i++) {
              if (e.users[i].user_id === currentUserId) { myUser = e.users[i]; break; }
            }
            if (myUser) {
              myNet = parseFloat(myUser.net_balance || 0);
              paidByMe = parseFloat(myUser.paid_share || 0) > 0;
            }
          }
          return {
            id: e.id,
            description: e.description || '(no description)',
            date: Utilities.formatDate(new Date(e.date), Session.getScriptTimeZone(), 'dd MMM yyyy'),
            cost: parseFloat(e.cost),
            currency: e.currency_code,
            myNet: myNet,        // +ve = owed to me, -ve = I owe
            paidByMe: paidByMe,
            isPayment: e.payment === true
          };
        });
    } catch(e) {}

    // ── 4. Validate LKR ──
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CC_SW_CL_INST');
    const sheetBalance = sheet.getRange('C44').getValue();
    const lkrTotal = totals['LKR'] || 0;
    const validationResult = validateBalance(lkrTotal, sheetBalance);

    showBalanceModal(balanceData, totals, validationResult, sheetBalance, recentTransactions);

  } catch (error) {
    // silent fail
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

function showBalanceModal(balanceData, totals, validationResult, sheetBalance, recentTransactions) {
  try {
    var html = HtmlService.createTemplateFromFile('SplitwiseBalances');
    html.balanceData = balanceData;
    html.totals = totals;
    html.validationResult = validationResult;
    html.sheetBalance = sheetBalance;
    html.recentTransactions = recentTransactions || [];

    var htmlOutput = html.evaluate()
      .setWidth(2000)
      .setHeight(2000);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '💰 Splitwise Balances');
  } catch (error) {
    // Silent fail
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

  const utEarn    = Number(sheet.getRange("F2").getValue());
  const stockEarn = Number(sheet.getRange("I2").getValue());
  const goldEarn  = Number(sheet.getRange("C3").getValue());
  const utInvest  = Number(sheet.getRange("F1").getValue());
  const stockInvest = Number(sheet.getRange("I1").getValue());
  const goldInvest  = Number(sheet.getRange("C2").getValue());

  const startRow = 46;
  const lastRow  = sheet.getLastRow();

  const existing = lastRow >= startRow
    ? sheet.getRange(startRow, 1, lastRow - startRow + 1, 7).getValues().filter(r => r.some(c => c !== ""))
    : [];

  // ── Duplicate check: compare current values against the last logged row ──
  if (existing.length > 0) {
    const last = existing[existing.length - 1];
    const sameEarnings  = Number(last[1]) === utEarn    &&
                          Number(last[2]) === stockEarn  &&
                          Number(last[3]) === goldEarn;
    const sameInvested  = Number(last[4]) === utInvest  &&
                          Number(last[5]) === stockInvest &&
                          Number(last[6]) === goldInvest;

    if (sameEarnings && sameInvested) {
      return {
        logged:  false,
        reason:  'duplicate',
        message: 'No change detected — values are identical to the last log entry (' + last[0] + '). Nothing was saved.'
      };
    }
  }

  existing.push([month, utEarn, stockEarn, goldEarn, utInvest, stockInvest, goldInvest]);
  sheet.getRange(startRow, 1, existing.length, 7).setValues(existing);

  return {
    logged:  true,
    message: 'Logged successfully for ' + month
  };
}


// ── Log current portfolio total (B16) into growth log (P:Q:R from row 3) ──
function logPortfolioAmount(note) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Portfolio');
  if (!sheet) throw new Error('Portfolio sheet not found');

  var currentAmount = Number(sheet.getRange('B16').getValue());
  if (!currentAmount) throw new Error('B16 is empty or zero — nothing to log');

  var dateLabel = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/yyyy');
  var noteStr   = (note || '').toString().trim();

  // Read existing growth log: P(col16), Q(col17), R(col18) from row 3
  var lastRow  = sheet.getLastRow();
  var lastCol  = sheet.getLastColumn();
  var existing = [];

  if (lastRow >= 3 && lastCol >= 16) {
    var numCols = Math.min(3, lastCol - 15);
    var raw = sheet.getRange(3, 16, lastRow - 2, numCols).getValues();
    existing = raw.filter(function(r) { return r[0] !== '' || r[1] !== ''; });
  }

  // ── Duplicate check: compare amount against last logged entry ──
  if (existing.length > 0) {
    var last = existing[existing.length - 1];
    if (Number(last[1]) === currentAmount) {
      return {
        logged:  false,
        reason:  'duplicate',
        message: 'Portfolio value is unchanged (₨' +
                 currentAmount.toLocaleString('en-IN') +
                 ') — same as last log (' + last[0] + '). Nothing saved.'
      };
    }
  }

  // ── Append new row ──
  var newRow = [dateLabel, currentAmount, noteStr];
  var writeRow = existing.length + 1; // 1-based offset from row 3
  sheet.getRange(2 + writeRow, 16, 1, 3).setValues([newRow]);

  return {
    logged:  true,
    amount:  currentAmount,
    date:    dateLabel,
    message: 'Portfolio logged: ₨' + currentAmount.toLocaleString('en-IN') + ' for ' + dateLabel
  };
}

// ═══════════════════════════════════════════════════════════════
// MONTHLY EXPENSE DASHBOARD
// ═══════════════════════════════════════════════════════════════

function showExpenseDashboard() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('expenseDashboard')
      .setWidth(2000)
      .setHeight(2000);
    SpreadsheetApp.getUi().showModalDialog(html, '📊 Monthly Expense Dashboard');
  } catch (error) {
    // silent fail
  }
}

function getExpenseData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Monthly Expences');
    if (!sheet) throw new Error('Sheet "Monthly Expences" not found');

    // ── Row 6: month headers (cols A onward, fixed 30-col window) ─
    // Col A = "Category", then month names, last meaningful col = "TOTAL"
    var headerRow = sheet.getRange(6, 1, 1, 30).getValues()[0];
    var monthCols = []; // [{label, colIndex, year}]  colIndex = 0-based in this array
    for (var c = 1; c < headerRow.length; c++) {
      var h = (headerRow[c] || '').toString().trim();
      if (!h) continue;                          // blank — past the data
      if (h.toUpperCase() === 'TOTAL') continue; // skip grand-total column
      // Extract year: header contains "20XX" (e.g. "JUNE 2025", "JAN 2026")
      var yr = 'Unknown';
      var ym = h.match(/20\d{2}/);
      if (ym) yr = ym[0];
      monthCols.push({ label: h, colIndex: c, year: yr });
    }

    // ── Rows 7–17: 10 category rows + 1 Total row ────────────────
    // Read exactly 11 rows, same 30-col window so indices align
    var tableData = sheet.getRange(7, 1, 11, 30).getValues();
    var categories    = [];
    var monthlyTotals = monthCols.map(function() { return 0; });

    tableData.forEach(function(row) {
      var catName = (row[0] || '').toString().trim();
      if (!catName) return; // skip blank rows

      var vals = monthCols.map(function(mc) {
        return parseFloat(row[mc.colIndex]) || 0;
      });

      // Row 17 label contains "total" (case-insensitive)
      if (catName.toLowerCase().indexOf('total') >= 0) {
        monthlyTotals = vals;
      } else {
        categories.push({ name: catName, values: vals });
      }
    });

    // ── Group by year ─────────────────────────────────────────────
    // yearGroups: { '2025': { months:[], indices:[], totals:[], catTotals:{} }, ... }
    var yearGroups = {};
    monthCols.forEach(function(mc, idx) {
      var yr = mc.year;
      if (!yearGroups[yr]) {
        yearGroups[yr] = { months: [], indices: [], totals: [], catTotals: {} };
        categories.forEach(function(c) { yearGroups[yr].catTotals[c.name] = 0; });
      }
      yearGroups[yr].months.push(mc.label);
      yearGroups[yr].indices.push(idx);
      yearGroups[yr].totals.push(monthlyTotals[idx] || 0);
      categories.forEach(function(c) {
        yearGroups[yr].catTotals[c.name] = (yearGroups[yr].catTotals[c.name] || 0) + (c.values[idx] || 0);
      });
    });

    return {
      months:        monthCols.map(function(m) { return m.label; }),
      monthlyTotals: monthlyTotals,
      categories:    categories,
      yearGroups:    yearGroups,
      latestMonth:   monthCols.length > 0 ? monthCols[monthCols.length - 1].label : ''
    };

  } catch (error) {
    throw new Error('Failed to load expense data: ' + error.message);
  }
}

function getExpensesForFriend(friendId) {
  var apiKey = 'Lo46GCiwIVipgzU3aV64dM5YysVZkLP3nY6vxtWv';
  var headers = { 'Authorization': 'Bearer ' + apiKey };
  var opts = { method: 'get', headers: headers, muteHttpExceptions: true };

  try {
    // ── 1. Current user ID ──
    var currentUserId = null;
    try {
      var meResp = UrlFetchApp.fetch('https://secure.splitwise.com/api/v3.0/get_current_user', opts);
      currentUserId = JSON.parse(meResp.getContentText()).user.id;
    } catch(e) {}

    // ── 2. Fetch expenses for this specific friend ──
    var url = 'https://secure.splitwise.com/api/v3.0/get_expenses?limit=20&offset=0&friend_id=' + friendId;
    var expResp = UrlFetchApp.fetch(url, opts);
    var expenses = JSON.parse(expResp.getContentText()).expenses || [];

    var transactions = expenses
      .filter(function(e) { return !e.deleted_at; })
      .slice(0, 10)
      .map(function(e) {
        var myNet = 0, paidByMe = false;
        if (currentUserId) {
          for (var i = 0; i < e.users.length; i++) {
            if (e.users[i].user_id === currentUserId) {
              myNet = parseFloat(e.users[i].net_balance || 0);
              paidByMe = parseFloat(e.users[i].paid_share || 0) > 0;
              break;
            }
          }
        }
        // All participant first names
        var participants = (e.users || []).map(function(u) {
          return (u.user && u.user.first_name) ? u.user.first_name : 'User';
        });
        return {
          id:           e.id,
          description:  e.description || '(no description)',
          date:         Utilities.formatDate(new Date(e.date), Session.getScriptTimeZone(), 'dd MMM yyyy'),
          cost:         parseFloat(e.cost),
          currency:     e.currency_code,
          myNet:        myNet,
          paidByMe:     paidByMe,
          isPayment:    e.payment === true,
          participants: participants
        };
      });

    return { success: true, transactions: transactions };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function showInvestmentDashboard() {
  const html = HtmlService.createHtmlOutputFromFile("investmentDashboard")
    .setWidth(2000)
    .setHeight(2000);
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

// ═══════════════════════════════════════════════════════════════
// CSE STOCK DASHBOARD
// ═══════════════════════════════════════════════════════════════

function showCSEDashboard() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('csiDashboard')
      .setWidth(2000)
      .setHeight(2000);
    SpreadsheetApp.getUi().showModalDialog(html, '📈 CSE Stock Dashboard');
  } catch(e) {}
}

// ── Stage 1: Instant – just reads sheet, zero API calls ─────────
function getCSESheetData() {
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName("UT_CRYPTO");
    var rows  = sheet.getRange("A20:G29").getValues();
    var stocks = [];
    rows.forEach(function(row) {
      var symbol = (row[0] || '').toString().trim();
      if (!symbol) return;
      stocks.push({
        symbol:      symbol,
        sheetPrice:  Number(row[1]) || 0,
        boughtPrice: Number(row[2]) || 0,
        qty:         Number(row[3]) || 0,
        name:        (row[4] || '').toString().trim(),
        sector:      (row[5] || '').toString().trim(),
        notes:       (row[6] || '').toString().trim()
      });
    });
    return { stocks: stocks };
  } catch(e) {
    return { stocks: [], error: e.message };
  }
}

// ── Stage 2: One live price per symbol (called per-card from HTML) ──
function getCSELivePrice(symbol) {
  try {
    var res  = UrlFetchApp.fetch("https://www.cse.lk/api/companyInfoSummery", {
      method: "post", payload: { symbol: symbol }, muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return { success: false, symbol: symbol };
    var json = JSON.parse(res.getContentText());
    if (!json || !json.reqSymbolInfo) return { success: false, symbol: symbol };
    var si = json.reqSymbolInfo;
    return {
      success:      true,
      symbol:       symbol,
      currentPrice: Number(si.lastTradedPrice)  || 0,
      change:       Number(si.change)           || Number(si.priceChange)     || 0,
      changePct:    Number(si.percentageChange) || Number(si.percentChange)   || 0,
      dayHigh:      Number(si.dayHigh)          || Number(si.highPrice)       || 0,
      dayLow:       Number(si.dayLow)           || Number(si.lowPrice)        || 0,
      volume:       Number(si.volume)           || Number(si.totalVolume)     || 0,
      pe:           Number(si.pe)               || Number(si.priceToEarnings) || 0,
      dividendYield:Number(si.dividendYield)    || 0,
      marketCap:    Number(si.marketCap)        || 0,
      yearHigh:     Number(si.yearHigh)         || Number(si['52WeekHigh'])   || 0,
      yearLow:      Number(si.yearLow)          || Number(si['52WeekLow'])    || 0,
      companyName:  si.companyName || si.name   || '',
      allFields:    si
    };
  } catch(e) {
    return { success: false, symbol: symbol, error: e.message };
  }
}

// ── Stage 3: Market + News + Dividends (lazy, on tab click) ─────
function getCSEMarketData() {
  var result = { market: {}, news: [], dividends: [] };
  var base   = "https://www.cse.lk/api/";

  var eps = [
    { key: 'summary',    path: 'marketSummary'          },
    { key: 'index',      path: 'allSharePriceIndex'      },
    { key: 'gainers',    path: 'topGainers'              },
    { key: 'losers',     path: 'topLosers'               },
    { key: 'mostActive', path: 'mostActive'              }
  ];
  eps.forEach(function(ep) {
    try {
      var r = UrlFetchApp.fetch(base + ep.path, { method: 'get', muteHttpExceptions: true });
      if (r.getResponseCode() === 200) {
        var body = r.getContentText().trim();
        if (body && body.charAt(0) !== '<') result.market[ep.key] = JSON.parse(body);
      }
    } catch(e) {}
  });
  try {
    var nr = UrlFetchApp.fetch(base + 'news', { method: 'get', muteHttpExceptions: true });
    if (nr.getResponseCode() === 200) {
      var nb = nr.getContentText().trim();
      if (nb && nb.charAt(0) !== '<') result.news = JSON.parse(nb);
    }
  } catch(e) {}
  // Try multiple dividend endpoint variants
  var divEndpoints = ['dividendAnnouncements','dividendAnnouncement','dividend',
                      'corporateActions','corporateAction','announcement'];
  for (var di = 0; di < divEndpoints.length; di++) {
    try {
      var dr = UrlFetchApp.fetch(base + divEndpoints[di], { method: 'get', muteHttpExceptions: true });
      if (dr.getResponseCode() === 200) {
        var db = dr.getContentText().trim();
        if (db && db.charAt(0) !== '<' && db.length > 5) {
          var parsed = JSON.parse(db);
          var list = Array.isArray(parsed) ? parsed : (parsed.data || parsed.list || parsed.dividends || parsed.results || []);
          if (list.length > 0) { result.dividends = list; break; }
        }
      }
    } catch(e) {}
    // also try as POST
    try {
      var drp = UrlFetchApp.fetch(base + divEndpoints[di], { method: 'post', payload: {}, muteHttpExceptions: true });
      if (drp.getResponseCode() === 200) {
        var dbp = drp.getContentText().trim();
        if (dbp && dbp.charAt(0) !== '<' && dbp.length > 5) {
          var parsedp = JSON.parse(dbp);
          var listp = Array.isArray(parsedp) ? parsedp : (parsedp.data || parsedp.list || parsedp.dividends || parsedp.results || []);
          if (listp.length > 0) { result.dividends = listp; break; }
        }
      }
    } catch(e) {}
  }

  return result;
}








// ── Per-symbol dividend fetch (tries multiple endpoints per symbol) ──
function getCSEDividendsForHoldings(symbols) {
  var base = "https://www.cse.lk/api/";
  var result = [];
  var now = new Date();
  var curYM = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM');

  var variants = ['dividendHistory','dividendAnnouncements','dividend','dividendAnnouncement','corporateActions'];

  symbols.forEach(function(sym) {
    var found = false;
    for (var i = 0; i < variants.length && !found; i++) {
      // Try JSON POST
      try {
        var r = UrlFetchApp.fetch(base + variants[i], {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify({ symbol: sym }),
          muteHttpExceptions: true
        });
        if (r.getResponseCode() === 200) {
          var body = r.getContentText().trim();
          if (body && (body.charAt(0) === '[' || body.charAt(0) === '{')) {
            var parsed = JSON.parse(body);
            var list = Array.isArray(parsed) ? parsed : (parsed.data || parsed.list || parsed.dividends || parsed.results || []);
            if (list.length > 0) {
              list.forEach(function(d) {
                if (!d.symbol) d.symbol = sym;
                result.push(d);
              });
              found = true;
            }
          }
        }
      } catch(e) {}
      if (found) break;
      // Try form POST
      try {
        var r2 = UrlFetchApp.fetch(base + variants[i], {
          method: 'post', payload: { symbol: sym }, muteHttpExceptions: true
        });
        if (r2.getResponseCode() === 200) {
          var body2 = r2.getContentText().trim();
          if (body2 && (body2.charAt(0) === '[' || body2.charAt(0) === '{')) {
            var parsed2 = JSON.parse(body2);
            var list2 = Array.isArray(parsed2) ? parsed2 : (parsed2.data || parsed2.list || parsed2.dividends || parsed2.results || []);
            if (list2.length > 0) {
              list2.forEach(function(d) {
                if (!d.symbol) d.symbol = sym;
                result.push(d);
              });
              found = true;
            }
          }
        }
      } catch(e) {}
    }
  });

  return { dividends: result, month: curYM };
}

// ── Gold price data (current scrape + UT_CRYPTO monthly history + price chart) ──
function getGoldData() {
  var result = { currentPrice: 0, updated: '', history: [], priceHistory: [], currentInvested: 0, lkrRate: 0 };

  // Fetch live price from ravijewellers.lk
  try {
    var html = UrlFetchApp.fetch('https://ravijewellers.lk/', { muteHttpExceptions: true }).getContentText();
    var m = html.match(/goldrate-rate[^>]*>.*?LKR\s*([\d,]+)/i)
             || html.match(/class="[^"]*gold[^"]*"[^>]*>[\s\S]{0,200}?LKR\s*([\d,]+)/i)
             || html.match(/LKR\s*([\d,]+)/);
    if (m) {
      result.currentPrice = Number(m[1].replace(/,/g, ''));
      result.updated = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy, HH:mm');
    }
  } catch(e) {}

  // Read from UT_CRYPTO sheet
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName('UT_CRYPTO');
    var c1 = sheet.getRange('C1').getValue();
    var c2 = sheet.getRange('C2').getValue();
    if (c1 && !result.currentPrice) result.currentPrice = Number(c1) || 0;
    result.currentInvested = Number(c2) || 0;

    // Monthly history rows 46+ : col A=date, col D=goldEarnings, col G=goldInvested
    var lastRow = sheet.getLastRow();
    if (lastRow >= 46) {
      var raw = sheet.getRange(46, 1, lastRow - 45, 7).getValues();
      result.history = raw
        .filter(function(r) { return r[0]; })
        .map(function(r) {
          return {
            month: Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'yyyy-MM'),
            label: Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'MMM yy'),
            goldInvested: Number(r[6]) || 0,
            goldEarnings: Number(r[3]) || 0
          };
        });
    }
  } catch(e) {}

  // Historical gold price chart: XAU/USD (Yahoo Finance) → LKR per gram
  try {
    // Get USD/LKR exchange rate
    var lkrRate = 320;
    try {
      var er = UrlFetchApp.fetch('https://open.er-api.com/v6/latest/USD', { muteHttpExceptions: true });
      if (er.getResponseCode() === 200) {
        var erJson = JSON.parse(er.getContentText());
        if (erJson.rates && erJson.rates.LKR) lkrRate = erJson.rates.LKR;
      }
    } catch(e2) {}
    result.lkrRate = lkrRate;

    // 6 months weekly gold price from Yahoo Finance (GC=F = Gold futures)
    var yf = UrlFetchApp.fetch(
      'https://query1.finance.yahoo.com/v8/finance/chart/GC=F?interval=1wk&range=6mo',
      { muteHttpExceptions: true }
    );
    if (yf.getResponseCode() === 200) {
      var yfJson = JSON.parse(yf.getContentText());
      var res0 = yfJson.chart && yfJson.chart.result && yfJson.chart.result[0];
      if (res0) {
        var timestamps = res0.timestamp || [];
        var closes = res0.indicators && res0.indicators.quote[0].close || [];
        result.priceHistory = [];
        for (var i = 0; i < timestamps.length; i++) {
          if (!closes[i]) continue;
          var d = new Date(timestamps[i] * 1000);
          result.priceHistory.push({
            label: Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd MMM'),
            usd:   Math.round(closes[i] * 100) / 100,
            lkr:   Math.round(closes[i] * lkrRate / 31.1035)
          });
        }
      }
    }
  } catch(e) {}

  return result;
}
