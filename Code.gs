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
    showExpenseDashboard();
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
    // Data starts at row 2, columns A–M (M = status JSON)
    var values = sheet.getRange(2, 1, lastRow - 1, 13).getValues();

    function parseAmt(val) {
      if (!val && val !== 0) return 0;
      var s = val.toString().trim();
      if (s === '' || s === 'NTM' || s === 'N T M') return 0;
      // Strip any NTM suffix mixed in (e.g. "NTM Stocks going down")
      if (s.toUpperCase().startsWith('NTM')) return 0;
      // FAILED cell → -1 so frontend can distinguish from NTM (0)
      if (s.toUpperCase().startsWith('FAILED') || s.toUpperCase() === 'FAILED') return -1;
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
        notes:         (row[11] || '').toString().trim(),
      status:        (function(v){ try{ return JSON.parse((v||'').toString().trim()||'{}'); }catch(e){ return {}; } })(row[12])
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

function saveInvestmentStatus(monthStr, statusJson) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Money Flow and invest');
    if (!sheet) throw new Error('Sheet "Money Flow and invest" not found');
    var vals  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      var raw = vals[i][0];
      var m   = raw instanceof Date
        ? Utilities.formatDate(raw, Session.getScriptTimeZone(), 'MMM/yyyy')
        : (raw || '').toString().trim();
      if (m === monthStr) {
        sheet.getRange(i + 2, 13).setValue(statusJson);
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    return { success: false, error: 'Month "' + monthStr + '" not found' };
  } catch(e) { return { success: false, error: e.message }; }
}

/**
 * Overwrites the investment amount cell with "FAILED" text.
 * Called only when the user explicitly sets a status to "Failed" in the dashboard UI.
 * Column mapping mirrors getMoneyFlowData() row[] indices (0-based → +1 for 1-based col).
 */
function writeInvestmentFailed(monthStr, key) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Money Flow and invest');
    if (!sheet) throw new Error('Sheet "Money Flow and invest" not found');

    // key → 1-based sheet column  (matches parseAmt row[] order in getMoneyFlowData)
    var keyColMap = { sentHome:2, cal:3, ndb:4, binance:5, stock:6, fd:7, gold:8 };
    var colIdx = keyColMap[key];
    if (!colIdx) throw new Error('Unknown investment key: ' + key);

    // Locate the row for this month
    var lastRow = sheet.getLastRow();
    var monthVals = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < monthVals.length; i++) {
      var raw = monthVals[i][0];
      var m   = raw instanceof Date
        ? Utilities.formatDate(raw, Session.getScriptTimeZone(), 'MMM/yyyy')
        : (raw || '').toString().trim();
      if (m === monthStr) {
        sheet.getRange(i + 2, colIdx).setValue('FAILED');
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    return { success: false, error: 'Month "' + monthStr + '" not found' };
  } catch(e) { return { success: false, error: e.message }; }
}

function saveMoneyFlowNotes(monthStr, notes) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Money Flow and invest');
    if (!sheet) throw new Error('Sheet "Money Flow and invest" not found');
    var vals  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      var raw = vals[i][0];
      var m   = raw instanceof Date
        ? Utilities.formatDate(raw, Session.getScriptTimeZone(), 'MMM/yyyy')
        : (raw || '').toString().trim();
      if (m === monthStr) {
        sheet.getRange(i + 2, 12).setValue(notes);
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    return { success: false, error: 'Month "' + monthStr + '" not found' };
  } catch(e) { return { success: false, error: e.message }; }
}

/**
 * Saves investment amounts for a given month from the Money Flow dashboard.
 * amountsJson: JSON string like {"sentHome":10000,"cal":5000,"ndb":0,...}
 * key → 1-based sheet column mapping matches getMoneyFlowData() indices.
 */
function saveMoneyFlowAmounts(monthStr, amountsJson, statusJson) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Money Flow and invest');
    if (!sheet) throw new Error('Sheet "Money Flow and invest" not found');

    var amounts = JSON.parse(amountsJson || '{}');
    var status = JSON.parse(statusJson || '{}');
    var keyColMap = { sentHome:2, cal:3, ndb:4, binance:5, stock:6, fd:7, gold:8 };

    // Colors
    var ntmColor = '#cfe2f3';    // Light blue for NTM
    var valueColor = '#fff2cc';  // Light yellow for pending values
    var doneColor = '#d9ead3';   // Light green for DONE
    var failedColor = '#f4cccc'; // Light red for Failed

    var vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      var raw = vals[i][0];
      var m   = raw instanceof Date
        ? Utilities.formatDate(raw, Session.getScriptTimeZone(), 'MMM/yyyy')
        : (raw || '').toString().trim();
      if (m === monthStr) {
        var rowNum = i + 2;
        Object.keys(amounts).forEach(function(key) {
          var colIdx = keyColMap[key];
          if (colIdx) {
            var val = amounts[key];
            var st = status[key] || '';
            var cell = sheet.getRange(rowNum, colIdx);
            // Write value and apply color based on status
            if (val === 0 || val === '' || val === null) {
              cell.setValue('NTM');
              cell.setBackground(ntmColor);
            } else {
              cell.setValue(val);
              if (st === 'DONE') {
                cell.setBackground(doneColor);
              } else if (st === 'Failed') {
                cell.setBackground(failedColor);
              } else {
                cell.setBackground(valueColor);
              }
            }
          }
        });
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    return { success: false, error: 'Month "' + monthStr + '" not found' };
  } catch(e) { return { success: false, error: e.message }; }
}

/**
 * Adds a new month row to the Money Flow sheet.
 * payloadJson: JSON string { month: "MMM/yyyy", amounts: {...}, notes: "" }
 */
function addMoneyFlowMonth(payloadJson) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Money Flow and invest');
    if (!sheet) throw new Error('Sheet "Money Flow and invest" not found');

    var payload = JSON.parse(payloadJson || '{}');
    var monthStr = payload.month || '';
    var amounts = payload.amounts || {};
    var income = payload.income || 0;
    var notes = payload.notes || '';

    if (!monthStr) throw new Error('Month is required');

    // Check if month already exists
    var lastRow = sheet.getLastRow();
    var monthVals = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < monthVals.length; i++) {
      var raw = monthVals[i][0];
      var m = raw instanceof Date
        ? Utilities.formatDate(raw, Session.getScriptTimeZone(), 'MMM/yyyy')
        : (raw || '').toString().trim();
      if (m === monthStr) {
        return { success: false, error: 'Month "' + monthStr + '" already exists' };
      }
    }

    // Parse month string to date (1st of month)
    var monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var parts = monthStr.split('/');
    var mIdx = monthNames.indexOf(parts[0]);
    var year = parseInt(parts[1]);
    var monthDate = new Date(year, mIdx, 1);

    // Column mapping: A=month, B=sentHome, C=cal, D=ndb, E=binance, F=stock, G=fd, H=gold, I=totalInvested, J=income, K=savingPct, L=notes, M=status
    var keyColMap = { sentHome:2, cal:3, ndb:4, binance:5, stock:6, fd:7, gold:8 };

    // Build row values
    var newRow = [];
    newRow[0] = monthDate; // A: month date

    // B-H: investment categories
    ['sentHome','cal','ndb','binance','stock','fd','gold'].forEach(function(key, idx) {
      var val = amounts[key] || 0;
      newRow[idx + 1] = val === 0 ? 'NTM' : val;
    });

    // I: totalInvested (formula or sum)
    var totalInvested = 0;
    ['cal','ndb','binance','stock','fd','gold'].forEach(function(key) {
      var v = amounts[key] || 0;
      if (v > 0) totalInvested += v;
    });
    newRow[8] = totalInvested; // I

    newRow[9] = income;  // J: income
    // K: savingPct = (totalInvested / income) * 100
    var savingPct = income > 0 ? Math.round((totalInvested / income) * 100) : 0;
    newRow[10] = savingPct; // K: saving percentage
    newRow[11] = notes; // L: notes
    newRow[12] = '{}'; // M: status (empty JSON)

    // Find last row with a month value (Date or "MMM/yyyy" format)
    var lastMonthRow = 1; // default to header row
    for (var j = 0; j < monthVals.length; j++) {
      var cellVal = monthVals[j][0];
      var isMonth = false;

      // Check if it's a Date object
      if (cellVal instanceof Date && !isNaN(cellVal.getTime())) {
        isMonth = true;
      } else {
        // Check if string contains month name
        var cellStr = (cellVal || '').toString().trim();
        for (var mi = 0; mi < monthNames.length; mi++) {
          if (cellStr.indexOf(monthNames[mi]) >= 0) {
            isMonth = true;
            break;
          }
        }
      }

      if (isMonth) {
        lastMonthRow = j + 2; // j is 0-based from row 2, so actual row = j + 2
      }
    }

    // Insert after the last month row
    var insertRow = lastMonthRow + 1;
    if (insertRow <= lastRow) {
      sheet.insertRowBefore(insertRow);
    }

    sheet.getRange(insertRow, 1, 1, 13).setValues([newRow]);

    // Apply custom background colors
    var ntmColor = '#cfe2f3';    // Light blue for NTM
    var valueColor = '#fff2cc';  // Light yellow for values with pending status

    // Column A (month) - no special color
    sheet.getRange(insertRow, 1).setBackground('#ffffff');

    // Columns B-H: investment categories - color based on NTM or value
    ['sentHome','cal','ndb','binance','stock','fd','gold'].forEach(function(key, idx) {
      var colNum = idx + 2; // B=2, C=3, etc.
      var val = amounts[key] || 0;
      if (val === 0) {
        sheet.getRange(insertRow, colNum).setBackground(ntmColor); // NTM - light blue
      } else {
        sheet.getRange(insertRow, colNum).setBackground(valueColor); // Value - light yellow
      }
    });

    // Columns I-M: totals and notes - white/default
    sheet.getRange(insertRow, 9, 1, 5).setBackground('#ffffff');

    SpreadsheetApp.flush();

    return { success: true, month: monthStr };
  } catch(e) {
    return { success: false, error: e.message };
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


// Function to get card data from sheet
// Function to get card data from the sheet
function getCardData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('CC_SW_CL_INST');
    
    if (!sheet) {
      throw new Error('Sheet "CC_SW_CL_INST" not found');
    }
    
    // Get data from range A4:I6 (data starts at row 4, col I = Available to Spend)
    var range = sheet.getRange('A4:I7');
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

    // ── EPF + ETF total: find "TOTAL" row in col A, read col F ──────
    var epfEtf = 0;
    try {
      var epfSheet = ss.getSheetByName('EPF ETF');
      if (epfSheet) {
        var epfLastRow = epfSheet.getLastRow();
        var epfColA = epfSheet.getRange(1, 1, epfLastRow, 1).getValues();
        for (var ei = 0; ei < epfColA.length; ei++) {
          if ((epfColA[ei][0] + '').trim().toLowerCase() === 'total') {
            epfEtf = parseFloat(epfSheet.getRange(ei + 1, 6).getValue()) || 0;
            break;
          }
        }
      }
    } catch(ee) {
      Logger.log('EPF ETF read failed: ' + ee.message);
    }

    // ── Aims / Remaining: A1:D15 (row1=headers, row15=total, rows2-14=data) ──
    var aims = [];
    try {
      var aimVals = sheet.getRange('A1:D15').getValues();
      // row 0 = headers, rows 1-13 = data, row 14 = total
      for (var ai = 1; ai <= 13; ai++) {
        var row = aimVals[ai];
        var cat  = (row[0] + '').trim();
        var amt  = parseFloat(row[1]) || 0;
        var aim  = parseFloat(row[2]) || 0;
        var rem  = parseFloat(row[3]) || 0;
        if (!cat || cat === '') continue;
        if (aim <= 0 && rem <= 0) continue; // skip rows with no aim and no remaining
        aims.push({ category: cat, amount: amt, aim: aim, remaining: rem });
      }
      // Total row (A15:D15)
      var totRow = aimVals[14];
      var aimsTotal = {
        amount:    parseFloat(totRow[1]) || 0,
        aim:       parseFloat(totRow[2]) || 0,
        remaining: parseFloat(totRow[3]) || 0
      };
    } catch(ea) {
      Logger.log('Aims data read failed: ' + ea.message);
      var aimsTotal = { amount: 0, aim: 0, remaining: 0 };
    }

    return { assets: assets, growth: growth, epfEtf: epfEtf, aims: aims, aimsTotal: aimsTotal };
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

// ── Bank Balance Management ──
function getBankBalances() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Portfolio');
    if (!sheet) throw new Error('Sheet "Portfolio" not found');
    
    var data = sheet.getRange("G1:I10").getValues();
    var bankBalances = [];
    
    for (var i = 0; i < data.length; i++) {
      var type = (data[i][0] || '').toString().trim();
      if (type.toUpperCase() === 'BANK') {
        bankBalances.push({
          row: i + 1, // 1-based row relative to G1
          source: (data[i][1] || '').toString().trim(),
          amount: parseFloat(data[i][2]) || 0
        });
      }
    }
    return bankBalances;
  } catch (e) {
    throw new Error('Failed to load bank balances: ' + e.message);
  }
}

function updateBankBalances(updates) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Portfolio');
    if (!sheet) throw new Error('Sheet "Portfolio" not found');
    
    updates.forEach(function(update) {
      if (update.source !== undefined) {
        sheet.getRange(update.row, 8).setValue(update.source);
      }
      if (update.amount !== undefined) {
        sheet.getRange(update.row, 9).setValue(update.amount);
      }
    });
    
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    throw new Error('Failed to update bank balances: ' + e.message);
  }
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
  const headers = { 'Authorization': 'Bearer ' + apiKey };
  const opts = { method: 'get', headers: headers };

  try {
    // ── 1. Friend balances ──
    const friendsResp = UrlFetchApp.fetch('https://secure.splitwise.com/api/v3.0/get_friends', opts);
    const friends = JSON.parse(friendsResp.getContentText()).friends;

    const balanceData = [];
    const totals = {};

    friends.forEach(function(friend) {
      const fullName = (friend.first_name + ' ' + (friend.last_name || '')).trim();
      friend.balance.forEach(function(bal) {
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

    // ── 2. Current user ID ──
    var currentUserId = null;
    try {
      const meResp = UrlFetchApp.fetch('https://secure.splitwise.com/api/v3.0/get_current_user', opts);
      currentUserId = JSON.parse(meResp.getContentText()).user.id;
    } catch(e) {}

    // ── 3. Recent expenses — fetch latest 15 ──
    var recentTransactions = [];
    try {
      var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
      var expUrl  = 'https://secure.splitwise.com/api/v3.0/get_expenses?limit=15&offset=0';
      var expResp = UrlFetchApp.fetch(expUrl, opts);
      var allTxns = (JSON.parse(expResp.getContentText()).expenses || [])
                      .filter(function(e) { return !e.deleted_at; })
                      .slice(0, 15);
      recentTransactions = allTxns.map(function(e) {
        var myNet = 0, paidByMe = false;
        if (currentUserId) {
          for (var i = 0; i < e.users.length; i++) {
            if (e.users[i].user_id === currentUserId) {
              myNet    = parseFloat(e.users[i].net_balance || 0);
              paidByMe = parseFloat(e.users[i].paid_share  || 0) > 0;
              break;
            }
          }
        }
        var rawDate = new Date(e.date);
        return {
          id:          e.id,
          description: e.description || '(no description)',
          date:        Utilities.formatDate(rawDate, tz, 'dd MMM yyyy'),
          rawDate:     rawDate.getTime(),
          cost:        parseFloat(e.cost),
          currency:    e.currency_code,
          myNet:       myNet,
          paidByMe:    paidByMe,
          isPayment:   e.payment === true
        };
      });
    } catch(e) {}

    // ── 4. Validate LKR balance ──
    const portfolioSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Portfolio');
    const sheetBalance = portfolioSheet ? portfolioSheet.getRange('B14').getValue() : 0;
    const lkrTotal = totals['LKR'] || 0;
    const validationResult = validateBalance(lkrTotal, sheetBalance);

    return {
      balanceData: balanceData,
      recentTransactions: recentTransactions,
      validationResult: validationResult
    };

  } catch (error) {
    Logger.log('getSplitwiseData error: ' + error);
    return { balanceData: [], recentTransactions: [], validationResult: null };
  }
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
      if (h.toUpperCase() === 'TOTAL') continue;                    // skip grand-total column
      if (h.toLowerCase().indexOf('control') >= 0) continue;       // skip Control Plan Amount column
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

/* ─────────────────────────────────────────────────────────────────
   getCurrentMonthExpensesList
   Returns every non-zero expense entry for the current month
   from the daily rows in "Monthly Expences" sheet,
   excluding Credit Card and Splitwise columns.
───────────────────────────────────────────────────────────────── */
function getCurrentMonthExpensesList() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Monthly Expences');
    if (!sheet) return { success: false, error: 'Sheet "Monthly Expences" not found' };

    var tz       = ss.getSpreadsheetTimeZone();
    var now      = new Date();
    var curMonth = parseInt(Utilities.formatDate(now, tz, 'M'), 10); // 1-12
    var curYear  = parseInt(Utilities.formatDate(now, tz, 'yyyy'), 10);

    // Column definitions (1-based), exclude CC (AE=31,AF=32,AG=33,AH=34,AI=35)
    // and SW (T=20, U=21, W=23)
    var CATS = [
      { name: 'Food',               amtCol: 3,  refCol: 4  },  // C / D
      { name: 'Supermarket',        amtCol: 5,  refCol: 6  },  // E / F
      { name: 'Uber',               amtCol: 7,  refCol: 0  },  // G
      { name: 'Uber Work',          amtCol: 8,  refCol: 0  },  // H
      { name: 'Movies & Outing',    amtCol: 10, refCol: 9  },  // J / I
      { name: 'Other',              amtCol: 12, refCol: 11 },  // L / K
      { name: 'Bus Fair',           amtCol: 13, refCol: 0  },  // M
      { name: 'Party', amtCol: 15, refCol: 14 },  // O / N
      { name: 'Dress & Appearance',              amtCol: 17, refCol: 16 },  // Q / P
      { name: 'Rent',               amtCol: 18, refCol: 0  }   // R
    ];

    var lastRow = sheet.getLastRow();
    var numCols = 35; // A through AI
    if (numCols > sheet.getLastColumn()) numCols = sheet.getLastColumn();
    var allData = sheet.getRange(1, 1, lastRow, numCols).getValues();

    var entries = [];

    for (var r = 0; r < allData.length; r++) {
      var row  = allData[r];
      // Normalise col A to "M-d" using shared helper
      var dateKey = _cellStr(row[0], true).trim();
      if (!dateKey) continue;

      // Match "M-d" pattern and check current month
      var parts = dateKey.match(/^(\d{1,2})-(\d{1,2})$/);
      if (!parts) continue;
      var rowMonth = parseInt(parts[1], 10);
      var rowDay   = parseInt(parts[2], 10);
      if (rowMonth !== curMonth) continue;

      // Extract non-zero values per category
      CATS.forEach(function(cat) {
        var rawAmt = row[cat.amtCol - 1]; // 0-indexed
        var amt = 0;
        if (typeof rawAmt === 'number') {
          amt = rawAmt;
        } else {
          // Could be a formula string like "=800+350" – evaluate naively
          amt = parseFloat(String(rawAmt).replace(/[^0-9.+-]/g, '')) || 0;
        }
        if (isNaN(amt) || amt <= 0) return;

        var ref = '';
        if (cat.refCol > 0) {
          ref = String(row[cat.refCol - 1] || '').trim();
        }

        entries.push({
          day:       rowDay,
          category:  cat.name,
          reference: ref,
          amount:    amt
        });
      });
    }

    // Sort by day ascending, then category
    entries.sort(function(a, b) {
      return a.day !== b.day ? a.day - b.day : a.category.localeCompare(b.category);
    });

    // ── Read Control Plan from summary rows 6-16 ──
    var controlPlan = {};
    try {
      var cpHdr = sheet.getRange(6, 1, 1, sheet.getLastColumn()).getValues()[0];
      var cpCol = -1;
      for (var ci = 0; ci < cpHdr.length; ci++) {
        if (String(cpHdr[ci]).toLowerCase().indexOf('control') >= 0) { cpCol = ci; break; }
      }
      if (cpCol >= 0) {
        var cpRows = sheet.getRange(7, 1, 10, cpCol + 1).getValues(); // 10 category rows
        // Build raw map keyed by normalized name
        var cpRaw = {};
        cpRows.forEach(function(r) {
          var nm = String(r[0] || '').toLowerCase().replace(/\s+/g, ' ').trim();
          if (nm) cpRaw[nm] = parseFloat(r[cpCol]) || 0;
        });
        // Map to CATS names: exact match first, then first-word partial, then row-order fallback
        CATS.forEach(function(cat, idx) {
          var nm  = cat.name.toLowerCase();
          var val = 0;
          if (cpRaw[nm] !== undefined) {
            val = cpRaw[nm];
          } else {
            var fw = nm.split(' ')[0];
            Object.keys(cpRaw).forEach(function(k) {
              if (k.indexOf(fw) >= 0) val = cpRaw[k];
            });
          }
          // Row-order fallback: if still 0 and sheet row exists
          if (val === 0 && cpRows[idx]) {
            val = parseFloat(cpRows[idx][cpCol]) || 0;
          }
          controlPlan[cat.name] = val;
        });
      }
    } catch(e2) { /* control plan unavailable — not critical */ }

    // ── Last month totals from summary rows ──
    // Column order from right: Control Plan Amount | Total | Current Month | Prev Month | ...
    var lastMonthTotals = {};
    try {
      var hdrFull = sheet.getRange(6, 1, 1, sheet.getLastColumn()).getValues()[0];
      // Find Control Plan Amount column (0-based)
      var cpColIdx2 = -1;
      for (var ci2 = 0; ci2 < hdrFull.length; ci2++) {
        if (String(hdrFull[ci2] || '').toLowerCase().indexOf('control') >= 0) {
          cpColIdx2 = ci2; break;
        }
      }
      // prevMonthCol = Control - 3
      var prevColIdx = cpColIdx2 >= 3 ? cpColIdx2 - 3 : -1;
      if (prevColIdx >= 1) {
        var sumRows = sheet.getRange(7, 1, 10, prevColIdx + 1).getValues();
        // Build raw map keyed by normalised sheet name
        var lmtRaw = {};
        sumRows.forEach(function(row) {
          var nm = String(row[0] || '').trim();
          if (!nm || nm.toLowerCase().indexOf('total') >= 0) return;
          lmtRaw[nm.toLowerCase().replace(/\s+/g,' ')] = { raw: nm, val: parseFloat(row[prevColIdx]) || 0 };
        });
        // Map to CATS names: exact → first-word partial → row-order fallback (same as controlPlan)
        CATS.forEach(function(cat, idx) {
          var key = cat.name.toLowerCase().replace(/\s+/g,' ');
          var entry = lmtRaw[key];
          if (!entry) {
            var fw = key.split(' ')[0];
            Object.keys(lmtRaw).forEach(function(k) { if (k.indexOf(fw) >= 0) entry = lmtRaw[k]; });
          }
          var val = entry ? entry.val : 0;
          // Row-order fallback
          if (val === 0 && sumRows[idx]) val = parseFloat(sumRows[idx][prevColIdx]) || 0;
          if (val > 0) lastMonthTotals[cat.name] = val;
        });
      }
    } catch(e3) {}

    return { success: true, month: curMonth, year: curYear, entries: entries, controlPlan: controlPlan, lastMonthTotals: lastMonthTotals };
  } catch(err) {
    return { success: false, error: err.message || String(err) };
  }
}

function saveControlPlan(updates) {
  // updates: { "Food": 1200, "Supermarket": 800, ... }
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Monthly Expences');
    if (!sheet) return { success: false, error: 'Sheet not found' };

    // Find the Control Plan column in row 6 header
    var hdrRow = sheet.getRange(6, 1, 1, sheet.getLastColumn()).getValues()[0];
    var cpCol  = -1;
    for (var ci = 0; ci < hdrRow.length; ci++) {
      if (String(hdrRow[ci]).toLowerCase().indexOf('control') >= 0) { cpCol = ci + 1; break; } // 1-based
    }
    if (cpCol < 0) return { success: false, error: 'Control Plan column not found in row 6' };

    // Read category names from rows 7-16 (col A)
    var catNames = sheet.getRange(7, 1, 10, 1).getValues(); // 10 rows
    for (var ri = 0; ri < catNames.length; ri++) {
      var sheetName = String(catNames[ri][0] || '').trim();
      if (!sheetName) continue;
      var sheetKey  = sheetName.toLowerCase().replace(/\s+/g, ' ');
      // Try exact match first
      var matched = false;
      Object.keys(updates).forEach(function(catName) {
        if (matched) return;
        var updKey = catName.toLowerCase().replace(/\s+/g, ' ');
        if (updKey === sheetKey) {
          sheet.getRange(7 + ri, cpCol).setValue(updates[catName]);
          matched = true;
        }
      });
      // First-word partial if not matched
      if (!matched) {
        var fw = sheetKey.split(' ')[0];
        Object.keys(updates).forEach(function(catName) {
          if (matched) return;
          var updKey = catName.toLowerCase();
          if (updKey.indexOf(fw) >= 0 || fw.indexOf(updKey.split(' ')[0]) >= 0) {
            sheet.getRange(7 + ri, cpCol).setValue(updates[catName]);
            matched = true;
          }
        });
      }
    }
    SpreadsheetApp.flush();
    return { success: true };
  } catch(err) {
    return { success: false, error: err.message || String(err) };
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

    // ── 2. Fetch ALL expenses for this friend (up to 200 per Splitwise API limit) ──
    var allExpenses = [];
    var offset = 0;
    var pageLimit = 200;
    while (true) {
      var url = 'https://secure.splitwise.com/api/v3.0/get_expenses?limit=' + pageLimit
              + '&offset=' + offset + '&friend_id=' + friendId;
      var expResp = UrlFetchApp.fetch(url, opts);
      var page = JSON.parse(expResp.getContentText()).expenses || [];
      var active = page.filter(function(e) { return !e.deleted_at; });
      allExpenses = allExpenses.concat(active);
      if (page.length < pageLimit) break;   // last page
      offset += pageLimit;
      if (offset >= 1000) break;            // safety cap
    }

    var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    var transactions = allExpenses.map(function(e) {
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
        var participants = (e.users || []).map(function(u) {
          return (u.user && u.user.first_name) ? u.user.first_name : 'User';
        });
        var rawDate = new Date(e.date);
        return {
          id:           e.id,
          description:  e.description || '(no description)',
          date:         Utilities.formatDate(rawDate, tz, 'dd MMM yyyy'),
          rawDate:      rawDate.getTime(),   // ms timestamp for suspect analysis
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
// UNIT TRUST DATA
// ═══════════════════════════════════════════════════════════════

function getUTData() {
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName('UT_CRYPTO');
    if (!sheet) throw new Error('UT_CRYPTO sheet not found');
    // Rows 8–15, cols A–E: Fund Name, Invested, Units, Unit Price, Earnings
    var lastRow = Math.min(sheet.getLastRow(), 15);
    if (lastRow < 8) return { funds: [] };
    var rows = sheet.getRange(8, 1, lastRow - 7, 5).getValues();
    var funds = [];
    rows.forEach(function(row, i) {
      var name = (row[0] || '').toString().trim();
      if (!name || name.toLowerCase() === 'total') return;
      // derive company from first word (CAL, NDB, etc.)
      var company = name.split(' ')[0].toUpperCase();
      funds.push({
        rowIndex: i,           // 0-based offset from row 8
        name:      name,
        company:   company,
        invested:  Number((row[1] || '').toString().replace(/,/g,'')) || 0,
        units:     Number((row[2] || '').toString().replace(/,/g,'')) || 0,
        unitPrice: Number((row[3] || '').toString().replace(/,/g,'')) || 0,
        earnings:  Number((row[4] || '').toString().replace(/,/g,'')) || 0
      });
    });
    return { funds: funds };
  } catch(e) {
    return { funds: [], error: e.message };
  }
}

// Saves updated invested amount and units for a single fund row back to the sheet.
// rowIndex is 0-based (row 8 = index 0).
function saveUTEntry(rowIndex, invested, units) {
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName('UT_CRYPTO');
    if (!sheet) throw new Error('UT_CRYPTO sheet not found');
    var sheetRow = 8 + Number(rowIndex);
    if (sheetRow < 8 || sheetRow > 15) throw new Error('Invalid row index: ' + rowIndex);
    var inv = Number(invested);
    var uni = Number(units);
    if (isNaN(inv) || isNaN(uni)) throw new Error('Invalid numbers supplied');
    sheet.getRange(sheetRow, 2).setValue(inv);  // col B = Invested
    sheet.getRange(sheetRow, 3).setValue(uni);  // col C = Units
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ═══════════════════════════════════════════════════════════════
// Investment DASHBOARD
// ═══════════════════════════════════════════════════════════════

function showInvestmentDashboard() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('investmentDashboard')
      .setWidth(2000)
      .setHeight(2000);
    SpreadsheetApp.getUi().showModalDialog(html, '📈 Investment Dashboard');
  } catch(e) {}
}

/**
 * Saves stock buying price and quantity to UT_CRYPTO sheet.
 * rowIdx: 0-based index (0 = row 20, 1 = row 21, etc.)
 * boughtPrice: column C, qty: column D
 */
function saveStockData(symbol, rowIdx, boughtPrice, qty) {
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName("UT_CRYPTO");
    if (!sheet) throw new Error('UT_CRYPTO sheet not found');

    var sheetRow = 20 + Number(rowIdx);
    if (sheetRow < 20 || sheetRow > 29) throw new Error('Invalid row index');

    // Verify symbol matches
    var currentSymbol = sheet.getRange(sheetRow, 1).getValue();
    if (currentSymbol !== symbol) {
      throw new Error('Symbol mismatch: expected ' + symbol + ', found ' + currentSymbol);
    }

    // Update bought price (col C = 3) and qty (col D = 4)
    sheet.getRange(sheetRow, 3).setValue(Number(boughtPrice) || 0);
    sheet.getRange(sheetRow, 4).setValue(Number(qty) || 0);

    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

/**
 * Save all stock holdings data at once
 * updatesJson: JSON array of {symbol, rowIdx, bp, qty}
 */
function saveAllStockData(updatesJson) {
  try {
    var updates = JSON.parse(updatesJson);
    var sheet = SpreadsheetApp.getActive().getSheetByName("UT_CRYPTO");
    if (!sheet) throw new Error('UT_CRYPTO sheet not found');

    updates.forEach(function(u) {
      var sheetRow = 20 + Number(u.rowIdx);
      if (sheetRow >= 20 && sheetRow <= 29) {
        sheet.getRange(sheetRow, 3).setValue(Number(u.bp) || 0);
        sheet.getRange(sheetRow, 4).setValue(Number(u.qty) || 0);
      }
    });

    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
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

// ── Fixed Deposits: reads Portfolio sheet rows 21+ columns A–I ──────
function getFDData() {
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName('Portfolio');
    if (!sheet) throw new Error('Portfolio sheet not found');
    var lastRow = sheet.getLastRow();
    if (lastRow < 21) return { fds: [] };
    var rows = sheet.getRange(21, 1, lastRow - 20, 9).getValues();
    var fds  = [];
    var tz   = Session.getScriptTimeZone();
    rows.forEach(function(row) {
      var id = row[0];
      if (!id || isNaN(Number(id))) return; // skip header / Total row / blank
      var opened = row[3] ? new Date(row[3]) : null;
      var end    = row[4] ? new Date(row[4]) : null;
      fds.push({
        id:       Number(id),
        bank:     (row[1] || '').toString().trim().toUpperCase(),
        amount:   Number((row[2] || '').toString().replace(/,/g, '')) || 0,
        opened:   opened ? Utilities.formatDate(opened, tz, 'yyyy-MM-dd') : '',
        end:      end    ? Utilities.formatDate(end,    tz, 'yyyy-MM-dd') : '',
        type:     (row[5] || '').toString().trim(),
        interest: Number((row[6] || '').toString().replace(/,/g, '')) || 0,
        status:   (row[7] || '').toString().trim(),
        remarks:  (row[8] || '').toString().trim()
      });
    });
    return { fds: fds };
  } catch(e) {
    return { fds: [], error: e.message };
  }
}

// ── Update FD status and/or remarks for a given FD id ─────────────
function updateFDItem(fdId, newStatus, newRemarks) {
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName('Portfolio');
    if (!sheet) throw new Error('Portfolio sheet not found');
    var lastRow = sheet.getLastRow();
    if (lastRow < 21) throw new Error('No FD rows found');
    var rows = sheet.getRange(21, 1, lastRow - 20, 9).getValues();
    for (var i = 0; i < rows.length; i++) {
      var id = rows[i][0];
      if (!id || isNaN(Number(id))) continue;
      if (Number(id) === Number(fdId)) {
        var sheetRow = 21 + i;
        // col H (8) = status, col I (9) = remarks
        if (newStatus !== undefined && newStatus !== null) {
          sheet.getRange(sheetRow, 8).setValue(newStatus);
        }
        if (newRemarks !== undefined && newRemarks !== null) {
          sheet.getRange(sheetRow, 9).setValue(newRemarks);
        }
        SpreadsheetApp.flush();
        return { ok: true };
      }
    }
    return { ok: false, error: 'FD id ' + fdId + ' not found' };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Add a new FD row to Portfolio sheet, auto-incrementing ID ────────
function addFDItem(bank, amount, opened, end, type, interest, status, remarks) {
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName('Portfolio');
    if (!sheet) throw new Error('Portfolio sheet not found');
    
    var lastRow = sheet.getLastRow();
    var nextId = 1;
    var insertRow = lastRow + 1; // Default to end

    if (lastRow >= 21) {
      var range = sheet.getRange(21, 1, lastRow - 20 + 1, 1); // Include potential last row
      var values = range.getValues();
      
      for (var i = 0; i < values.length; i++) {
        var val = values[i][0];
        // Check for Total row
        if (val === 'Total' || (typeof val === 'string' && val.toLowerCase().indexOf('total') > -1)) {
          insertRow = 21 + i;
          break;
        }
        // Auto-increment ID logic
        var n = Number(val);
        if (!isNaN(n) && n >= nextId) nextId = n + 1;
      }
    }

    // Insert new row if we found a Total row or similar
    if (insertRow <= lastRow) {
      sheet.insertRowBefore(insertRow);
    } else {
      insertRow = lastRow + 1;
    }

    var openedDate = opened ? new Date(opened) : '';
    var endDate    = end    ? new Date(end)    : '';
    
    sheet.getRange(insertRow, 1).setValue(nextId);
    sheet.getRange(insertRow, 2).setValue((bank    || '').toUpperCase());
    sheet.getRange(insertRow, 3).setValue(Number(amount)   || 0);
    if (openedDate) sheet.getRange(insertRow, 4).setValue(openedDate);
    if (endDate)    sheet.getRange(insertRow, 5).setValue(endDate);
    sheet.getRange(insertRow, 6).setValue(type    || '');
    sheet.getRange(insertRow, 7).setValue(Number(interest) || 0);
    sheet.getRange(insertRow, 8).setValue(status  || 'In Progress');
    sheet.getRange(insertRow, 9).setValue(remarks || '');
    
    SpreadsheetApp.flush();
    return { ok: true, id: nextId };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Stage 2 (NEW): Batch fetch all stocks from LOLC Stock Screener ──
function getLOLCScreenerAll() {
  var epoch = new Date().getTime();
  var url   = 'https://www.lolcsecurities.lk/api/stock-screener/?s=' + epoch;
  var res   = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      'Accept':           'application/json, text/javascript, */*; q=0.01',
      'User-Agent':       'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
      'X-Requested-With': 'XMLHttpRequest',
      'Referer':          'https://www.lolcsecurities.lk/stock-screener/'
    }
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('LOLC screener HTTP ' + res.getResponseCode());
  }
  var json = JSON.parse(res.getContentText());
  var data = json.data || [];

  function pNum(v) {
    if (!v || v === '-') return 0;
    return parseFloat(String(v).replace(/,/g, '')) || 0;
  }
  function pPct(v) {
    if (!v || v === '-') return 0;
    return parseFloat(String(v).replace(/%/g, '')) || 0;
  }

  return data.map(function(row) {
    var price = pNum(row['Market Price (LKR)']);
    return {
      success:        price > 0,
      symbol:         (row['Company Tiker'] || '').trim(),
      companyName:    (row['Company Name']  || '').trim(),
      sector:         (row['Sector']        || '').trim(),
      currentPrice:   price,
      marketCap:      pNum(row['Market Capitalization (LKR Mn)']),
      foreignHolding: pPct(row['Foreign Holding%']),
      earnings:       pNum(row['4QT Earnings (LKR Mn)']),
      pe:             pNum(row['PE (x)']),
      sectorPe:       pNum(row['Sector PE (x)']),
      pbv:            pNum(row['PBV (x)']),
      sectorPbv:      pNum(row['Sector PBV (x)']),
      dy:             pPct(row['DY (%)']),
      dps:            pNum(row['DPS (LKR)']),
      eps:            pNum(row['EPS 4QT (LKR)']),
      nav:            pNum(row['NAV (LKR)']),
      roe:            pPct(row['ROE (%)']),
      allFields:      row
    };
  });
}

// ── Stage 2 (LEGACY): One live price per symbol (fallback only) ──
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








// ── LOLC Dividend Calendar CSV — future/today payments only ──────────
function getDividendCalendarCSV() {
  var url = 'https://www.lolcsecurities.lk/dividend-calendar/dividends_db.csv';
  var res = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { 'User-Agent': 'Mozilla/5.0' }
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('LOLC dividend CSV returned HTTP ' + res.getResponseCode());
  }
  var raw = res.getContentText('UTF-8');
  var lines = raw.split('\n');
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  // Keep records from up to 60 days in the past (≈ 2 months back)
  var cutoff = new Date(today.getTime() - 60 * 24 * 60 * 60 * 1000);
  var header = lines[0];
  var kept = [header];
  for (var i = 1; i < lines.length; i++) {
    var ln = lines[i].trim();
    if (!ln) continue;
    // D_PAY is the 3rd column (index 2); format yyyy-MM-dd (ISO)
    var cols = ln.split(',');
    var payRaw = cols[2] ? cols[2].replace(/"/g, '').trim() : '';
    var parts = payRaw.split('-');
    if (parts.length === 3) {
      var payDate = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
      if (payDate < cutoff) continue; // skip records older than 2 months
    }
    kept.push(ln);
  }
  return kept.join('\n');
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
            lkr:   Math.round(closes[i] * lkrRate / 31.1035 * (22/24)) // 22K per gram
          });
        }
      }
    }
  } catch(e) {}

  return result;
}


/* ═══════════════════════════════════════════════════════════════
   SPLITWISE — CREATE EXPENSE
   payload: { description, amount, currency, date,
              splits:[{userId,share,paid,name}],
              receiptBase64 }
   splits must include a "me" entry (userId==='me') and friend entries.
   The person who paid (paid > 0) is always the current user.
═══════════════════════════════════════════════════════════════ */
function createSplitwiseExpense(payload) {
  try {
    var apiKey = 'Lo46GCiwIVipgzU3aV64dM5YysVZkLP3nY6vxtWv';
    var headers = { 'Authorization': 'Bearer ' + apiKey };

    // 1. Get current user ID
    var meResp = UrlFetchApp.fetch('https://secure.splitwise.com/api/v3.0/get_current_user', { method:'get', headers:headers });
    var meData = JSON.parse(meResp.getContentText());
    var meId   = meData.user && meData.user.id ? meData.user.id : null;
    if (!meId) return { success:false, error:'Could not retrieve current user ID' };

    // 2. Build form payload for create_expense
    var total   = parseFloat(payload.amount) || 0;
    var splits  = payload.splits || [];
    var params  = {};
    params['cost']          = total.toFixed(2);
    params['description']   = String(payload.description || 'Expense');
    params['currency_code'] = String(payload.currency || 'LKR');
    params['date']          = String(payload.date || new Date().toISOString().slice(0,10));
    params['group_id']      = '0'; // non-group expense

    var userIdx = 0;
    splits.forEach(function(s) {
      var uid = String(s.userId) === 'me' ? String(meId) : String(s.userId);
      params['users__' + userIdx + '__user_id']    = uid;
      params['users__' + userIdx + '__paid_share'] = (parseFloat(s.paid) || 0).toFixed(2);
      params['users__' + userIdx + '__owed_share'] = (parseFloat(s.share) || 0).toFixed(2);
      userIdx++;
    });

    // 3. Submit expense (JSON body is simpler than multipart for non-image)
    var body = JSON.stringify(params);
    var opts = {
      method:      'post',
      headers:     Object.assign({}, headers, { 'Content-Type': 'application/json' }),
      payload:     body,
      muteHttpExceptions: true
    };
    var resp    = UrlFetchApp.fetch('https://secure.splitwise.com/api/v3.0/create_expense', opts);
    var result  = JSON.parse(resp.getContentText());
    var code    = resp.getResponseCode();

    if (code !== 200 && code !== 201) {
      var errMsg = (result.errors && (result.errors.base||[]).join(', ')) || ('HTTP '+code);
      Logger.log('createSplitwiseExpense error: ' + errMsg);
      return { success:false, error: errMsg };
    }

    var expenses = result.expenses || [];
    if (!expenses.length) return { success:false, error:'No expense returned by API' };

    var expId = expenses[0].id;

    // 4. Attach receipt image if provided (multipart)
    if (payload.receiptBase64 && expId) {
      try {
        var b64 = payload.receiptBase64;
        // strip data URL prefix if present
        var commaIdx = b64.indexOf(',');
        if (commaIdx >= 0) b64 = b64.slice(commaIdx + 1);
        var mimeMatch = payload.receiptBase64.match(/^data:([^;]+);/);
        var mime = mimeMatch ? mimeMatch[1] : 'image/jpeg';
        var imgBytes = Utilities.base64Decode(b64);
        var blob = Utilities.newBlob(imgBytes, mime, 'receipt.jpg');
        var imgOpts = {
          method:  'post',
          headers: headers,
          payload: { 'receipt[original]': blob },
          muteHttpExceptions: true
        };
        UrlFetchApp.fetch('https://secure.splitwise.com/api/v3.0/update_expense/'+expId, imgOpts);
      } catch(imgErr) {
        Logger.log('Receipt upload warning: ' + imgErr);
        // not fatal — expense was created
      }
    }

    return { success:true, expenseId: expId };

  } catch(err) {
    Logger.log('createSplitwiseExpense exception: ' + err);
    return { success:false, error: err.message || String(err) };
  }
}



// ═══════════════════════════════════════════════════════════════
// ADD EXPENSE DIALOG  (Trigger: Monthly Expences · J1)
// ═══════════════════════════════════════════════════════════════

function showAddExpDialog() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ADD_EXP')
      .setWidth(2000)
      .setHeight(2000);
    SpreadsheetApp.getUi().showModalDialog(html, '💸 Add Expense');
  } catch (err) { /* silent fail on mobile */ }
}

// Returns raw ADD_EXP.html content so render.html can embed it in an iframe
function getAddExpHtmlContent() {
  return HtmlService.createHtmlOutputFromFile('ADD_EXP').getContent();
}

// ── Helper: column letter(s) → 1-based column number ───────────
function _colToNum(col) {
  col = col.toUpperCase();
  var n = 0;
  for (var i = 0; i < col.length; i++) {
    n = n * 26 + (col.charCodeAt(i) - 64);
  }
  return n;
}

// ── Helper: normalise a raw cell value to string ────────────────
// dateMode=true → formats Date as "M-d" (no leading zeros) to match col A keys like "4-4"
// IMPORTANT: must use the *spreadsheet* timezone so Date objects match what the
// user sees in the sheet (e.g. "4-4" for April 4 Colombo time).
// Using Session.getScriptTimeZone() causes an off-by-one if the script tz ≠ sheet tz.
function _cellStr(v, dateMode) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date && !isNaN(v)) {
    if (dateMode) {
      var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
      return Utilities.formatDate(v, tz, 'M-d');
    }
    return ''; // non-date-mode: skip Date cells (they're not expense values)
  }
  return String(v);
}

/* ── getExpenseRow ──────────────────────────────────────────────
   dateKey: "M-D" string matching column A, e.g. "4-4" for Apr 4
   Returns an object keyed by column letter with raw cell values,
   plus _row (1-based row number) and _sheet (sheet name).
   Also accepts the previous day's key as prevKey for TZ fallback.
─────────────────────────────────────────────────────────────── */
function getExpenseRow(dateKey, prevKey) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Monthly Expences');
    if (!sheet) return { error: 'Sheet "Monthly Expences" not found' };

    var lastRow = sheet.getLastRow();
    // Read only column A to find the matching row
    var colA = sheet.getRange(1, 1, lastRow, 1).getValues();

    var targetRow  = -1;
    var fallbackRow = -1;
    var isPrevDay   = false;

    for (var i = 0; i < colA.length; i++) {
      // Pass dateMode=true so Date objects are formatted as "M-d"
      var cell = _cellStr(colA[i][0], true).trim();
      if (cell === dateKey)  { targetRow  = i + 1; break; }
      if (prevKey && cell === prevKey) { fallbackRow = i + 1; }
    }

    if (targetRow < 0) {
      if (fallbackRow > 0) {
        targetRow = fallbackRow;
        isPrevDay = true;
      } else {
        return { error: 'Row not found for date: ' + dateKey };
      }
    }

    // Read both values (computed results) and formulas for the row
    var numCols  = 37; // A through AK
    var rowRange = sheet.getRange(targetRow, 1, 1, numCols);
    var rowVals  = rowRange.getValues()[0];
    var rowFmls  = rowRange.getFormulas()[0];  // "" if cell has no formula

    var COLS = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S',
                'T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK'];

    var result = { _row: targetRow, _sheet: 'Monthly Expences', _prevDay: isPrevDay };
    for (var c = 0; c < COLS.length && c < rowVals.length; c++) {
      var fml = rowFmls[c];  // e.g. "=800+350" or ""
      var val = rowVals[c];  // e.g. 1150 (computed)
      if (fml && fml.trim() !== '') {
        // Preserve the formula string so the input box shows "=800+350"
        result[COLS[c]] = fml;
      } else {
        // Use computed value; format Date in col A as "M-d"
        result[COLS[c]] = _cellStr(val, COLS[c] === 'A');
      }
    }
    return result;

  } catch (err) {
    return { error: err.message || String(err) };
  }
}

/* ── updateExpenseCells ─────────────────────────────────────────
   payload: { dateKey, prevKey, cells: { 'C': '=800+350', 'D': 'Kottu', ... } }
   Finds the matching row then updates each specified cell individually.
   Formula strings (starting with =) are written as formulas.
   Empty string → writes empty string (clears the cell).
─────────────────────────────────────────────────────────────── */
function updateExpenseCells(payload) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Monthly Expences');
    if (!sheet) return { success: false, error: 'Sheet "Monthly Expences" not found' };

    // Re-find the row (safe against stale _row if sheet was edited)
    var rowInfo = getExpenseRow(payload.dateKey, payload.prevKey);
    if (rowInfo.error) return { success: false, error: rowInfo.error };

    var rowNum = rowInfo._row;
    var cells  = payload.cells || {};
    var updated = [];

    Object.keys(cells).forEach(function(col) {
      var colNum = _colToNum(col);
      var val    = cells[col];
      var cell   = sheet.getRange(rowNum, colNum);

      if (val === '' || val === null || val === undefined) {
        cell.clearContent();
      } else if (String(val).trim().charAt(0) === '=') {
        // Write as a formula — Google Sheets will evaluate it
        cell.setFormula(String(val).trim());
      } else {
        // Try to preserve numbers as numbers (not strings)
        var num = parseFloat(String(val).replace(/,/g,''));
        if (!isNaN(num) && String(val).trim() === String(num)) {
          cell.setValue(num);
        } else {
          cell.setValue(val);
        }
      }
      updated.push(col);
    });

    SpreadsheetApp.flush();
    return { success: true, updatedCols: updated, row: rowNum };

  } catch (err) {
    Logger.log('updateExpenseCells error: ' + err);
    return { success: false, error: err.message || String(err) };
  }
}

function getBankHolidays(year) {
  if (!year) year = new Date().getFullYear();
  var url = 'https://www.cbsl.gov.lk/en/about/about-the-bank/bank-holidays-' + year;
  try {
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) return [];
    var html = response.getContentText();
    var holidays = [];
    
    var rowRegex = /<tr>\s*<td[^>]*>(.*?)<\/td>\s*<td[^>]*>(.*?)<\/td>\s*<\/tr>/gi;
    var match;
    while ((match = rowRegex.exec(html)) !== null) {
      var dateStr = match[1].replace(/&nbsp;/g, ' ').replace(/<[^>]+>/g, '').trim();
      var descStr = match[2].replace(/&nbsp;/g, ' ').replace(/<[^>]+>/g, '').replace(/[^ -~]/g, ' ').replace(/\s+/g, ' ').trim();
      if (dateStr && descStr && dateStr.length > 3 && descStr.length > 3) {
        if (dateStr.toLowerCase().indexOf('date') === -1 && descStr.toLowerCase().indexOf('holiday') === -1) {
           if (dateStr.match(/^[a-zA-Z]+\s+\d{1,2}/)) {
             holidays.push({ date: dateStr, desc: descStr });
           }
        }
      }
    }
    return holidays;
  } catch (e) {
    Logger.log(e);
    return [];
  }
}

/**
 * Get subscriptions data from OTHERS sheet
 * Columns M to T (13-20), starting from row 2
 * Headers in row 2:
 * M: Month (until)
 * N: Platform
 * O: Plan/Type
 * P: Cost
 * Q: Collections
 * R: collection from
 * S: net cost
 * T: Note
 */
function getSubscriptions() {
  try {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName('OTHERS');

    if (!sheet) {
      Logger.log('OTHERS sheet not found');
      return { success: false, error: 'OTHERS sheet not found' };
    }

    // Get data from M3 onwards (skip header row 2)
    var lastRow = sheet.getLastRow();
    Logger.log('Last row: ' + lastRow);

    if (lastRow < 3) {
      Logger.log('No data rows found');
      return { success: true, entries: [], total: 0 };
    }

    // Read M3:T[lastRow] (columns 13-20)
    var data = sheet.getRange(3, 13, lastRow - 2, 8).getValues();
    Logger.log('Data rows: ' + data.length);

    var entries = [];
    var total = 0;

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var month = row[0] ? String(row[0]).trim() : '';
      var platform = row[1] ? String(row[1]).trim() : '';
      var plan = row[2] ? String(row[2]).trim() : '';
      var cost = row[3] ? Number(row[3]) || 0 : 0;
      var collections = row[4] ? Number(row[4]) || 0 : 0;
      var collectionFrom = row[5] ? String(row[5]).trim() : '';
      var netCost = row[6] ? Number(row[6]) || 0 : 0;
      var note = row[7] ? String(row[7]).trim() : '';

      // Skip empty rows or Total row
      if (!platform || platform.toLowerCase() === 'total') {
        if (platform.toLowerCase() === 'total' && netCost > 0) {
          total = netCost;
        }
        continue;
      }

      entries.push({
        month: month,
        platform: platform,
        plan: plan,
        cost: cost,
        collections: collections,
        collectionFrom: collectionFrom,
        netCost: netCost,
        note: note
      });
    }

    Logger.log('Entries found: ' + entries.length);
    Logger.log('Total: ' + total);

    // If no total found in sheet, calculate it
    if (total === 0) {
      for (var j = 0; j < entries.length; j++) {
        total += entries[j].netCost;
      }
    }

    var result = {
      success: true,
      entries: entries,
      total: total
    };

    Logger.log('Returning result with ' + entries.length + ' entries');
    return result;

  } catch (e) {
    Logger.log('Error: ' + e.message);
    return { success: false, error: e.message };
  }
}

