/**
 * =========================================================================
 *                   GOOGLE APPS SCRIPT FOR PATIENT DATABASE
 * =========================================================================
 * 
 * INSTRUCTIONS FOR DEPLOYMENT:
 * 1. Open your Google Sheet in your web browser.
 * 2. Click "Extensions" from the top menu -> Select "Apps Script".
 * 3. Delete any default code in the editor, copy and paste ALL of this code below.
 * 4. Click the Save icon (floppy disk).
 * 5. Click "Deploy" (top right) -> Select "Manage deployments".
 * 6. Click the gear icon next to "Select type" -> Select "Web app".
 * 7. Configure deployment:
 *    - Description: Patient Database API (Batch & Rectangular Optimized)
 *    - Execute as: "Me" (your email)
 *    - Who has access: "Anyone" (essential so the Flask server can communicate)
 * 8. Click "Deploy". You will be asked to authorize permissions.
 * 9. Copy the "Web app URL" that is generated.
 * 10. Paste this URL into your Flask .env file under GOOGLE_APPS_SCRIPT_URL=...
 */

var TABLES = [
  "Demographics",
  "DMP",
  "RMD",
  "FamilyPersonalHistory",
  "Pathology",
  "Followup",
  "Samples",
  "Analysis"
];

function normalizePatientId(pid) {
  var clean = String(pid).trim();
  if (clean.indexOf(".") !== -1) {
    var num = parseFloat(clean);
    if (!isNaN(num) && num === Math.floor(num)) {
      return String(Math.floor(num));
    }
  }
  return clean;
}

function doGet(e) {
  var action = e.parameter.action;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (action === "getJoinedData") {
    var filters = {};
    for (var key in e.parameter) {
      if (key !== "action") {
        filters[key] = e.parameter[key].split(",");
      }
    }
    try {
      var data = getJoinedDataImpl(ss, filters);
      return ContentService.createTextOutput(JSON.stringify({ success: true, data: data }))
                           .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.message }))
                           .setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  if (action === "getFilters") {
    var filterCols = {
      "FamilyPersonalHistory": ["Family_history_of_Ca_Breast", "Multiple_neoplasia"],
      "Pathology": ["Type_of_tumor", "Grade", "Lympho_infiltrate", "Desmoplasia", "Adj_breast", "DCIS"],
      "RMD": ["Menopausal_status"],
      "DMP": ["Clinical_staging_simplified"]
    };
    var result = {};
    for (var tab in filterCols) {
      var sheet = ss.getSheetByName(tab);
      if (!sheet) continue;
      var data = sheet.getDataRange().getValues();
      if (data.length <= 1) continue;
      var headers = data[0];
      
      var cols = filterCols[tab];
      cols.forEach(function(col) {
        var idx = headers.indexOf(col);
        if (idx !== -1) {
          var vals = [];
          for (var r = 1; r < data.length; r++) {
            var val = String(data[r][idx]).trim();
            if (val && vals.indexOf(val) === -1) {
              vals.push(val);
            }
          }
          result[col] = vals.sort();
        } else {
          result[col] = [];
        }
      });
    }
    return ContentService.createTextOutput(JSON.stringify({ success: true, filters: result }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === "getHeaders") {
    var result = {};
    TABLES.forEach(function(tab) {
      var sheet = ss.getSheetByName(tab);
      if (sheet) {
        var data = sheet.getDataRange().getValues();
        result[tab] = data.length > 0 ? data[0] : [];
      } else {
        result[tab] = [];
      }
    });
    return ContentService.createTextOutput(JSON.stringify({ success: true, headers: result }))
                         .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "getAuthorizedUsers") {
    var sheet = ss.getSheetByName("AuthorizedUsers");
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({ success: true, users: [] }))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    var data = sheet.getDataRange().getValues();
    var users = [];
    if (data.length > 1) {
      var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
      var userIdx = headers.indexOf("username");
      var passIdx = headers.indexOf("password");
      if (userIdx !== -1 && passIdx !== -1) {
        for (var r = 1; r < data.length; r++) {
          var u = String(data[r][userIdx]).trim();
          var p = String(data[r][passIdx]).trim();
          if (u && p) {
            users.push({ username: u, password: p });
          }
        }
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ success: true, users: users }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ success: false, message: "Invalid action" }))
                       .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (action === "bulk_upload") {
      var tablesData = payload.tables;
      var results = {};
      var warnings = [];
      var newColumnsAll = {};
      
      // Auto demographics index checking
      var demoPids = {};
      var demoSheet = ss.getSheetByName("Demographics");
      if (demoSheet) {
        var demoVals = demoSheet.getDataRange().getValues();
        if (demoVals.length > 0) {
          var dPidIdx = demoVals[0].indexOf("Patient_ID");
          if (dPidIdx !== -1) {
            for (var d = 1; d < demoVals.length; d++) {
              var dp = normalizePatientId(demoVals[d][dPidIdx]);
              if (dp) demoPids[dp] = true;
            }
          }
        }
      }
      
      var newDemos = [];
      
      // Process each table tab batch
      for (var tabName in tablesData) {
        var tabData = tablesData[tabName];
        var excelHeaders = tabData.headers;
        var rows = tabData.rows;
        
        var sheet = ss.getSheetByName(tabName);
        if (!sheet) {
          sheet = ss.insertSheet(tabName);
          sheet.appendRow(excelHeaders);
        }
        
        var dataRange = sheet.getDataRange();
        var values = dataRange.getValues();
        var sheetHeaders = values.length > 0 ? values[0] : [];
        
        var newCols = [];
        excelHeaders.forEach(function(h) {
          if (h && sheetHeaders.indexOf(h) === -1) {
            newCols.push(h);
          }
        });
        
        if (newCols.length > 0) {
          newCols.forEach(function(col) {
            sheetHeaders.push(col);
          });
          sheet.getRange(1, 1, 1, sheetHeaders.length).setValues([sheetHeaders]);
          values = sheet.getDataRange().getValues();
          newColumnsAll[tabName] = newCols;
        }
        
        var pidIdx = sheetHeaders.indexOf("Patient_ID");
        if (pidIdx === -1) {
          warnings.push("Sheet " + tabName + " is missing 'Patient_ID'.");
          continue;
        }
        
        var pidToRowIdx = {};
        for (var r = 1; r < values.length; r++) {
          var pid = normalizePatientId(values[r][pidIdx]);
          if (pid) {
            pidToRowIdx[pid] = r;
          }
        }
        
        var inserted = 0;
        var updated = 0;
        var skipped = 0;
        
        rows.forEach(function(row) {
          var pid = normalizePatientId(row["Patient_ID"]);
          if (!pid) {
            skipped++;
            return;
          }
          
          if (tabName !== "Demographics" && demoSheet && !demoPids[pid]) {
            demoPids[pid] = true;
            newDemos.push(pid);
          }
          
          if (pidToRowIdx[pid] !== undefined) {
            var rowIdx = pidToRowIdx[pid];
            var existingRow = values[rowIdx];
            
            // Expand existing row to new headers length if needed
            while (existingRow.length < sheetHeaders.length) {
              existingRow.push("");
            }
            
            excelHeaders.forEach(function(h) {
              var sIdx = sheetHeaders.indexOf(h);
              if (sIdx !== -1) {
                existingRow[sIdx] = row[h] !== undefined ? String(row[h]) : "";
              }
            });
            updated++;
          } else {
            var rowValues = [];
            sheetHeaders.forEach(function(h) {
              rowValues.push(row[h] !== undefined ? String(row[h]) : "");
            });
            values.push(rowValues);
            inserted++;
          }
        });
        
        // CRITICAL PERFORMANCE FIX: Pad all rows to make a perfect rectangular 2D array!
        for (var i = 0; i < values.length; i++) {
          while (values[i].length < sheetHeaders.length) {
            values[i].push("");
          }
          for (var j = 0; j < values[i].length; j++) {
            if (values[i][j] === undefined || values[i][j] === null) {
              values[i][j] = "";
            }
          }
        }
        
        sheet.getRange(1, 1, values.length, sheetHeaders.length).setValues(values);
        results[tabName] = { inserted: inserted, updated: updated, skipped: skipped };
      }
      
      // Batch write new Patient IDs to Demographics sheet if needed
      if (newDemos.length > 0 && demoSheet) {
        var demoVals = demoSheet.getDataRange().getValues();
        var demoHeaders = demoVals.length > 0 ? demoVals[0] : ["Patient_ID"];
        var dIdx = demoHeaders.indexOf("Patient_ID");
        if (dIdx !== -1) {
          newDemos.forEach(function(p) {
            var r = [];
            demoHeaders.forEach(function(h) {
              r.push(h === "Patient_ID" ? p : "");
            });
            demoVals.push(r);
          });
          
          for (var i = 0; i < demoVals.length; i++) {
            while (demoVals[i].length < demoHeaders.length) {
              demoVals[i].push("");
            }
          }
          
          demoSheet.getRange(1, 1, demoVals.length, demoHeaders.length).setValues(demoVals);
          warnings.push("Auto-created " + newDemos.length + " new Patient ID(s) inside 'Demographics' tab.");
        }
      }
      
      return ContentService.createTextOutput(JSON.stringify({ 
        success: true, 
        message: "Bulk multi-category upload completed successfully.",
        results: results,
        new_columns: newColumnsAll,
        warnings: warnings
      })).setMimeType(ContentService.MimeType.JSON);
    }
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: "Apps Script error: " + err.message }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

function getJoinedDataImpl(ss, filters) {
  var patients = {};
  
  TABLES.forEach(function(tab) {
    var sheet = ss.getSheetByName(tab);
    if (!sheet) return;
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;
    var headers = data[0];
    
    var pidIdx = headers.indexOf("Patient_ID");
    if (pidIdx === -1) return;
    
    for (var r = 1; r < data.length; r++) {
      var rawPid = String(data[r][pidIdx]).trim();
      var pid = normalizePatientId(rawPid);
      if (!pid) continue;
      
      if (!patients[pid]) {
        patients[pid] = { "Patient_ID": pid };
      }
      
      for (var c = 0; c < headers.length; c++) {
        if (c !== pidIdx && headers[c]) {
          patients[pid][headers[c]] = String(data[r][c]).trim();
        }
      }
    }
  });
  
  var result = [];
  for (var pid in patients) {
    var p = patients[pid];
    var match = true;
    for (var col in filters) {
      if (filters[col].length > 0 && filters[col][0] !== "") {
        var val = String(p[col] || "").trim();
        if (filters[col].indexOf(val) === -1) {
          match = false;
          break;
        }
      }
    }
    if (match) {
      result.push(p);
    }
  }
  
  result.sort(function(a, b) {
    var ai = parseInt(a.Patient_ID);
    var bi = parseInt(b.Patient_ID);
    if (isNaN(ai)) return 1;
    if (isNaN(bi)) return -1;
    return ai - bi;
  });
  
  return result;
}
