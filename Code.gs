/**
 * @fileoverview Main server-side script for the Training Management System.
 * Handles Google Sheets menu functions, web app backend logic, and onEdit triggers.
 */

// --- CONFIGURATION ---
const SPREADSHEET_ID = '13smCfWTRiFPF5GE21gExRbBbOJ8Lipa9EQPJDyW9IMo'; // Replace with your Spreadsheet ID
const SHEET_NAMES = {
  RECORDS: 'Training_Records',
  EMPLOYEES: 'Employee_Database',
  REFERENCE: 'Reference_Data',
  ATTENDANCE: 'Training Attendance',
  ITR: 'ITR'
};

// --- HELPER FUNCTION ---
function toSimpleKey(str) {
  if (!str || typeof str !== 'string') return '';
  return str.toLowerCase().replace(/\s+/g, '').replace(/\W/g, '');
}

// --- GOOGLE SHEETS UI FUNCTIONS ---
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Training Tools')
    .addItem('Generate Attendance Sheet', 'generateAttendance')
    .addItem('Generate ITR (Individual Training Record)', 'generateITR')
    .addToUi();
}


function generateAttendance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attendanceSheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE);
  if (!attendanceSheet) {
    SpreadsheetApp.getUi().alert(`"${SHEET_NAMES.ATTENDANCE}" sheet not found.`);
    return;
  }

  const course = attendanceSheet.getRange("D6").getValue().toString().trim();
  const dateStartValue = attendanceSheet.getRange("D7").getValue();
  const dateEndValue = attendanceSheet.getRange("D8").getValue();

  const clearAllOutputs = () => {
    attendanceSheet.getRange("H6:H8").clearContent();
    attendanceSheet.getRange(11, 3, attendanceSheet.getMaxRows() - 10, 5).clearContent();
  };

  if (!course || !dateStartValue || !dateEndValue) {
    clearAllOutputs();
    SpreadsheetApp.getUi().alert('Please fill in Course, Start Date, and End Date (D6:D8) before generating.');
    return;
  }

  const dateStart = new Date(dateStartValue);
  const dateEnd = new Date(dateEndValue);

  const employeeMap = getEmployeeMap(ss);
  const sourceSheet = ss.getSheetByName(SHEET_NAMES.RECORDS);
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert(`"${SHEET_NAMES.RECORDS}" sheet not found.`);
    return;
  }

  const data = sourceSheet.getDataRange().getValues();
  const headers = data.shift().map(h => String(h).trim());
  const headerMap = new Map(headers.map((h, i) => [h, i]));

  // Get indices for our columns
  const empCodeIdx = headerMap.get('EMPLOYEE CODE');
  const empNameIdx = headerMap.get('EMPLOYEE NAME');
  const plantIdx = headerMap.get('PLANT');
  const courseIdx = headerMap.get('TRAINING COURSE');
  const dateStartedIdx = headerMap.get('DATE STARTED');
  const dateEndedIdx = headerMap.get('DATE ENDED');
  const companyIdx = headerMap.get('COMPANY NAME');
  const positionIdx = headerMap.get('POSITION');


  const filteredRecords = data.filter(row => {
    const recordCourse = row[courseIdx];
    return recordCourse && recordCourse.toString().trim() === course &&
      isSameDate(new Date(row[dateStartedIdx]), dateStart) &&
      isSameDate(new Date(row[dateEndedIdx]), dateEnd);
  });

  if (filteredRecords.length > 0) {
    const firstRecord = filteredRecords[0];
    attendanceSheet.getRange("H6").setValue(firstRecord[headerMap.get('TRAINING VENUE')]);
    attendanceSheet.getRange("H7").setValue(firstRecord[headerMap.get('TIME STARTED')]);
    attendanceSheet.getRange("H8").setValue(firstRecord[headerMap.get('TIME ENDED')]);

    const output = filteredRecords.map(row => {
      const empCode = row[empCodeIdx] ? row[empCodeIdx].toString() : null;
      const empDetails = employeeMap[empCode] || { position: '', company: '' };

      const company = (companyIdx !== undefined && row[companyIdx]) ? row[companyIdx] : empDetails.company;
      const position = (positionIdx !== undefined && row[positionIdx]) ? row[positionIdx] : empDetails.position;

      return [
        empCode,
        row[empNameIdx],
        company,
        row[plantIdx],
        position
      ];
    });

    attendanceSheet.getRange(11, 3, attendanceSheet.getMaxRows() - 10, 5).clearContent();
    attendanceSheet.getRange(11, 3, output.length, 5).setValues(output);
    SpreadsheetApp.getUi().alert(`Successfully generated attendance for ${output.length} participants.`);
  } else {
    clearAllOutputs();
    SpreadsheetApp.getUi().alert('No matching records found for the specified criteria.');
  }
}

/**
 * --- FINAL UPDATED FUNCTION ---
 * Generates an Individual Training Record.
 * Includes a fallback to search 'Training_Records' if an employee code is not found
 * in the main 'Employee_Database', and populates all available fields (including FA/BU and DEPT).
 */
function generateITR() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const itrSheet = ss.getSheetByName(SHEET_NAMES.ITR);
  if (!itrSheet) {
    SpreadsheetApp.getUi().alert(`"${SHEET_NAMES.ITR}" sheet not found.`);
    return;
  }
  const empCode = itrSheet.getRange("I7").getValue();
  if (!empCode) {
    SpreadsheetApp.getUi().alert('Please enter an Employee Code in cell I7.');
    return;
  }
  // Clear previous data
  itrSheet.getRangeList(["I8:I9", "P7:P9"]).clearContent();
  if (itrSheet.getMaxRows() > 12) {
    itrSheet.getRange("D13:S" + itrSheet.getMaxRows()).clearContent();
  }

  let employeeFound = false;

  // Step 1: Try to find employee in the Employee Database
  const empSheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);
  if (empSheet) {
    const empData = empSheet.getDataRange().getValues();
    const empHeaders = empData.shift();
    const employeeRow = empData.find(row => row[empHeaders.indexOf('EMPLOYEE CODE')] == empCode);
    if (employeeRow) {
      itrSheet.getRange("I8").setValue(employeeRow[empHeaders.indexOf('EMPLOYEE NAME')]);
      itrSheet.getRange("I9").setValue(employeeRow[empHeaders.indexOf('FA/BU')]);
      itrSheet.getRange("P7").setValue(employeeRow[empHeaders.indexOf('DEPT')]);
      itrSheet.getRange("P8").setValue(employeeRow[empHeaders.indexOf('POSITION')]);
      itrSheet.getRange("P9").setValue(employeeRow[empHeaders.indexOf('WORK LOCATION')]);
      employeeFound = true;
    }
  }

  // Step 2: If not in database, search Training_Records as a fallback
  const dbSheet = ss.getSheetByName(SHEET_NAMES.RECORDS);
  if (!employeeFound && dbSheet) {
    const dbData = dbSheet.getDataRange().getValues();
    const headers = dbData.shift().map(h => String(h).trim().toUpperCase());
    const codeIdx = headers.indexOf('EMPLOYEE CODE');
    const nameIdx = headers.indexOf('EMPLOYEE NAME');
    const positionIdx = headers.indexOf('POSITION');
    const plantIdx = headers.indexOf('PLANT');
    const fabuIdx = headers.indexOf('FA/BU');
    const deptIdx = headers.indexOf('DEPT');
    
    const recordRow = dbData.find(row => row[codeIdx] == empCode);
    
    if (recordRow) {
      itrSheet.getRange("I8").setValue(recordRow[nameIdx]);
      itrSheet.getRange("I9").setValue(recordRow[fabuIdx]);
      itrSheet.getRange("P7").setValue(recordRow[deptIdx]);
      itrSheet.getRange("P8").setValue(recordRow[positionIdx]);
      itrSheet.getRange("P9").setValue(recordRow[plantIdx]);
      employeeFound = true;
      SpreadsheetApp.getUi().alert('Note: This is a non-database employee. Displaying data from training history.');
    }
  }
  
  if (!employeeFound) {
    SpreadsheetApp.getUi().alert('Employee Code not found in Employee Database or Training Records.');
    return;
  }

  // The rest of the function remains the same
  const dbData = dbSheet.getDataRange().getValues();
  const headers = dbData.shift().map(h => String(h).trim().toUpperCase());
  const trainingCodeIdx = headers.indexOf('EMPLOYEE CODE');
  const dateStartedIdx = headers.indexOf('DATE STARTED');
  const completionStatusIdx = headers.indexOf('COMPLETION STATUS');

  if (completionStatusIdx === -1) {
    SpreadsheetApp.getUi().alert('"COMPLETION STATUS" column not found in Training_Records sheet.');
    return;
  }

  const trainingRecords = dbData
    .filter(row => {
      if (row[trainingCodeIdx] != empCode) return false;
      const completionStatus = String(row[completionStatusIdx]).trim().toLowerCase();
      const isCompleted = (completionStatus === 'completed' || completionStatus === 'completed (no assessment)');
      return isCompleted;
    })
    .sort((a, b) => new Date(b[dateStartedIdx]) - new Date(a[dateStartedIdx]));

  if (trainingRecords.length > 0) {
    const outputData = trainingRecords.map(row => [
      row[headers.indexOf('TRAINING COURSE')], null, null, null, null, null, null, null, null,
      row[headers.indexOf('DATE STARTED')],
      row[headers.indexOf('DATE ENDED')],
      row[headers.indexOf('TOTAL HRS')],
      null,
      row[headers.indexOf('TRAINING FACILITATOR')],
      row[headers.indexOf('TRAINING CATEGORY')],
      row[headers.indexOf('TRAINING STATUS')]
    ]);
    itrSheet.getRange(13, 4, outputData.length, outputData[0].length).setValues(outputData);
    SpreadsheetApp.getUi().alert(`Found and sorted ${trainingRecords.length} completed training records.`);
  } else {
    SpreadsheetApp.getUi().alert(`No completed training records found for employee ${empCode}.`);
  }
}

function onEdit(e) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  if (sheetName === "Corporate Training") {
    const editedRange = e.range;
    const EMPLOYEE_DATABASE_SHEET_NAME = "Employee Database";
    const NAME_COLUMN_INDEX_CT = 3;
    const PLANT_COLUMN_INDEX_CT = 4;
    const EMPLOYEE_CODE_COLUMN_INDEX_CT = 2;
    const DATABASE_START_ROW = 3;
    const DATABASE_CODE_COLUMN_INDEX_DB = 4;
    const DATABASE_FULL_NAME_COLUMN_INDEX_DB = 5;
    const DATABASE_MIDDLE_NAME_COLUMN_INDEX_DB = 9;
    const DATABASE_WORK_LOCATION_COLUMN_INDEX_DB = 12;
    const LAST_NAME_MATCH_THRESHOLD_STRICT = 0.85;
    const LAST_NAME_MATCH_THRESHOLD_NORMAL = 0.7;
    const FIRST_NAME_MIN_LENGTH_FOR_PARTIAL = 3;
    const PARTIAL_MATCH_THRESHOLD_LOW = 0.6;
    const CLEAR_CODE_ON_NO_MATCH = true;
    const LOCATION_MATCH_BONUS = 0.5;
    const STRONG_NAME_MATCH_BONUS = 0.7;
    const MIDDLE_NAME_MATCH_BONUS = 0.2;
    const startRow = editedRange.getRow();
    const numRows = editedRange.getNumRows();
    const startCol = editedRange.getColumn();
    const numCols = editedRange.getNumColumns();

    if (
      (startCol > NAME_COLUMN_INDEX_CT || (startCol + numCols - 1) < NAME_COLUMN_INDEX_CT) &&
      (startCol > PLANT_COLUMN_INDEX_CT || (startCol + numCols - 1) < PLANT_COLUMN_INDEX_CT)
    ) return;

    const dbSheet = e.source.getSheetByName(EMPLOYEE_DATABASE_SHEET_NAME);
    if (!dbSheet) return;

    const dbLastRow = dbSheet.getLastRow();
    if (dbLastRow < DATABASE_START_ROW) return;

    const dbData = dbSheet.getRange(
      DATABASE_START_ROW,
      DATABASE_CODE_COLUMN_INDEX_DB,
      dbLastRow - DATABASE_START_ROW + 1,
      9
    ).getDisplayValues();

    for (let i = 0; i < numRows; i++) {
      const row = startRow + i;
      const nameCell = sheet.getRange(row, NAME_COLUMN_INDEX_CT);
      const plantCell = sheet.getRange(row, PLANT_COLUMN_INDEX_CT);
      const empCodeCell = sheet.getRange(row, EMPLOYEE_CODE_COLUMN_INDEX_CT);
      const inputName = nameCell.getValue();
      const inputPlant = onEdit_normalizeNameString(plantCell.getValue());

      if (!inputName) {
        empCodeCell.clearContent();
        continue;
      }

      const inputParts = onEdit_extractNameParts(inputName);
      let bestMatch = { code: "", score: 0 };

      for (let j = 0; j < dbData.length; j++) {
        const dbRow = dbData[j];
        const code = dbRow[DATABASE_CODE_COLUMN_INDEX_DB - DATABASE_CODE_COLUMN_INDEX_DB];
        const dbFullName = dbRow[DATABASE_FULL_NAME_COLUMN_INDEX_DB - DATABASE_CODE_COLUMN_INDEX_DB];
        const dbMiddleName = dbRow[DATABASE_MIDDLE_NAME_COLUMN_INDEX_DB - DATABASE_CODE_COLUMN_INDEX_DB];
        const dbWorkLocation = dbRow[DATABASE_WORK_LOCATION_COLUMN_INDEX_DB - DATABASE_CODE_COLUMN_INDEX_DB];
        if (!code) continue;

        const dbParts = onEdit_extractNameParts(dbFullName);
        const normalizedDbWorkLocation = onEdit_normalizeNameString(dbWorkLocation);
        const normalizedDbMiddleName = onEdit_normalizeNameString(dbMiddleName);
        let currentScore = 0;
        let locationMatch = false;

        if (inputPlant && normalizedDbWorkLocation) {
          if (inputPlant.includes(normalizedDbWorkLocation) || normalizedDbWorkLocation.includes(inputPlant)) {
            locationMatch = true;
          }
        } else if (inputPlant === normalizedDbWorkLocation) {
          locationMatch = true;
        }

        const lastNameMatchScore = onEdit_getMatchScore(inputParts.last, dbParts.last);
        const firstNamePartialMatchScore = onEdit_getPartialMatchScore(inputParts.first, dbParts.first);
        const middleNamePartialMatchScore = onEdit_getPartialMatchScore(inputParts.middle, normalizedDbMiddleName);
        const fullNamePartialMatchScore = onEdit_getPartialMatchScore(inputParts.full, dbParts.full);
        const firstNamePresent = inputParts.first.length >= FIRST_NAME_MIN_LENGTH_FOR_PARTIAL;
        const dbFirstNamePresent = dbParts.first.length >= FIRST_NAME_MIN_LENGTH_FOR_PARTIAL;

        if (lastNameMatchScore >= LAST_NAME_MATCH_THRESHOLD_STRICT && firstNamePresent && dbFirstNamePresent && firstNamePartialMatchScore >= PARTIAL_MATCH_THRESHOLD_LOW) {
          currentScore = lastNameMatchScore + firstNamePartialMatchScore + STRONG_NAME_MATCH_BONUS;
          if (locationMatch) currentScore += LOCATION_MATCH_BONUS;
          if (normalizedDbMiddleName && inputParts.middle && middleNamePartialMatchScore >= PARTIAL_MATCH_THRESHOLD_LOW) {
            currentScore += MIDDLE_NAME_MATCH_BONUS;
          } else if (!normalizedDbMiddleName && !inputParts.middle) {
            currentScore += MIDDLE_NAME_MATCH_BONUS / 2;
          }
        } else if (locationMatch) {
          if (lastNameMatchScore >= LAST_NAME_MATCH_THRESHOLD_NORMAL && firstNamePartialMatchScore >= PARTIAL_MATCH_THRESHOLD_LOW) {
            currentScore = lastNameMatchScore + firstNamePartialMatchScore + LOCATION_MATCH_BONUS;
            if (middleNamePartialMatchScore >= PARTIAL_MATCH_THRESHOLD_LOW) {
              currentScore += MIDDLE_NAME_MATCH_BONUS;
            }
          } else if (fullNamePartialMatchScore >= PARTIAL_MATCH_THRESHOLD_LOW + 0.1) {
            currentScore = fullNamePartialMatchScore + LOCATION_MATCH_BONUS;
          }
        } else if (inputPlant === "") {
          if (lastNameMatchScore >= LAST_NAME_MATCH_THRESHOLD_NORMAL && firstNamePartialMatchScore >= PARTIAL_MATCH_THRESHOLD_LOW) {
            currentScore = lastNameMatchScore + firstNamePartialMatchScore;
            if (middleNamePartialMatchScore >= PARTIAL_MATCH_THRESHOLD_LOW) {
              currentScore += MIDDLE_NAME_MATCH_BONUS;
            }
          } else if (fullNamePartialMatchScore >= PARTIAL_MATCH_THRESHOLD_LOW) {
            currentScore = fullNamePartialMatchScore;
          }
        }

        if (currentScore > bestMatch.score) {
          bestMatch = { code, score: currentScore };
        }
      }

      if (bestMatch.code) {
        empCodeCell.setValue(bestMatch.code);
      } else if (CLEAR_CODE_ON_NO_MATCH) {
        empCodeCell.clearContent();
      }
    }
  }

  if (sheetName === "Training Attendance") {
    const editedRange = e.range;
    const sheet = e.source.getSheetByName("Training Attendance");

    const editedA1Notation = editedRange.getA1Notation();
    if (editedA1Notation !== "D6" && editedA1Notation !== "D7" && editedA1Notation !== "D8") {
        return;
    }

    const course = sheet.getRange("D6").getValue().toString().trim();
    const dateStartValue = sheet.getRange("D7").getValue();
    const dateEndValue = sheet.getRange("D8").getValue();

    const clearAllOutputs = () => {
      sheet.getRange("H6:H8").clearContent();
      const outputRange = sheet.getRange(11, 3, sheet.getMaxRows() - 10, 5);
      outputRange.clearContent();
    };

    if (!course || !dateStartValue || !dateEndValue) {
      clearAllOutputs();
      return;
    }

    const dateStart = new Date(dateStartValue);
    const dateEnd = new Date(dateEndValue);

    if (isNaN(dateStart.getTime()) || isNaN(dateEnd.getTime())) return;

    const sourceSheet = e.source.getSheetByName("Corporate Training");
    const data = sourceSheet.getDataRange().getValues();
    let matchFound = false;
    let output = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const trainingCourse = row[5];
      const startDate = new Date(row[11]);
      const endDate = new Date(row[12]);

      if (
        trainingCourse.toString().trim() === course &&
        isSameDate(startDate, dateStart) &&
        isSameDate(endDate, dateEnd)
      ) {
        if (!matchFound) {
          sheet.getRange("H6").setValue(row[13]);
          sheet.getRange("H7").setValue(row[15]);
          sheet.getRange("H8").setValue(row[16]);
          matchFound = true;
        }

        const empCode = row[1];
        const fullName = row[2];
        const plant = row[3];
        const companyName = row[32];
        const position = row[33];
        output.push([empCode, fullName, companyName, plant, position]);
      }
    }

    const startRow = 11;
    const outputRange = sheet.getRange(startRow, 3, sheet.getMaxRows() - 10, 5);
    outputRange.clearContent();

    if (output.length > 0) {
      sheet.getRange(startRow, 3, output.length, 5).setValues(output);
    }
  }
}


// --- WEB APP ---
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Training Management System Database')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getReferenceData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.REFERENCE);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const referenceData = {};
  headers.forEach((header, index) => {
    const uniqueValues = [...new Set(data.map(row => row[index]).filter(String))];
    referenceData[header] = uniqueValues;
  });
  return referenceData;
}

function searchEmployees() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const h = new Map(headers.map((col, i) => [col.trim(), i]));
  return data
    .filter(row => row[h.get('ACTIVE')] && row[h.get('ACTIVE')].toString().trim().toUpperCase() === 'YES')
    .map(row => ({
      employeeCode: row[h.get('EMPLOYEE CODE')],
      employeeName: row[h.get('EMPLOYEE NAME')],
      workLocation: row[h.get('WORK LOCATION')],
      active: row[h.get('ACTIVE')],
      companyName: row[h.get('COMPANY NAME')],
      plant: row[h.get('WORK LOCATION')],
      dept: row[h.get('DEPT')],
      position: row[h.get('POSITION')],
      fabu: row[h.get('FA/BU')] // <-- ADD THIS LINE
    }));
}

function getDashboardData(dateRange) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.RECORDS);
  if (!sheet || sheet.getLastRow() < 2) {
    return { totalTrainings: 0, totalHours: 0, totalParticipants: 0, totalCost: 0, participantsByCategory: {}, trainingsByConduct: {}, upcomingTrainings: [], hoursByDept: {} };
  }
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const h = new Map(headers.map((col, i) => [col.trim(), i]));
  const filteredData = filterDataByDateRange(data, h.get('DATE STARTED'), dateRange);
  const today = new Date();
  const thirtyDaysFromNow = new Date();
  thirtyDaysFromNow.setDate(today.getDate() + 30);
  const dashboard = { totalHours: 0, totalCost: 0, participantsByCategory: {}, uniqueTrainings: new Set(), trainingsByConduct: {}, upcomingTrainings: [] };
  filteredData.forEach(row => {
    dashboard.totalHours += parseFloat(row[h.get('TOTAL HRS')]) || 0;
    dashboard.totalCost += parseFloat(row[h.get('COST')]) || 0;
    const trainingId = `${row[h.get('TRAINING COURSE')]}|${row[h.get('DATE STARTED')]}|${row[h.get('TRAINING VENUE')]}`;
    if (!dashboard.uniqueTrainings.has(trainingId)) {
      dashboard.uniqueTrainings.add(trainingId);
      const conduct = row[h.get('TYPE OF CONDUCT')] || 'N/A';
      dashboard.trainingsByConduct[conduct] = (dashboard.trainingsByConduct[conduct] || 0) + 1;
    }
    const category = row[h.get('TRAINING CATEGORY')] || 'Uncategorized';
    dashboard.participantsByCategory[category] = (dashboard.participantsByCategory[category] || 0) + 1;
    const trainingDate = new Date(row[h.get('DATE STARTED')]);
    if (trainingDate >= today && trainingDate <= thirtyDaysFromNow) {
      dashboard.upcomingTrainings.push({
        course: row[h.get('TRAINING COURSE')],
        date: Utilities.formatDate(trainingDate, Session.getScriptTimeZone(), 'MM/dd/yyyy'),
        venue: row[h.get('TRAINING VENUE')]
      });
    }
  });
  const hoursByDept = getTrainingHoursByDepartment(filteredData, headers);
  return {
    totalTrainings: dashboard.uniqueTrainings.size, totalHours: dashboard.totalHours.toFixed(2), totalParticipants: filteredData.length, totalCost: dashboard.totalCost.toFixed(2),
    participantsByCategory: dashboard.participantsByCategory, trainingsByConduct: dashboard.trainingsByConduct, upcomingTrainings: dashboard.upcomingTrainings, hoursByDept
  };
}

/**
 * --- UPDATED FUNCTION ---
 * Saves new training records, now converting Employee Code and other key fields to uppercase.
 */
function saveTrainingRecords(records) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.RECORDS);
  if (!sheet) throw new Error('Training_Records sheet not found.');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastId = sheet.getRange(sheet.getLastRow(), 1).getValue();
  let startNumber = (typeof lastId === 'number' && lastId > 0) ? lastId + 1 : 1;

  const participantCount = records.length; // <-- ADD THIS LINE

  const dataToAppend = records.map((record, index) => {
    if (new Date(record.dateended) < new Date(record.datestarted)) {
      throw new Error('Date Ended cannot be before Date Started.');
    }

  // --- ADD THESE LINES to calculate cost per participant ---
    let costPerParticipant = 0;
    const totalCost = parseFloat(record.cost);
    if (!isNaN(totalCost) && totalCost > 0 && participantCount > 0) {
      costPerParticipant = totalCost / participantCount;
    }
  // --- END of new lines ---

    const preScore = parseFloat(record.pretestscore);
    const postScore = parseFloat(record.posttestscore);
    const preTotal = parseFloat(record.pretesttotalitems);
    const postTotal = parseFloat(record.posttesttotalitems);

    const { finalRatingValue, finalRatingRemark } = calculateFinalRating(preScore, postScore, preTotal, postTotal);

    const newRecord = {
      ...record,
      cost: costPerParticipant, // <-- UPDATE THIS LINE
      finalratingremark: finalRatingRemark,
      no: startNumber + index
    };

    const numericRating = parseFloat(finalRatingValue);
    if (!isNaN(numericRating)) {
      newRecord.finalrating = numericRating / 100;
    } else {
      newRecord.finalrating = '';
    }

    // Convert specific text fields to uppercase for consistency
    const fieldsToUppercase = ['employeecode', 'employeename', 'plant', 'companyname', 'position', 'fabu', 'dept'];
    fieldsToUppercase.forEach(field => {
      if (newRecord[field] && typeof newRecord[field] === 'string') {
        newRecord[field] = newRecord[field].toUpperCase();
      }
    });

    return headers.map(header => {
      const fieldName = toSimpleKey(header);
      const value = newRecord[fieldName];
      if (value instanceof Date) {
        return Utilities.formatDate(value, Session.getScriptTimeZone(), 'MM/dd/yyyy');
      }
      return value ?? '';
    });
  });

  if (dataToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, dataToAppend.length, dataToAppend[0].length).setValues(dataToAppend);
  }
  return { status: 'success', message: `${dataToAppend.length} records saved successfully.` };
}

function getTrainingSessionDetails(sessionDetails) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.RECORDS);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const spreadsheetTimezone = ss.getSpreadsheetTimeZone();
  const allData = sheet.getDataRange().getValues();
  const headers = allData.shift().map(h => String(h).trim());
  const headerMap = new Map(headers.map((h, i) => [toSimpleKey(h), i]));

  const dateStartedIdx = headerMap.get('datestarted');
  const courseIdx = headerMap.get('trainingcourse');

  if (dateStartedIdx === undefined || courseIdx === undefined) {
    Logger.log("Error: 'Date Started' or 'Training Course' columns not found.");
    return [];
  }

  const clientDateStr = String(sessionDetails.datestarted).trim();

  const filteredRows = allData.filter(row => {
    const courseName = String(row[courseIdx]).trim();
    const dateValue = row[dateStartedIdx];
    if (!(dateValue instanceof Date)) return false;
    const recordDateStr = Utilities.formatDate(dateValue, spreadsheetTimezone, 'MM/dd/yyyy');
    return courseName === String(sessionDetails.trainingcourse).trim() && recordDateStr === clientDateStr;
  });

  const records = filteredRows.map(row => {
    const record = {};
    headers.forEach((h, j) => {
      const key = toSimpleKey(h);
      let value = row[j];

      if (value instanceof Date) {
        if (key.includes('time')) {
          record[key] = Utilities.formatDate(value, spreadsheetTimezone, 'HH:mm');
        } else {
          record[key] = Utilities.formatDate(value, spreadsheetTimezone, 'MM/dd/yyyy');
        }
      } else {
        record[key] = value;
      }
    });
    return record;
  });

  return records;
}

/**
 * --- UPDATED FUNCTION ---
 * Updates existing training records, now converting Employee Code and other key fields to uppercase.
 */
function updateTrainingRecord(updatedRecord, finalParticipantList) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.RECORDS);
  if (!sheet) throw new Error('Training_Records sheet not found.');

  const spreadsheetTimezone = ss.getSpreadsheetTimeZone();
  const allData = sheet.getDataRange().getValues();
  const headers = allData.shift();
  const headerMap = new Map(headers.map((h, i) => [toSimpleKey(String(h)), i]));

  const trainingId = toSimpleKey(updatedRecord.trainingcourse) + toSimpleKey(updatedRecord.datestarted);
  const dateStartedIdx = headerMap.get('datestarted');
  const courseIdx = headerMap.get('trainingcourse');

  const rowsToDelete = [];
  allData.forEach((row, index) => {
    if (row[dateStartedIdx] instanceof Date) {
      const rowDate = Utilities.formatDate(new Date(row[dateStartedIdx]), spreadsheetTimezone, 'MM/dd/yyyy');
      const rowTrainingId = toSimpleKey(String(row[courseIdx]).trim()) + toSimpleKey(rowDate);
      if (rowTrainingId === trainingId) {
        rowsToDelete.push(index + 2);
      }
    }
  });

  rowsToDelete.sort((a, b) => b - a).forEach(rowNum => {
    sheet.deleteRow(rowNum);
  });

  const lastId = sheet.getLastRow() > 1 ? sheet.getRange(sheet.getLastRow(), 1).getValue() : 0;
  let startNumber = (typeof lastId === 'number' && lastId > 0) ? lastId + 1 : 1;

// --- ADD THIS NEW BLOCK to calculate cost per participant ---
    const participantCount = finalParticipantList.length;
    let costPerParticipant = 0;
    const totalCost = parseFloat(updatedRecord.cost);
    if (!isNaN(totalCost) && totalCost > 0 && participantCount > 0) {
      costPerParticipant = totalCost / participantCount;
    }
    // --- END of new block ---

  const dataToAppend = finalParticipantList.map((participant, index) => {
    const fullRecord = { ...participant, ...updatedRecord, cost: costPerParticipant, no: startNumber + index };
    
    // Convert specific text fields to uppercase for consistency
    const fieldsToUppercase = ['employeecode', 'employeename', 'plant', 'companyname', 'position', 'fabu', 'dept'];
    fieldsToUppercase.forEach(field => {
      if (fullRecord[field] && typeof fullRecord[field] === 'string') {
        fullRecord[field] = fullRecord[field].toUpperCase();
      }
    });
    
    const preScore = parseFloat(fullRecord.pretestscore);
    const postScore = parseFloat(fullRecord.posttestscore);
    const preTotal = parseFloat(fullRecord.pretesttotalitems);
    const postTotal = parseFloat(fullRecord.posttesttotalitems);

    const canRecalculate = !isNaN(preScore) && !isNaN(postScore) && preTotal > 0 && postTotal > 0;

    if (canRecalculate) {
        const { finalRatingValue, finalRatingRemark } = calculateFinalRating(preScore, postScore, preTotal, postTotal);
        fullRecord.finalratingremark = finalRatingRemark;
        const numericRating = parseFloat(finalRatingValue);
        if (!isNaN(numericRating)) {
            fullRecord.finalrating = numericRating / 100;
        } else {
            fullRecord.finalrating = '';
        }
    }

    return headers.map(header => {
      const key = toSimpleKey(String(header));
      const value = fullRecord[key];
      return value ?? '';
    });
  });

  if (dataToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, dataToAppend.length, dataToAppend[0].length).setValues(dataToAppend);
  }

  return { status: 'success', message: 'Training record updated successfully.' };
}


function deleteTrainingRecord(record) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.RECORDS);
  if (!sheet) throw new Error('Training_Records sheet not found.');

  const allData = sheet.getDataRange().getValues();
  const headers = allData[0].map(h => String(h).trim());
  const headerMap = new Map(headers.map((h, i) => [toSimpleKey(h), i]));

  const trainingId = toSimpleKey(record.trainingcourse) + toSimpleKey(record.datestarted);

  let rowsToDelete = [];
  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    const rowDate = Utilities.formatDate(new Date(row[headerMap.get('datestarted')]), Session.getScriptTimeZone(), 'MM/dd/yyyy');
    const rowTrainingId = toSimpleKey(String(row[headerMap.get('trainingcourse')]).trim()) + toSimpleKey(rowDate);
    if (rowTrainingId === trainingId) {
      rowsToDelete.push(i + 1);
    }
  }

  if (rowsToDelete.length > 0) {
    rowsToDelete.sort((a, b) => b - a);
    rowsToDelete.forEach(rowNum => sheet.deleteRow(rowNum));
    return { status: 'success', message: `${rowsToDelete.length} records for the training session have been deleted.` };
  } else {
    return { status: 'error', message: 'Training session not found.' };
  }
}

function getIndividualTrainingRecord(searchTerm) {
  if (!searchTerm || searchTerm.trim() === '') {
    return { employee: null, records: [] };
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const employeeSheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);
  const recordsSheet = ss.getSheetByName(SHEET_NAMES.RECORDS);

  if (!employeeSheet || !recordsSheet) {
    throw new Error('Required sheets (Employee Database or Training Records) not found.');
  }

  // 1. Find the employee
  const employeeData = employeeSheet.getDataRange().getValues();
  const empHeaders = employeeData.shift().map(h => String(h).trim().toUpperCase());
  const empCodeIdx = empHeaders.indexOf('EMPLOYEE CODE');
  const empNameIdx = empHeaders.indexOf('EMPLOYEE NAME');
  
  const normalizedSearch = searchTerm.trim().toLowerCase();
  let foundEmployeeRow = null;

  for (const row of employeeData) {
    const code = row[empCodeIdx] ? String(row[empCodeIdx]).trim().toLowerCase() : '';
    const name = row[empNameIdx] ? String(row[empNameIdx]).trim().toLowerCase() : '';
    if (code === normalizedSearch || name.includes(normalizedSearch)) {
      foundEmployeeRow = row;
      break;
    }
  }

  if (!foundEmployeeRow) {
    return { employee: null, records: [] };
  }

  // 2. Get Employee Details
  const empDeptIdx = empHeaders.indexOf('DEPT');
  const empPosIdx = empHeaders.indexOf('POSITION');
  const empFabuIdx = empHeaders.indexOf('FA/BU');
  const empLocIdx = empHeaders.indexOf('WORK LOCATION');

  const employeeDetails = {
    employeeNumber: foundEmployeeRow[empCodeIdx],
    employeeName: foundEmployeeRow[empNameIdx],
    department: foundEmployeeRow[empDeptIdx],
    position: foundEmployeeRow[empPosIdx],
    fabu: foundEmployeeRow[empFabuIdx],
    location: foundEmployeeRow[empLocIdx]
  };
  
  const employeeCode = employeeDetails.employeeNumber;

  // 3. Get Training History
  const recordsData = recordsSheet.getDataRange().getValues();
  const recHeaders = recordsData.shift().map(h => String(h).trim().toUpperCase());
  const recCodeIdx = recHeaders.indexOf('EMPLOYEE CODE');
  const recCourseIdx = recHeaders.indexOf('TRAINING COURSE');
  const recStartIdx = recHeaders.indexOf('DATE STARTED');
  const recEndIdx = recHeaders.indexOf('DATE ENDED');
  const recHoursIdx = recHeaders.indexOf('TOTAL HRS');
  const recFacilitatorIdx = recHeaders.indexOf('TRAINING FACILITATOR');
  const recCategoryIdx = recHeaders.indexOf('TRAINING CATEGORY');
  const recCompletionStatusIdx = recHeaders.indexOf('COMPLETION STATUS');

  const trainingHistory = recordsData
    .filter(row => {
      if (row[recCodeIdx] != employeeCode) return false;
      const completionStatus = String(row[recCompletionStatusIdx]).trim().toLowerCase();
      return completionStatus === 'completed' || completionStatus === 'completed (no assessment)';
    })
    .map((row, index) => {
      const startDate = row[recStartIdx] instanceof Date ? Utilities.formatDate(row[recStartIdx], ss.getSpreadsheetTimeZone(), 'MM/dd/yyyy') : '';
      const endDate = row[recEndIdx] instanceof Date ? Utilities.formatDate(row[recEndIdx], ss.getSpreadsheetTimeZone(), 'MM/dd/yyyy') : '';
      return {
        no: index + 1,
        courseTitle: row[recCourseIdx],
        trainingStarted: startDate,
        trainingEnded: endDate,
        trainingDuration: `${row[recHoursIdx]} hours`,
        facilitator: row[recFacilitatorIdx],
        trainingCategory: row[recCategoryIdx],
        remarks: row[recCompletionStatusIdx]
      };
    });

  return { employee: employeeDetails, records: trainingHistory };
}


// --- NEW FUNCTIONS FOR RECORDS MANAGEMENT ---
/**
 * Searches for unique training course names from the records sheet.
 * @param {string} query - The search string to filter course names.
 * @returns {string[]} A list of unique, filtered training course names.
 */
function searchTrainingSessions(query) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.RECORDS);
    if (!sheet || sheet.getLastRow() < 2) return [];

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const courseIndex = headers.findIndex(h => h.trim().toUpperCase() === 'TRAINING COURSE');

    if (courseIndex === -1) {
      Logger.log('"TRAINING COURSE" column not found in Training_Records.');
      return [];
    }

    const uniqueCourses = [...new Set(data.map(row => row[courseIndex]).filter(String))];
    const normalizedQuery = query ? query.trim().toLowerCase() : '';

    if (!normalizedQuery) {
      return uniqueCourses.sort();
    }

    return uniqueCourses.filter(course =>
      course.toLowerCase().includes(normalizedQuery)
    ).sort();

  } catch (e) {
    Logger.log(`Error in searchTrainingSessions: ${e.toString()}`);
    return [];
  }
}

/**
 * Gets all unique start dates for a given training course.
 * @param {string} courseName - The name of the training course.
 * @returns {string[]} A list of unique, formatted dates for that course.
 */
function getTrainingDatesForCourse(courseName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.RECORDS);
    if (!sheet || sheet.getLastRow() < 2 || !courseName) return [];

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const courseIndex = headers.findIndex(h => h.trim().toUpperCase() === 'TRAINING COURSE');
    const dateIndex = headers.findIndex(h => h.trim().toUpperCase() === 'DATE STARTED');

    if (courseIndex === -1 || dateIndex === -1) {
      Logger.log('Required columns (TRAINING COURSE or DATE STARTED) not found.');
      return [];
    }

    const dates = new Set();
    const spreadsheetTimezone = ss.getSpreadsheetTimeZone();

    data.forEach(row => {
      if (row[courseIndex] && row[courseIndex].toString().trim() === courseName.trim() && row[dateIndex] instanceof Date) {
        const formattedDate = Utilities.formatDate(row[dateIndex], spreadsheetTimezone, 'MM/dd/yyyy');
        dates.add(formattedDate);
      }
    });

    return Array.from(dates).sort((a, b) => new Date(b) - new Date(a));

  } catch (e) {
    Logger.log(`Error in getTrainingDatesForCourse: ${e.toString()}`);
    return [];
  }
}

function getParticipantsList(courseName, trainingDate) {
  if (!courseName || !trainingDate) {
    throw new Error('Training course and date are required.');
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const recordsSheet = ss.getSheetByName(SHEET_NAMES.RECORDS);
  const employeeSheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);

  if (!recordsSheet || !employeeSheet) {
    throw new Error('Required sheets not found.');
  }

  const employeeData = employeeSheet.getDataRange().getValues();
  const empHeaders = employeeData.shift().map(h => String(h).trim().toUpperCase());
  const empCodeIdx = empHeaders.indexOf('EMPLOYEE CODE');
  const empLevelIdx = empHeaders.indexOf('LEVEL');
  const empStatusIdx = empHeaders.indexOf('ACTIVE');
  const empWorkLocationIdx = empHeaders.indexOf('WORK LOCATION');
  const empFabuIdx = empHeaders.indexOf('FA/BU');
  const empPositionIdx = empHeaders.indexOf('POSITION');
  const empCompanyIdx = empHeaders.indexOf('COMPANY NAME');

  const employeeMap = new Map();
  employeeData.forEach(row => {
    const code = row[empCodeIdx];
    if (code) {
      employeeMap.set(String(code).trim(), {
        level: empLevelIdx > -1 ? row[empLevelIdx] : 'N/A',
        status: row[empStatusIdx],
        workLocation: row[empWorkLocationIdx],
        fabu: row[empFabuIdx],
        position: row[empPositionIdx],
        companyName: row[empCompanyIdx]
      });
    }
  });

  const recordsData = recordsSheet.getDataRange().getValues();
  const recHeaders = recordsData.shift();
  const recCourseIdx = recHeaders.indexOf('TRAINING COURSE');
  const recDateIdx = recHeaders.indexOf('DATE STARTED');
  const recEmpCodeIdx = recHeaders.indexOf('EMPLOYEE CODE');
  const recEmpNameIdx = recHeaders.indexOf('EMPLOYEE NAME');
  const recWorkLocationIdx = recHeaders.indexOf('PLANT');
  const recFabuIdx = recHeaders.indexOf('FA/BU');
  const recPositionIdx = recHeaders.indexOf('POSITION');
  const recCompanyIdx = recHeaders.indexOf('COMPANY NAME');

  const spreadsheetTimezone = ss.getSpreadsheetTimeZone();
  const formattedTrainingDate = Utilities.formatDate(new Date(trainingDate), spreadsheetTimezone, 'MM/dd/yyyy');

  const participants = recordsData
    .filter(row => {
      const rowCourse = String(row[recCourseIdx]).trim();
      if (!row[recDateIdx] || !(row[recDateIdx] instanceof Date)) return false;
      const rowDate = Utilities.formatDate(new Date(row[recDateIdx]), spreadsheetTimezone, 'MM/dd/yyyy');
      return rowCourse === courseName && rowDate === formattedTrainingDate;
    })
    .map(row => {
      const empCode = String(row[recEmpCodeIdx]).trim();
      const employeeDetails = employeeMap.get(empCode);

      const recordData = {
        workLocation: row[recWorkLocationIdx],
        fabu: row[recFabuIdx],
        position: row[recPositionIdx],
        companyName: row[recCompanyIdx],
        level: 'N/A',
        status: 'N/A'
      };

      const finalData = employeeDetails ? employeeDetails : recordData;

      return {
        employeeNumber: empCode,
        employeeName: row[recEmpNameIdx],
        workLocation: finalData.workLocation,
        fabu: finalData.fabu,
        position: finalData.position,
        level: finalData.level,
        companyName: finalData.companyName,
        status: String(finalData.status).trim().toUpperCase() === 'YES' ? 'Active' : 'Inactive'
      };
    });

  return participants;
}

// --- HELPER & UTILITY FUNCTIONS ---
function getTrainingHoursByDepartment(data, headers) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const employeeMap = getEmployeeMap(ss);
  const empCodeTrIdx = headers.indexOf('EMPLOYEE CODE');
  const totalHrsIdx = headers.indexOf('TOTAL HRS');
  const hoursByDept = {};
  data.forEach(row => {
    const empCode = row[empCodeTrIdx];
    const dept = employeeMap[empCode] ? employeeMap[empCode].dept : 'N/A';
    const hours = parseFloat(row[totalHrsIdx]) || 0;
    hoursByDept[dept] = (hoursByDept[dept] || 0) + hours;
  });
  return hoursByDept;
}

function getEmployeeMap(ss) {
  const spreadsheet = ss || SpreadsheetApp.openById(SPREADSHEET_ID);
  const empSheet = spreadsheet.getSheetByName(SHEET_NAMES.EMPLOYEES);
  if (!empSheet) return {};
  const empData = empSheet.getDataRange().getValues();
  const empHeaders = empData.shift();
  const h = new Map(empHeaders.map((col, i) => [col.trim(), i]));
  const employeeMap = {};
  empData.forEach(row => {
    const code = row[h.get('EMPLOYEE CODE')] ? row[h.get('EMPLOYEE CODE').toString()] : null;
    if (code) {
      employeeMap[code] = {
        employeeName: row[h.get('EMPLOYEE NAME')],
        position: row[h.get('POSITION')],
        company: row[h.get('COMPANY NAME')],
        dept: row[h.get('DEPT')],
        division: row[h.get('FA/BU')],
        active: row[h.get('ACTIVE')],
        plant: row[h.get('WORK LOCATION')]
      };
    }
  });
  return employeeMap;
}

function calculateFinalRating(preScore, postScore, preTotal, postTotal) {
  let finalRatingValue = '';
  let finalRatingRemark = '';
  if (!isNaN(preScore) && !isNaN(postScore) && preTotal > 0 && postTotal > 0) {
    const prePercentage = preScore / preTotal;
    const postPercentage = postScore / postTotal;
    if (prePercentage === 0 && postPercentage > 0) {
      finalRatingValue = (postPercentage * 100).toFixed(2);
      finalRatingRemark = 'SCORE';
    } else if (prePercentage > 0 && postPercentage > prePercentage) {
      const improvement = ((postPercentage - prePercentage) / prePercentage) * 100;
      finalRatingValue = improvement.toFixed(2);
      finalRatingRemark = 'IMPROVEMENT';
    } else if (postPercentage === prePercentage) {
      finalRatingValue = '0.00';
      finalRatingRemark = 'NO CHANGE';
    } else if (prePercentage > postPercentage) {
      const decrease = ((prePercentage - postPercentage) / prePercentage) * 100;
      finalRatingValue = `-${decrease.toFixed(2)}`;
      finalRatingRemark = 'DECREASE';
    }
  }
  return { finalRatingValue, finalRatingRemark };
}

function isSameDate(d1, d2) {
  return d1 instanceof Date && d2 instanceof Date &&
    d1.getFullYear() === d2.getFullYear() &&
    d1.getMonth() === d2.getMonth() &&
    d1.getDate() === d2.getDate();
}

function filterDataByDateRange(data, dateIndex, range) {
  if (range === 'all') return data;
  const now = new Date();
  let startDate = new Date();
  startDate.setHours(0, 0, 0, 0);
  if (range === 'quarter') {
    const quarter = Math.floor(now.getMonth() / 3);
    startDate.setMonth(quarter * 3, 1);
  } else if (range === 'year') {
    startDate.setMonth(0, 1);
  }
  return data.filter(row => {
    const rowDate = new Date(row[dateIndex]);
    return rowDate >= startDate && rowDate <= now;
  });
}

// THE FIX: Renamed all helper functions for the onEdit trigger to be unique.
function onEdit_extractNameParts(name) {
  const normalized = onEdit_normalizeNameString(name);
  const commaSplit = normalized.split(",");
  let last = "";
  let first = "";
  let middle = "";
  if (commaSplit.length >= 2) {
    last = commaSplit[0].trim();
    const rest = commaSplit[1].trim().split(" ");
    first = rest[0] || "";
    middle = rest.length > 1 ? rest.slice(1).join(" ") : "";
  } else {
    const parts = normalized.split(" ");
    first = parts[0] || "";
    middle = parts.length > 2 ? parts.slice(1, -1).join(" ") : "";
    last = parts[parts.length - 1] || "";
  }
  return {
    first,
    middle,
    last,
    full: normalized.replace(/,/g, "")
  };
}

function onEdit_normalizeNameString(text) {
  if (!text) return "";
  return text.toString().toLowerCase().replace(/[^a-z\s,]/g, "").replace(/\s+/g, " ").trim();
}

function onEdit_getMatchScore(str1, str2) {
  if (!str1 || !str2) return 0;
  const minLength = Math.min(str1.length, str2.length);
  let score = 0;
  for (let i = 0; i < minLength; i++) {
    if (str1[i] === str2[i]) score++;
  }
  return minLength > 0 ? score / minLength : 0;
}

function onEdit_getPartialMatchScore(str1, str2) {
  if (!str1 || !str2) return 0;
  const words1 = str1.split(" ").filter(Boolean);
  const words2 = str2.split(" ").filter(Boolean);
  if (words1.length === 0) return 1;
  let matches = 0;
  for (const word1 of words1) {
    if (words2.includes(word1)) matches++;
  }
  return matches / words1.length;
}
