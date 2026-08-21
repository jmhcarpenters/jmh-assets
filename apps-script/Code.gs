function doGet(e) {
  const iconUrl = "https://raw.githubusercontent.com/jmhcarpenters/jmh-assets/main/jmh-icon.png";
  const isManager = e && e.parameter && String(e.parameter.page || "").toLowerCase() === "manager";
  return HtmlService.createHtmlOutputFromFile(isManager ? "Admin" : "Index")
    .setTitle(isManager ? "JMH Manager Portal" : "JMH Crew Time Log")
    .setFaviconUrl(iconUrl)
    .addMetaTag("viewport", "width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover");
}


function getJobs() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Jobs");

  if (!sheet) {
    throw new Error(
      'The "Jobs" sheet was not found.'
    );
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const jobs = sheet
    .getRange(2, 2, lastRow - 1, 1)
    .getValues()
    .flat()
    .map(function (job) {
      return String(job || "").trim();
    })
    .filter(Boolean);

  return [...new Set(jobs)];
}


function getCategories() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Tasks");

  if (!sheet) {
    throw new Error('The "Tasks" sheet was not found.');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const categories = sheet
    .getRange(2, 1, lastRow - 1, 1)
    .getDisplayValues()
    .flat()
    .map(function (category) { return String(category || "").trim(); })
    .filter(Boolean);

  return [...new Set(categories)];
}


function getPhases(category) {
  const sheet = getJmhSpreadsheet_().getSheetByName("Tasks");
  if (!sheet) throw new Error('The "Tasks" sheet was not found.');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const selectedCategory = String(category || "").trim().toLowerCase();
  if (!selectedCategory) return [];

  const phases = sheet
    .getRange(2, 1, lastRow - 1, 2)
    .getDisplayValues()
    .filter(function (row) {
      return String(row[0] || "").trim().toLowerCase() === selectedCategory;
    })
    .map(function (row) { return String(row[1] || "").trim(); })
    .filter(Boolean);

  return [...new Set(phases)];
}


function submitTimeAction(data) {
  if (!data) throw new Error("No information was received.");

  const action = String(data.action || "").trim();
  const employee = String(data.employee || "").trim();
  const job = String(data.job || "").trim();
  const category = String(data.category || "").trim();
  const phase = String(data.phase || "").trim();
  const task = String(data.task || "").trim();
  const notes = String(data.notes || "").trim();

  if (!employee) throw new Error("Please select an employee.");

  const passcode = String(data.passcode || "").trim();
  verifyCrewDeviceOrPasscode(employee, passcode, data.deviceToken);

  const spreadsheet = getJmhSpreadsheet_();
  const timeLog = spreadsheet.getSheetByName("Time Log");
  const taskLog = spreadsheet.getSheetByName("Task Log");

  if (!timeLog) throw new Error('The "Time Log" sheet was not found.');
  if (!taskLog) throw new Error('The "Task Log" sheet was not found.');

  const timestamp = data.originalDeviceTime ? new Date(data.originalDeviceTime) : new Date();
  timestamp.setMilliseconds(0);

  if (action === "clockIn") {
    if (!job || !category || !phase || !task) {
      throw new Error("Please select a job and starting task before clocking in.");
    }
    const clockInTimestamp = prepareClockInTimestamp(spreadsheet, job, timestamp);
    clockIn(timeLog, employee, clockInTimestamp, job, notes);
    startTask(timeLog, taskLog, employee, clockInTimestamp, job, phase, task, notes);
    rmStampTimeLogMetadata(2, data);
    updateTimeLogCalculation(2);
    return employee + " clocked in successfully.";
  }

  if (action === "startTask") {
    if (!job || !category || !phase || !task) {
      throw new Error("Please select the category, phase, and task.");
    }
    startTask(timeLog, taskLog, employee, timestamp, job, phase, task, notes);
    return employee + " started: " + task;
  }

  if (action === "startChangeOrder") {
    const changeOrderId = String(data.changeOrderId || "").trim();
    if (!job || !changeOrderId) throw new Error("Please select a change order.");
    startChangeOrder(timeLog, employee, timestamp, job, changeOrderId, notes);
    return employee + " started Change Order " + changeOrderId;
  }

  if (action === "clockOut") {
    const closingRow = findOpenRow(timeLog, employee, timestamp);
    const roundedClockOut = prepareClockOutTimestamp(timestamp);
    clockOut(timeLog, taskLog, employee, roundedClockOut);
    if (closingRow) updateTimeLogCalculation(closingRow);
    return employee + " clocked out successfully.";
  }

  throw new Error("The selected action was not recognized.");
}

function prepareClockInTimestamp(spreadsheet, job, timestamp) {
  const policy = getJobClockInPolicy(spreadsheet, job);
  enforceEarliestClockIn(timestamp, policy);
  const interval = 15 * 60 * 1000;
  const rounded = new Date(Math.round(timestamp.getTime() / interval) * interval);
  rounded.setSeconds(0, 0);
  return rounded;
}

function prepareClockOutTimestamp(timestamp) {
  const interval = 15 * 60 * 1000;
  const rounded = new Date(Math.round(timestamp.getTime() / interval) * interval);
  rounded.setSeconds(0, 0);
  return rounded;
}

function getJobClockInPolicy(spreadsheet, job) {
  const sheet = spreadsheet.getSheetByName("Jobs");
  if (!sheet) throw new Error('The "Jobs" sheet was not found.');
  const lastColumn = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(String);
  const col = function (name) { return headers.indexOf(name); };
  const jobCol = col("Job Name");
  const addressCol = col("Location");
  const earliestCol = col("Earliest Clock In");
  const latCol = col("Latitude");
  const lngCol = col("Longitude");
  const radiusCol = col("GPS Radius (ft)");
  if ([jobCol, addressCol, earliestCol, latCol, lngCol, radiusCol].some(function (i) { return i < 0; })) {
    throw new Error("The Jobs sheet is missing clock-in policy columns.");
  }
  const count = Math.max(lastRow - 1, 1);
  const values = sheet.getRange(2, 1, count, lastColumn).getValues();
  const display = sheet.getRange(2, 1, count, lastColumn).getDisplayValues();
  const selected = String(job || "").trim();
  let row = -1;
  for (let i = 0; i < display.length; i++) {
    if (String(display[i][jobCol] || "").trim() === selected) { row = i; break; }
  }
  if (row < 0) throw new Error("The selected job could not be found.");
  let latitude = Number(values[row][latCol]);
  let longitude = Number(values[row][lngCol]);
  const address = String(display[row][addressCol] || "").trim();
  if ((!Number.isFinite(latitude) || !Number.isFinite(longitude) || !latitude || !longitude) && address) {
    const result = Maps.newGeocoder().geocode(address);
    if (result && result.status === "OK" && result.results && result.results.length) {
      latitude = Number(result.results[0].geometry.location.lat);
      longitude = Number(result.results[0].geometry.location.lng);
      sheet.getRange(row + 2, latCol + 1).setValue(latitude);
      sheet.getRange(row + 2, lngCol + 1).setValue(longitude);
    }
  }
  return {
    job: selected,
    earliest: values[row][earliestCol],
    earliestDisplay: String(display[row][earliestCol] || "").trim(),
    latitude: latitude,
    longitude: longitude,
    radiusFeet: Number(values[row][radiusCol]) || 500
  };
}

function enforceEarliestClockIn(timestamp, policy) {
  if (!policy.earliestDisplay) return;
  let hours, minutes;
  if (policy.earliest instanceof Date) {
    hours = policy.earliest.getHours();
    minutes = policy.earliest.getMinutes();
  } else {
    const match = policy.earliestDisplay.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)?$/i);
    if (!match) throw new Error("Invalid earliest clock-in time for " + policy.job + ".");
    hours = Number(match[1]); minutes = Number(match[2]);
    const meridiem = String(match[3] || "").toUpperCase();
    if (meridiem === "PM" && hours < 12) hours += 12;
    if (meridiem === "AM" && hours === 12) hours = 0;
  }
  const opening = new Date(timestamp);
  opening.setHours(hours, minutes, 0, 0);
  if (timestamp < opening) {
    throw new Error("Clock-in for " + policy.job + " opens at " + policy.earliestDisplay + ".\nLa entrada para " + policy.job + " comienza a las " + policy.earliestDisplay + ".");
  }
}

function enforceJobsiteLocation(latitude, longitude, policy) {
  if (!Number.isFinite(policy.latitude) || !Number.isFinite(policy.longitude) || !policy.latitude || !policy.longitude) {
    throw new Error("GPS coordinates are not configured for " + policy.job + ".");
  }
  if (!Number.isFinite(latitude) || !Number.isFinite(longitude)) {
    throw new Error("Location permission is required to clock in.\nSe requiere permiso de ubicación para registrar la entrada.");
  }
  const toRad = Math.PI / 180;
  const lat1 = latitude * toRad, lat2 = policy.latitude * toRad;
  const dLat = (policy.latitude - latitude) * toRad;
  const dLng = (policy.longitude - longitude) * toRad;
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(lat1) * Math.cos(lat2) * Math.sin(dLng / 2) ** 2;
  const distanceFeet = 20902231 * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  if (distanceFeet > policy.radiusFeet) {
    throw new Error("You must be at the " + policy.job + " jobsite to clock in. You are about " + Math.round(distanceFeet) + " ft away.\nDebe estar en la obra de " + policy.job + " para registrar su entrada. Está aproximadamente a " + Math.round(distanceFeet) + " pies de distancia.");
  }
}

function clockIn(sheet, employee, timestamp, job, notes) {
  const openRow = findOpenRow(sheet, employee, timestamp);
  if (openRow) throw new Error(employee + " already has an open time entry.");
  appendTimeLogRow(sheet, timestamp, employee, job, notes);
  SpreadsheetApp.flush();
}

function startTask(timeLog, taskLog, employee, timestamp, job, phase, task, notes) {
  const openTimeRow = findOpenRow(timeLog, employee, timestamp);
  if (!openTimeRow) throw new Error(employee + " must clock in before starting a task.");

  const clockedInJob = String(timeLog.getRange(openTimeRow, 3).getDisplayValue() || "").trim();
  if (clockedInJob && clockedInJob.toLowerCase() !== job.toLowerCase()) {
    throw new Error(employee + " is clocked in to " + clockedInJob + ". Please use that job.");
  }
  if (!clockedInJob) timeLog.getRange(openTimeRow, 3).setValue(job);

  const openTaskRow = findOpenTaskRow(taskLog, employee, timestamp);
  if (openTaskRow) setTaskEndTime(taskLog, openTaskRow, timestamp);

  const spreadsheet = getJmhSpreadsheet_();
  const changeOrderLog = spreadsheet.getSheetByName("Change Order Log");
  if (changeOrderLog) {
    const openChangeOrderRow = findOpenChangeOrderRow(changeOrderLog, employee, timestamp);
    if (openChangeOrderRow) setChangeOrderEndTime(changeOrderLog, openChangeOrderRow, timestamp);
  }

  appendTaskLogRow(taskLog, timestamp, employee, job, phase, task, notes);
}

function findOpenChangeOrderRow(changeOrderLog, employee, timestamp) {
  if (!changeOrderLog) return null;
  const lastRow = changeOrderLog.getLastRow();
  if (lastRow < 2) return null;
  const data = changeOrderLog.getRange(2, 1, lastRow - 1, 9).getValues();
  const today = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy-MM-dd");
  for (let index = data.length - 1; index >= 0; index--) {
    const row = data[index];
    const rowEmployee = String(row[1] || "").trim();
    if (!(row[0] instanceof Date)) continue;
    const rowDate = Utilities.formatDate(row[0], Session.getScriptTimeZone(), "yyyy-MM-dd");
    const endTime = row[6];
    const sameEmployee = rowEmployee.toLowerCase() === employee.toLowerCase();
    const isOpen = endTime === "" || endTime === null;
    if (sameEmployee && rowDate === today && isOpen) return index + 2;
  }
  return null;
}

function setChangeOrderEndTime(changeOrderLog, row, timestamp) {
  const timeOnly = createTimeOnly(timestamp);
  changeOrderLog.getRange(row, 7).setValue(timeOnly).setNumberFormat("h:mm AM/PM");
}

function clockOut(timeLog, taskLog, employee, timestamp) {
  const openTimeRow = findOpenRow(timeLog, employee, timestamp);
  if (!openTimeRow) throw new Error("No open time entry was found for " + employee + ".");
  const openTaskRow = findOpenTaskRow(taskLog, employee, timestamp);
  if (openTaskRow) setTaskEndTime(taskLog, openTaskRow, timestamp);
  const spreadsheet = getJmhSpreadsheet_();
  const changeOrderLog = spreadsheet.getSheetByName("Change Order Log");
  if (changeOrderLog) {
    const openChangeOrderRow = findOpenChangeOrderRow(changeOrderLog, employee, timestamp);
    if (openChangeOrderRow) setChangeOrderEndTime(changeOrderLog, openChangeOrderRow, timestamp);
  }
  setEndTime(timeLog, openTimeRow, timestamp);
}

function appendTimeLogRow(sheet, timestamp, employee, job, notes) {
  const dateOnly = new Date(timestamp.getFullYear(), timestamp.getMonth(), timestamp.getDate());
  const timeOnly = new Date(1899, 11, 30, timestamp.getHours(), timestamp.getMinutes(), timestamp.getSeconds());
  if (shouldInsertWeeklySeparator(sheet, timestamp)) sheet.insertRowsBefore(2, 1);
  sheet.insertRowsBefore(2, 1);
  sheet.getRange(2, 1, 1, 6).setValues([[dateOnly, employee, job, timeOnly, "", ""]]);
  sheet.getRange(2, 6).setFormula('=IF(OR(D2="",E2=""),"",IF(MOD(E2-D2,1)*24>=6,1,0))');
  sheet.getRange(2, 7).setValue(0);
  sheet.getRange(2, 8).setFormula('=IF(OR(D2="",E2=""),"",ROUND(MOD(E2-D2,1)*24+N(G2),2))');
  sheet.getRange(2, 9).setFormula('=IF(H2="","",MAX(0,H2-F2))');
  sheet.getRange(2, 10).setFormula('=IF(H2="","",H2*L2)');
  sheet.getRange(2, 12).setValue(getEmployeeHourlyRateSnapshot_(employee));
  sheet.getRange(2, 19).setFormula('=IF(A2="","",IFS(AND(D2="",E2<>""),"Missing clock-in",AND(A2<TODAY(),D2<>"",E2=""),"Missing clock-out",COUNTIFS($A:$A,A2,$B:$B,B2,$D:$D,D2)>1,"Duplicate entry",AND(H2<>"",H2<8),"Short shift - "&TEXT(H2,"0.00")&" hrs (under 8)",AND(H2<>"",H2>10.5),"Long shift - "&TEXT(H2,"0.00")&" hrs (over 10.5)",AND(N2<>"",O2<>"",ABS(O2-N2)*1440>30),"Delayed sync - "&ROUND(ABS(O2-N2)*1440,0)&" minutes",TRUE,""))');
  var reviewStatusCell = sheet.getRange(2, 20);
  var reviewStatusRule = reviewStatusCell.getDataValidation();
  reviewStatusCell.clearDataValidations();
  reviewStatusCell.setFormula('=IF(S2="","","Needs Review")');
  if (reviewStatusRule) reviewStatusCell.setDataValidation(reviewStatusRule);
  sheet.getRange(2, 11).setValue(notes);
  sheet.getRange(2, 1).setNumberFormat("M/d/yy (dddd)");
  sheet.getRange(2, 4, 1, 2).setNumberFormat("h:mm AM/PM");
  SpreadsheetApp.flush();
}

function getEmployeeHourlyRateSnapshot_(employee) {
  const employees = getJmhSpreadsheet_().getSheetByName("Employees");
  if (!employees || employees.getLastRow() < 2) return "";
  const rows = employees.getRange(2, 1, employees.getLastRow() - 1, 3).getValues();
  const target = String(employee || "").trim().toLowerCase();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0] || "").trim().toLowerCase() === target) {
      return rows[i][2] === "" ? "" : Number(rows[i][2]);
    }
  }
  return "";
}

function findFirstEmptyTimeLogRow(sheet) {
  const firstDataRow = 2;
  const maxRows = sheet.getMaxRows();
  const employeeValues = sheet.getRange(firstDataRow, 2, maxRows - 1, 1).getDisplayValues();
  for (let index = 0; index < employeeValues.length; index++) {
    const employee = String(employeeValues[index][0] || "").trim();
    if (!employee) return index + firstDataRow;
  }
  sheet.insertRowAfter(maxRows);
  return maxRows + 1;
}

function findOpenRow(sheet, employee, timestamp) {
  const lastDataRow = getLastTimeLogDataRow(sheet);
  if (lastDataRow < 2) return null;
  const values = sheet.getRange(2, 1, lastDataRow - 1, 5).getValues();
  const targetDate = new Date(timestamp.getFullYear(), timestamp.getMonth(), timestamp.getDate()).getTime();
  for (let index = values.length - 1; index >= 0; index--) {
    const rowDate = values[index][0];
    const rowEmployee = String(values[index][1] || "").trim();
    const startTime = values[index][3];
    const endTime = values[index][4];
    if (!(rowDate instanceof Date)) continue;
    const normalizedDate = new Date(rowDate.getFullYear(), rowDate.getMonth(), rowDate.getDate()).getTime();
    const sameDate = normalizedDate === targetDate;
    const sameEmployee = rowEmployee.toLowerCase() === employee.toLowerCase();
    const hasStartTime = startTime instanceof Date;
    const isOpen = endTime === "" || endTime === null;
    if (sameDate && sameEmployee && hasStartTime && isOpen) return index + 2;
  }
  return null;
}

function appendTaskLogRow(sheet, timestamp, employee, job, phase, task, notes) {
  const dateOnly = new Date(timestamp.getFullYear(), timestamp.getMonth(), timestamp.getDate());
  const timeOnly = createTimeOnly(timestamp);
  sheet.insertRowsBefore(2, 1);
  sheet.getRange(2, 1, 1, 7).setValues([[dateOnly, employee, job, phase, task, timeOnly, ""]]);
  sheet.getRange(2, 9).setValue(notes);
  sheet.getRange(2, 1).setNumberFormat("M/d/yy (dddd)");
  sheet.getRange(2, 6, 1, 2).setNumberFormat("h:mm AM/PM");
  SpreadsheetApp.flush();
}

function findOpenTaskRow(sheet, employee, timestamp) {
  const lastDataRow = getLastTaskLogDataRow(sheet);
  if (lastDataRow < 2) return null;
  const values = sheet.getRange(2, 1, lastDataRow - 1, 7).getValues();
  const targetDate = new Date(timestamp.getFullYear(), timestamp.getMonth(), timestamp.getDate()).getTime();
  for (let index = values.length - 1; index >= 0; index--) {
    const rowDate = values[index][0];
    const rowEmployee = String(values[index][1] || "").trim();
    const startTime = values[index][5];
    const endTime = values[index][6];
    if (!(rowDate instanceof Date)) continue;
    const normalizedDate = new Date(rowDate.getFullYear(), rowDate.getMonth(), rowDate.getDate()).getTime();
    const sameDate = normalizedDate === targetDate;
    const sameEmployee = rowEmployee.toLowerCase() === employee.toLowerCase();
    const hasStartTime = startTime instanceof Date;
    const isOpen = endTime === "" || endTime === null;
    if (sameDate && sameEmployee && hasStartTime && isOpen) return index + 2;
  }
  return null;
}

function getLastTaskLogDataRow(sheet) {
  const maxRows = sheet.getMaxRows();
  const employeeValues = sheet.getRange(2, 2, maxRows - 1, 1).getDisplayValues();
  for (let index = employeeValues.length - 1; index >= 0; index--) {
    const employee = String(employeeValues[index][0] || "").trim();
    if (employee) return index + 2;
  }
  return 1;
}

function setTaskEndTime(sheet, row, timestamp) {
  const timeOnly = createTimeOnly(timestamp);
  sheet.getRange(row, 7).setValue(timeOnly).setNumberFormat("h:mm AM/PM");
}

function createTimeOnly(timestamp) {
  return new Date(1899, 11, 30, timestamp.getHours(), timestamp.getMinutes(), timestamp.getSeconds());
}

function getLastTimeLogDataRow(sheet) {
  const maxRows = sheet.getMaxRows();
  const employeeValues = sheet.getRange(2, 2, maxRows - 1, 1).getDisplayValues();
  for (let index = employeeValues.length - 1; index >= 0; index--) {
    const employee = String(employeeValues[index][0] || "").trim();
    if (employee) return index + 2;
  }
  return 1;
}

function setEndTime(sheet, row, timestamp) {
  const timeOnly = new Date(1899, 11, 30, timestamp.getHours(), timestamp.getMinutes(), timestamp.getSeconds());
  sheet.getRange(row, 5).setValue(timeOnly).setNumberFormat("h:mm AM/PM");
}

function startChangeOrder(timeLog, employee, timestamp, job, changeOrderId, notes) {
  const spreadsheet = getJmhSpreadsheet_();
  const taskLog = spreadsheet.getSheetByName("Task Log");
  const changeOrderLog = spreadsheet.getSheetByName("Change Order Log");
  const changeOrders = spreadsheet.getSheetByName("Change Orders");
  if (!taskLog) throw new Error('The "Task Log" sheet was not found.');
  if (!changeOrderLog) throw new Error('The "Change Order Log" sheet was not found.');
  if (!changeOrders) throw new Error('The "Change Orders" sheet was not found.');

  const openTimeRow = findOpenRow(timeLog, employee, timestamp);
  if (!openTimeRow) throw new Error(employee + " must clock in before starting a change order.");
  const clockedInJob = String(timeLog.getRange(openTimeRow, 3).getDisplayValue() || "").trim();
  if (clockedInJob && clockedInJob.toLowerCase() !== job.toLowerCase()) {
    throw new Error(employee + " is clocked in to " + clockedInJob + ". Please use that job.");
  }
  if (!clockedInJob) timeLog.getRange(openTimeRow, 3).setValue(job);

  const openTaskRow = findOpenTaskRow(taskLog, employee, timestamp);
  if (openTaskRow) setTaskEndTime(taskLog, openTaskRow, timestamp);
  const openChangeOrderRow = findOpenChangeOrderRow(changeOrderLog, employee, timestamp);
  if (openChangeOrderRow) setChangeOrderEndTime(changeOrderLog, openChangeOrderRow, timestamp);

  const lastChangeOrderRow = changeOrders.getLastRow();
  let description = "";
  if (lastChangeOrderRow >= 2) {
    const changeOrderData = changeOrders.getRange(2, 1, lastChangeOrderRow - 1, 3).getDisplayValues();
    for (let index = 0; index < changeOrderData.length; index++) {
      const rowId = String(changeOrderData[index][0] || "").trim();
      const rowJob = String(changeOrderData[index][1] || "").trim();
      if (rowId === changeOrderId && rowJob === job) {
        description = String(changeOrderData[index][2] || "").trim();
        break;
      }
    }
  }
  if (!description) throw new Error("The selected change order could not be found for this job.");

  const dateOnly = new Date(timestamp.getFullYear(), timestamp.getMonth(), timestamp.getDate());
  const timeOnly = createTimeOnly(timestamp);
  changeOrderLog.insertRowsBefore(2, 1);
  changeOrderLog.getRange(2, 1, 1, 7).setValues([[dateOnly, employee, job, changeOrderId, description, timeOnly, ""]]);
  changeOrderLog.getRange(2, 9).setValue(notes);
  changeOrderLog.getRange(2, 1).setNumberFormat("M/d/yy (dddd)");
  changeOrderLog.getRange(2, 6, 1, 2).setNumberFormat("h:mm AM/PM");
  SpreadsheetApp.flush();
}

function getEmployeeClockStatus(employee) {
  const name = String(employee || "").trim();
  if (!name) return {clockedIn:false, clockInTime:"", job:""};
  const timeLog = getJmhSpreadsheet_().getSheetByName("Time Log");
  if (!timeLog) throw new Error('The "Time Log" sheet was not found.');
  const timestamp = new Date();
  const openRow = findOpenRow(timeLog, name, timestamp);
  if (!openRow) return {clockedIn:false, clockInTime:"", job:""};
  const startTime = timeLog.getRange(openRow, 4).getValue();
  const job = String(timeLog.getRange(openRow, 3).getDisplayValue() || "").trim();
  return {
    clockedIn:true,
    clockInTime:Utilities.formatDate(startTime, Session.getScriptTimeZone(), "h:mm a"),
    job:job
  };
}

function initializeJobCoordinates() {
  const spreadsheet = getJmhSpreadsheet_();
  const sheet = spreadsheet.getSheetByName("Jobs");
  if (!sheet || sheet.getLastRow() < 2) return [];
  const jobs = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getDisplayValues().flat();
  return jobs.filter(Boolean).map(function (job) {
    const policy = getJobClockInPolicy(spreadsheet, job);
    return {job:job, latitude:policy.latitude, longitude:policy.longitude};
  });
}

function openManagerPortal() {
  const url = "https://script.google.com/macros/s/AKfycbwqzwDp_vGM9wO6HTadN12gzEUhnli1nvHeiL2NxC3CJPglZu2MGjBbtaghv-GBPid1cQ/exec?page=manager";
  const html = HtmlService.createHtmlOutput(`<script>window.open("${url}", "_blank");google.script.host.close();</script>`).setWidth(10).setHeight(10);
  SpreadsheetApp.getUi().showModalDialog(html, "Opening Manager Portal...");
}

function getCrewHours(employee, passcode, deviceToken) {
  employee = String(employee || "").trim();
  passcode = String(passcode || "").trim();
  if (!employee) throw new Error("Please select your name.\nSeleccione su nombre.");
  verifyCrewDeviceOrPasscode(employee, passcode, deviceToken);
  const sheet = getJmhSpreadsheet_().getSheetByName("Time Log");
  if (!sheet) throw new Error('The "Time Log" sheet was not found.');

  const now = new Date();
  const currentStart = crewWeekStart_(now);
  const nextStart = new Date(currentStart); nextStart.setDate(nextStart.getDate() + 7);
  const previousStart = new Date(currentStart); previousStart.setDate(previousStart.getDate() - 7);
  let currentHours = 0, previousHours = 0;

  if (sheet.getLastRow() >= 2) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
    rows.forEach(function (row) {
      if (String(row[1] || "").trim().toLowerCase() !== employee.toLowerCase()) return;
      const workDate = row[0]; if (!(workDate instanceof Date)) return;
      const day = new Date(workDate.getFullYear(), workDate.getMonth(), workDate.getDate());
      let paidHours = Number(row[7]);
      if (!Number.isFinite(paidHours) || row[7] === "") {
        const start = row[3], end = row[4]; if (!(start instanceof Date)) return;
        const endTime = end instanceof Date ? end : now;
        const startMinutes = start.getHours()*60 + start.getMinutes() + start.getSeconds()/60;
        let endMinutes = endTime.getHours()*60 + endTime.getMinutes() + endTime.getSeconds()/60;
        if (endMinutes < startMinutes) endMinutes += 24*60;
        paidHours = Math.max(0, (endMinutes-startMinutes)/60) + (Number(row[6]) || 0);
      }
      if (day >= currentStart && day < nextStart) currentHours += paidHours;
      else if (day >= previousStart && day < currentStart) previousHours += paidHours;
    });
  }
  const timezone = getJmhSpreadsheet_().getSpreadsheetTimeZone();
  return {
    employee:employee,
    currentHours:Math.round(currentHours*100)/100,
    previousHours:Math.round(previousHours*100)/100,
    currentLabel:crewWeekLabel_(currentStart, timezone),
    previousLabel:crewWeekLabel_(previousStart, timezone),
    updatedAt:Utilities.formatDate(now, timezone, "M/d/yy h:mm a")
  };
}

function crewWeekStart_(date) {
  const start = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const daysSinceMonday = (start.getDay() + 6) % 7;
  start.setDate(start.getDate() - daysSinceMonday);
  return start;
}

function crewWeekLabel_(start, timezone) {
  const end = new Date(start); end.setDate(end.getDate() + 6);
  return Utilities.formatDate(start, timezone, "M/d") + " - " + Utilities.formatDate(end, timezone, "M/d");
}

function switchJobAction(data) {
  data = data || {};
  if (!String(data.employee || "").trim()) throw new Error("Employee is required.");
  if (!String(data.job || "").trim() || !String(data.category || "").trim() || !String(data.phase || "").trim() || !String(data.task || "").trim()) {
    throw new Error("Select the new job, category, phase, and task.");
  }
  var baseKey = String(data.entryKey || (Date.now() + "-switch-job"));
  var clockOutData = JSON.parse(JSON.stringify(data));
  clockOutData.action = "clockOut";
  clockOutData.entryKey = baseKey + "-out";
  submitTimeAction(clockOutData);
  var clockInData = JSON.parse(JSON.stringify(data));
  clockInData.action = "clockIn";
  clockInData.entryKey = baseKey + "-in";
  return submitTimeAction(clockInData);
}

function shouldInsertWeeklySeparator(sheet, timestamp) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  var dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var newestExistingDate = null;
  for (var i = 0; i < dates.length; i++) {
    if (dates[i][0] !== "" && dates[i][0] != null) { newestExistingDate = dates[i][0]; break; }
  }
  if (!newestExistingDate) return false;
  return mondayWeekKey_(newestExistingDate) !== mondayWeekKey_(timestamp);
}

function mondayWeekKey_(value) {
  var date = value instanceof Date ? new Date(value.getTime()) : new Date(value);
  if (isNaN(date.getTime())) return "";
  date.setHours(0, 0, 0, 0);
  date.setDate(date.getDate() - ((date.getDay() + 6) % 7));
  return date.getFullYear() + "-" + String(date.getMonth() + 1).padStart(2, "0") + "-" + String(date.getDate()).padStart(2, "0");
}
