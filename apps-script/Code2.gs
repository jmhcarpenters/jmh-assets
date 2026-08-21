function getTasks(category, phase) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Tasks");

  if (!sheet) {
    throw new Error(
      'The "Tasks" sheet was not found.'
    );
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const selectedCategory = String(
    category || ""
  )
    .trim()
    .toLowerCase();

  const selectedPhase = String(
    phase || ""
  )
    .trim()
    .toLowerCase();

  if (
    !selectedCategory ||
    !selectedPhase
  ) {
    return [];
  }

  // Tasks sheet:
  // A = Category
  // B = Phase
  // C = Task
  const tasks = sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      3
    )
    .getDisplayValues()
    .filter(function (row) {
      const rowCategory = String(
        row[0] || ""
      )
        .trim()
        .toLowerCase();

      const rowPhase = String(
        row[1] || ""
      )
        .trim()
        .toLowerCase();

      return (
        rowCategory === selectedCategory &&
        rowPhase === selectedPhase
      );
    })
    .map(function (row) {
      return String(
        row[2] || ""
      ).trim();
    })
    .filter(function (task) {
      return task !== "";
    });

  return [...new Set(tasks)];
}


function getChangeOrders(job) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(
      "Change Orders"
    );

  if (!sheet) {
    throw new Error(
      'The "Change Orders" sheet was not found.'
    );
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const selectedJob = String(
    job || ""
  ).trim();

  const data = sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      3
    )
    .getDisplayValues();

  return data
    .filter(function (row) {
      const rowJob = String(
        row[1] || ""
      ).trim();

      return rowJob === selectedJob;
    })
    .map(function (row) {
      return {
        id: String(
          row[0] || ""
        ).trim(),

        description: String(
          row[2] || ""
        ).trim()
      };
    });
}


function getEmployees() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Employees");

  if (!sheet) {
    throw new Error(
      'The "Employees" sheet was not found.'
    );
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  // Employees sheet:
  // A = Employee name
  // F = Active
  const rows = sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      6
    )
    .getDisplayValues();

  return rows
    .filter(function (row) {
      const employeeName = String(
        row[0] || ""
      ).trim();

      const activeValue = String(
        row[5] || ""
      )
        .trim()
        .toLowerCase();

      const isActive =
        activeValue === "yes" ||
        activeValue === "true" ||
        activeValue === "active" ||
        activeValue === "1" ||
        activeValue === "";

      return (
        employeeName !== "" &&
        isActive
      );
    })
    .map(function (row) {
      return String(
        row[0] || ""
      ).trim();
    });
}

function finalizePayroll() {
  const ss = getJmhSpreadsheet_();
  const payroll = ss.getSheetByName("Payroll");
  const history = ss.getSheetByName("Payroll History");
  const ui = SpreadsheetApp.getUi();

  if (!payroll || !history) {
    ui.alert("Payroll or Payroll History sheet was not found.");
    return;
  }

  const namePrompt = ui.prompt(
    "Finalize Payroll — Admin Authorization",
    "Enter your full admin name (John Habitzruther, Matthew Habitzruther, or David Aguirre):",
    ui.ButtonSet.OK_CANCEL
  );
  if (namePrompt.getSelectedButton() !== ui.Button.OK) return;

  const adminName = String(namePrompt.getResponseText() || "").trim();
  const authorizedNames = [
    "john habitzruther",
    "matthew habitzruther",
    "david aguirre"
  ];
  if (authorizedNames.indexOf(rmNormalize(adminName)) === -1) {
    ui.alert("Only John, Matthew, or David can finalize payroll.");
    return;
  }

  const pinPrompt = ui.prompt(
    "Finalize Payroll — Admin Authorization",
    "Enter your four-digit admin passcode:",
    ui.ButtonSet.OK_CANCEL
  );
  if (pinPrompt.getSelectedButton() !== ui.Button.OK) return;

  try {
    requireAdminAccess(adminName, pinPrompt.getResponseText());
  } catch (error) {
    ui.alert("Payroll was not finalized. " + error.message);
    return;
  }

  const periodStart = payroll.getRange("N2").getValue();
  const periodEnd = payroll.getRange("O2").getValue();
  const payDate = payroll.getRange("H2").getValue();

  if (!(periodStart instanceof Date) || !(periodEnd instanceof Date)) {
    ui.alert("Please select a valid pay period first.");
    return;
  }

  const lastPayrollRow = payroll.getLastRow();
  const payrollData = payroll
    .getRange(7, 1, Math.max(lastPayrollRow - 6, 1), 5)
    .getValues()
    .filter(function(row) { return String(row[0] || "").trim() !== ""; });

  if (payrollData.length === 0) {
    ui.alert("There are no employees to finalize.");
    return;
  }

  const rowsToSave = payrollData.map(function(row) {
    return [periodStart, periodEnd, payDate, row[0], row[1], row[2], row[3], row[4]];
  });

  const timezone = ss.getSpreadsheetTimeZone();
  const periodLabel =
    Utilities.formatDate(periodStart, timezone, "M/d/yy") + " to " +
    Utilities.formatDate(periodEnd, timezone, "M/d/yy");

  const confirmation = ui.alert(
    "Finalize Payroll",
    "Finalize and lock payroll for " + periodLabel +
      "?\n\nPayroll remains live and editable unless an authorized admin confirms here.",
    ui.ButtonSet.YES_NO
  );
  if (confirmation !== ui.Button.YES) return;

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const sameDay = function(a, b) {
      return a instanceof Date && b instanceof Date &&
        a.getFullYear() === b.getFullYear() &&
        a.getMonth() === b.getMonth() &&
        a.getDate() === b.getDate();
    };

    let destinationRow = 2;
    const lastHistoryRow = history.getLastRow();

    if (lastHistoryRow >= 2) {
      const existingRange = history.getRange(2, 1, lastHistoryRow - 1, 8);
      const existingValues = existingRange.getValues();
      const existingFormulas = existingRange.getFormulas();
      const matchingIndexes = [];

      existingValues.forEach(function(row, index) {
        if (sameDay(row[0], periodStart) && sameDay(row[1], periodEnd)) {
          matchingIndexes.push(index);
        }
      });

      if (matchingIndexes.length > 0) {
        const hasLiveFormula = matchingIndexes.some(function(index) {
          return existingFormulas[index].some(function(formula) { return formula !== ""; });
        });
        if (!hasLiveFormula) {
          ui.alert("This pay period has already been finalized.\n\nNo duplicate was created.");
          return;
        }
        destinationRow = matchingIndexes[0] + 2;
      } else {
        history.insertRowsBefore(2, rowsToSave.length + 1);
      }
    } else {
      history.insertRowsBefore(2, rowsToSave.length + 1);
    }

    const finalizedRange = history.getRange(destinationRow, 1, rowsToSave.length, 8);
    finalizedRange.setValues(rowsToSave);
    history.getRange(destinationRow, 1, rowsToSave.length, 3).setNumberFormat("M/d/yy");
    history.getRange(destinationRow, 8, rowsToSave.length, 1).setNumberFormat("$#,##0.00");

    protectFinalizedPayrollRange_(finalizedRange, periodLabel);

    SpreadsheetApp.flush();
    const finalizedHistoricalData = finalizedRange.getValues();
    const delivery = deliverFinalizedPayrollSummaryToJohn_(
      finalizedHistoricalData,
      periodLabel,
      timezone
    );

    ui.alert(
      "Payroll finalized and locked successfully!\n\n" +
      periodLabel + "\n\n" + delivery
    );
  } finally {
    lock.releaseLock();
  }
}

function protectFinalizedPayrollRange_(range, periodLabel) {
  const protection = range.protect().setDescription("Finalized payroll — " + periodLabel);
  protection.setWarningOnly(false);
  const editors = protection.getEditors();
  if (editors.length) protection.removeEditors(editors);
  if (protection.canDomainEdit()) protection.setDomainEdit(false);
}

function deliverFinalizedPayrollSummaryToJohn_(historicalRows, periodLabel, timezone) {
  const recipient = findJohnPayrollEmail_();
  if (!recipient) {
    return "Delivery prepared, but John’s email was not found. Add an Email column for John on Employees or set the JOHN_PAYROLL_EMAIL script property.";
  }

  const payableRows = historicalRows.filter(function(row) {
    return String(row[3] || "").trim() !== "" &&
      (Number(row[4] || 0) !== 0 || Number(row[7] || 0) !== 0);
  });

  if (payableRows.length === 0) {
    return "No employees with payable hours were found in finalized Payroll History.";
  }

  const lines = payableRows.map(function(row) {
    const employee = String(row[3] || "").trim();
    const totalHours = Number(row[4] || 0);
    const regularHours = Number(row[5] || 0);
    const overtimeHours = Number(row[6] || 0);
    const payment = Number(row[7] || 0);
    return employee + ": " +
      totalHours.toFixed(2) + " total paid hrs (" +
      regularHours.toFixed(2) + " regular, " +
      overtimeHours.toFixed(2) + " OT), gross pay " +
      Utilities.formatString("$%.2f", payment);
  });

  const payDate = historicalRows.length && historicalRows[0][2] instanceof Date
    ? Utilities.formatDate(historicalRows[0][2], timezone, "M/d/yy")
    : "Not specified";

  const total = payableRows.reduce(function(sum, row) {
    return sum + Number(row[7] || 0);
  }, 0);

  const body = [
    "Finalized JMH payroll summary",
    "Pay period: " + periodLabel,
    "Pay date: " + payDate,
    "",
    lines.join("\n"),
    "",
    "Total payments: " + Utilities.formatString("$%.2f", total),
    "",
    "This automated payroll summary was sent by the JMH Construction Management 2026 spreadsheet."
  ].join("\n");

  const escapeHtml = function(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  };

  const formatCurrency = function(value) {
    return Number(value || 0).toLocaleString("en-US", {
      style: "currency",
      currency: "USD",
      minimumFractionDigits: 2
    });
  };

  const htmlRows = payableRows.map(function(row, index) {
    const background = index % 2 === 0 ? "#ffffff" : "#f3f7fb";
    return '<tr style="background:' + background + ';">' +
      '<td style="padding:9px 12px;border-bottom:1px solid #dfe5ec;text-align:left;">' +
        escapeHtml(String(row[3] || "").trim()) + '</td>' +
      '<td style="padding:9px 12px;border-bottom:1px solid #dfe5ec;text-align:right;">' +
        Number(row[4] || 0).toFixed(2) + '</td>' +
      '<td style="padding:9px 12px;border-bottom:1px solid #dfe5ec;text-align:right;">' +
        Number(row[5] || 0).toFixed(2) + '</td>' +
      '<td style="padding:9px 12px;border-bottom:1px solid #dfe5ec;text-align:right;">' +
        Number(row[6] || 0).toFixed(2) + '</td>' +
      '<td style="padding:9px 12px;border-bottom:1px solid #dfe5ec;text-align:right;white-space:nowrap;">' +
        formatCurrency(row[7]) + '</td>' +
    '</tr>';
  }).join("");

  const totalHours = payableRows.reduce(function(sum, row) {
    return sum + Number(row[4] || 0);
  }, 0);
  const totalRegular = payableRows.reduce(function(sum, row) {
    return sum + Number(row[5] || 0);
  }, 0);
  const totalOvertime = payableRows.reduce(function(sum, row) {
    return sum + Number(row[6] || 0);
  }, 0);

  const htmlBody =
    '<div style="font-family:Arial,Helvetica,sans-serif;color:#202124;max-width:850px;">' +
      '<h2 style="margin:0 0 14px;color:#174a7e;font-size:24px;">Finalized JMH Payroll Summary</h2>' +
      '<p style="margin:0 0 20px;line-height:1.55;">' +
        '<strong>Pay period:</strong> ' + escapeHtml(periodLabel) + '<br>' +
        '<strong>Pay date:</strong> ' + escapeHtml(payDate) +
      '</p>' +
      '<div style="overflow-x:auto;">' +
        '<table role="presentation" style="width:100%;border-collapse:collapse;min-width:620px;font-size:14px;">' +
          '<thead><tr style="background:#0b3b6d;color:#ffffff;">' +
            '<th style="padding:10px 12px;text-align:left;">Employee</th>' +
            '<th style="padding:10px 12px;text-align:right;">Total hours</th>' +
            '<th style="padding:10px 12px;text-align:right;">Regular</th>' +
            '<th style="padding:10px 12px;text-align:right;">OT</th>' +
            '<th style="padding:10px 12px;text-align:right;">Gross pay</th>' +
          '</tr></thead>' +
          '<tbody>' + htmlRows + '</tbody>' +
          '<tfoot><tr style="background:#eaf2fb;font-weight:bold;border-top:2px solid #174a7e;">' +
            '<td style="padding:10px 12px;text-align:left;">Total</td>' +
            '<td style="padding:10px 12px;text-align:right;">' + totalHours.toFixed(2) + '</td>' +
            '<td style="padding:10px 12px;text-align:right;">' + totalRegular.toFixed(2) + '</td>' +
            '<td style="padding:10px 12px;text-align:right;">' + totalOvertime.toFixed(2) + '</td>' +
            '<td style="padding:10px 12px;text-align:right;white-space:nowrap;">' + formatCurrency(total) + '</td>' +
          '</tr></tfoot>' +
        '</table>' +
      '</div>' +
      '<p style="margin:18px 0 0;color:#5f6368;font-size:12px;">' +
        'This automated payroll summary was sent by the JMH Construction Management 2026 spreadsheet.' +
      '</p>' +
    '</div>';

  MailApp.sendEmail({
    to: recipient,
    subject: "JMH Payroll Update — Finalized Pay Period " + periodLabel,
    body: body,
    htmlBody: htmlBody
  });

  return "The finalized Payroll History summary was emailed to John at " + recipient + ".";
}

function findJohnPayrollEmail_() {
  const configured = PropertiesService.getScriptProperties().getProperty("JOHN_PAYROLL_EMAIL");
  if (configured && configured.indexOf("@") > 0) return configured.trim();

  const sheet = getJmhSpreadsheet_().getSheetByName("Employees");
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return "";

  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0].map(function(value) { return rmNormalize(value); });
  let nameColumn = -1;
  let emailColumn = -1;

  ["employee", "employee name", "name"].some(function(header) {
    nameColumn = headers.indexOf(header);
    return nameColumn >= 0;
  });
  ["email", "email address", "work email"].some(function(header) {
    emailColumn = headers.indexOf(header);
    return emailColumn >= 0;
  });

  if (nameColumn < 0 || emailColumn < 0) return "";

  for (let i = 1; i < values.length; i++) {
    if (rmNormalize(values[i][nameColumn]).indexOf("john") === 0) {
      const email = String(values[i][emailColumn] || "").trim();
      if (email.indexOf("@") > 0) return email;
    }
  }
  return "";
}

function createChangeOrder(data) {
  data = data || {};
  const job = String(data.job || "").trim();
  const description = String(data.description || "").trim();
  const location = String(data.location || "").trim();
  if (!job || !description) {
    throw new Error("Job and description are required.");
  }
  const sheet = getJmhSpreadsheet_().getSheetByName("Change Orders");
  if (!sheet) throw new Error('The "Change Orders" sheet was not found.');
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const lastRow = sheet.getLastRow();
    const ids = lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat();
    let highest = 0;
    ids.forEach(function (id) {
      const match = String(id || "").match(/^CO-(\d+)$/i);
      if (match) highest = Math.max(highest, Number(match[1]));
    });
    const id = "CO-" + String(highest + 1).padStart(3, "0");
    // Insert beneath the header instead of appendRow(); spilled formulas make
    // otherwise empty bottom rows appear occupied. Match the live A:F schema:
    // ID, Job, Description, Location, Status, Billing Rate.
    sheet.insertRowsBefore(2, 1);
    sheet.getRange(2, 1, 1, 6).setValues([[
      id,
      job,
      description,
      location,
      "Open",
      ""
    ]]);
    SpreadsheetApp.flush();
    return { id: id, description: description, job: job };
  } finally {
    lock.releaseLock();
  }
}



function setJohnPayrollEmail() {
  PropertiesService.getScriptProperties().setProperty("JOHN_PAYROLL_EMAIL", "Jmhcarpenters@gmail.com");
  return "John payroll email configured.";
}
