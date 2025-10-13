function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Workshop Tools')
        .addItem('Open Dashboard', 'openDashboard')
        .addToUi();
}

function openDashboard() {
    const html = HtmlService.createHtmlOutputFromFile('Dashboard')
        .setWidth(900)
        .setHeight(600);
    SpreadsheetApp.getUi().showSidebar(html);
}

// === MAIN FUNCTION ===
// Collects logs — PREVIEW MODE by default (does NOT update sheet)
function moveAllRowsToWorkshop(previewOnly = true) {
    const ssUrl = 'https://docs.google.com/spreadsheets/d/16AOaXrNDgAIwfVaXsl3dIs35pkHwtd1exNYomT8ryqs/edit';
    const ss = SpreadsheetApp.openByUrl(ssUrl);
    const newHQSheet = ss.getSheetByName("NEW HQ");
    const hqWorkshopSheet = ss.getSheetByName("HQ Workshops 2025");

    const allData = newHQSheet.getDataRange().getValues();
    const hqData = hqWorkshopSheet.getDataRange().getValues();
    const bgColors = hqWorkshopSheet.getRange(1, 1, hqData.length, 1).getBackgrounds();

    const targetRows = findWorkshopRow(allData, hqData, bgColors);
    const logEntries = [];

    for (let i = 0; i < allData.length; i++) {
        const registrationData = extractWorkshopInfo(allData[i]);
        if (!registrationData) continue;

        const rowIndex = targetRows[i];
        let logEntry;

        if (rowIndex) {
            const message = simulateWorkshopUpdate(hqWorkshopSheet, rowIndex, registrationData);
            logEntry = {
                row: rowIndex,
                type: registrationData.isBalanceUpdate ? 'BALANCE update' : 'NEW registration',
                details: message
            };
        } else {
            logEntry = {
                row: 'N/A',
                type: 'ERROR',
                details: `No matching target row found for ${registrationData.title} (${registrationData.firstName} ${registrationData.lastName})`
            };
        }

        logEntries.push(logEntry);
    }

    return logEntries;
}

// === Extracts workshop registration info from sheet data ===
function extractWorkshopInfo(workshopData) {
    if (!workshopData || !workshopData[0]) return null;

    const titleText = workshopData[0];
    let level = null;
    let location = null;

    let match = titleText.match(/Animal Flow\s+(?:Level|Nivel)\s*(\d+)\s+(.+?)\s*\(/i);
    if (match) {
        level = parseInt(match[1], 10);
        location = match[2].trim();
    } else {
        match = titleText.match(/(.+?)\s+L(\d+)/i);
        if (match) {
            location = match[1].trim();
            level = parseInt(match[2], 10);
        }
    }

    const ticketValue = (workshopData[5] || "").trim();
    if (ticketValue.toUpperCase() === "TICKET") return null;

    const isBalanceUpdate = ticketValue.toUpperCase() === "BALANCE";
    const workshopTitle = `${location} L${level}`;

    return {
        title: workshopTitle,
        ticket: ticketValue,
        transactionNumber: workshopData[7],
        firstName: workshopData[8],
        lastName: workshopData[9],
        email: workshopData[10],
        date: workshopData[12],
        totalCost: workshopData[13],
        coupon: workshopData[14],
        couponCode: workshopData[15],
        affiliate: workshopData[16],
        customerNotes: workshopData[17],
        amountPaid: workshopData[18],
        balance: workshopData[19],
        isBalanceUpdate: isBalanceUpdate,
    };
}

// === Finds the correct rows for updates ===
function findWorkshopRow(newWorkshopData, sheetData, bgColors) {
    const nextAvailableRows = {};
    const workshopRowMap = {};

    for (let i = 0; i < newWorkshopData.length; i++) {
        const workshopInfo = extractWorkshopInfo(newWorkshopData[i]);
        if (!workshopInfo) continue;

        // Balance updates: find matching row by name and title
        if (workshopInfo.isBalanceUpdate) {
            for (let j = 0; j < sheetData.length; j++) {
                if (
                    normalize(sheetData[j][0]) === normalize(workshopInfo.title) &&
                    normalize(sheetData[j][8]) === normalize(workshopInfo.firstName) &&
                    normalize(sheetData[j][9]) === normalize(workshopInfo.lastName)
                ) {
                    workshopRowMap[i] = j + 1;
                    break;
                }
            }
            continue;
        }

        // New registrations: find first available white row
        const title = workshopInfo.title;
        let startRow = nextAvailableRows[title] ? nextAvailableRows[title] + 1 : 1;

        for (let j = startRow - 1; j < sheetData.length; j++) {
            const ticketCell = (sheetData[j][5] || "").trim();
            const totalCell = sheetData[j][6];
            const bgColor = bgColors[j][0];

            if (
                normalize(sheetData[j][0]) === normalize(title) &&
                ticketCell === "" &&
                totalCell != 25 &&
                (bgColor === "#ffffff" || bgColor.toLowerCase() === "white")
            ) {
                workshopRowMap[i] = j + 1;
                nextAvailableRows[title] = j + 1;
                break;
            }
        }
    }

    return workshopRowMap;
}

// === Formats and simulates updates ===
function simulateWorkshopUpdate(sheet, rowIndex, registrationData) {
    if (registrationData.isBalanceUpdate) {
        return `BALANCE update → Row ${rowIndex} would add <b>${registrationData.amountPaid}</b> to the Paid column.`;
    } else {
        const updates = {};
        for (const key in registrationData) {
            if (registrationData[key] !== "" && registrationData[key] !== null) updates[key] = registrationData[key];
        }
        return `NEW registration → Row ${rowIndex} would be updated with:<br>` +
            Object.entries(updates)
                .map(([k, v]) => `<strong>${k}:</strong> ${v}`)
                .join('<br>');
    }
}

// === Sends selected logs via email ===
function sendSelectedLogsEmail(selectedIndexes) {
    const logs = moveAllRowsToWorkshop(false);
    const selectedLogs = selectedIndexes.map(i => logs[i]);

    MailApp.sendEmail({
        to: Session.getActiveUser().getEmail(),
        subject: 'Workshop Update Logs',
        htmlBody: buildHtmlEmail(selectedLogs)
    });
    return 'Email sent for selected entries!';
}

// === Builds HTML Email ===
function buildHtmlEmail(logEntries) {
    let html = `<h2 style="font-family:sans-serif;">Workshop Update Logs</h2>
  <table style="width:100%; border-collapse:collapse; font-family:sans-serif;">
    <tr style="background-color:#4CAF50; color:white;">
      <th style="padding:8px;border:1px solid #ddd;">Row</th>
      <th style="padding:8px;border:1px solid #ddd;">Type</th>
      <th style="padding:8px;border:1px solid #ddd;">Details</th>
    </tr>`;

    logEntries.forEach((entry, index) => {
        const bgColor = index % 2 === 0 ? '#f9f9f9' : '#ffffff';
        html += `<tr style="background-color:${bgColor};">
      <td style="padding:8px;border:1px solid #ddd;">${entry.row}</td>
      <td style="padding:8px;border:1px solid #ddd;">${entry.type}</td>
      <td style="padding:8px;border:1px solid #ddd;">${entry.details}</td>
    </tr>`;
    });

    html += '</table>';
    return html;
}

// === Returns data to the HTML dashboard ===
function getWorkshopLogs() {
    return moveAllRowsToWorkshop(true);
}

// === Helper to normalize strings (trim + lowercase) ===
function normalize(value) {
    return (value || "").toString().trim().toLowerCase();
}
