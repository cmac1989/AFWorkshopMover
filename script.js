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

    for (let i = 1; i < allData.length; i++) {
        const registrationData = extractWorkshopInfo(allData[i]);
        if (!registrationData) continue;

        const rowIndex = targetRows[i];
        let logEntry;

        if (rowIndex) {
            let message;
            if (!previewOnly) {
                applyWorkshopUpdate(hqWorkshopSheet, rowIndex, registrationData);
                message = `✅ Row ${rowIndex} updated successfully (${registrationData.title} - ${registrationData.firstName} ${registrationData.lastName})`;
            } else {
                message = simulateWorkshopUpdate(hqWorkshopSheet, rowIndex, registrationData);
            }

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

    let titleText = (workshopData[0] || "").trim();

    // Clean up common prefixes
    titleText = titleText.replace(/^animal flow\s*/i, "").replace(/\s*\(.*\)$/, "");

    let level = null;
    let location = null;

    // Match both "Animal Flow Level 1 Lausanne" and "Lausanne L1"
    let match = titleText.match(/(?:level|nivel)?\s*(\d+)\s+(.+)/i);
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

    if (!location || !level) {
        return {
            title: titleText,
            firstName: workshopData[8],
            lastName: workshopData[9],
            isInvalid: true
        };
    }

    const ticketValue = (workshopData[5] || "").trim();
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
        waiver: workshopData[22],
        resides: workshopData[31],
        phoneNumber: workshopData[33],
        metaData: workshopData[37],
        isEmailSame: workshopData[39],
        fullName: workshopData[40],
        attendeeEmail: workshopData[41]
    };
}

// === Improved row matching ===
function findWorkshopRow(newWorkshopData, sheetData, bgColors) {
    const nextAvailableRows = {};
    const workshopRowMap = {};

    for (let i = 0; i < newWorkshopData.length; i++) {
        const workshopInfo = extractWorkshopInfo(newWorkshopData[i]);
        if (!workshopInfo) continue;

        const normalizedTitle = normalize(removeAccents(workshopInfo.title));
        const normalizedFirst = normalize(removeAccents(workshopInfo.firstName));
        const normalizedLast = normalize(removeAccents(workshopInfo.lastName));

        // Balance updates → title + name
        if (workshopInfo.isBalanceUpdate) {
            for (let j = 0; j < sheetData.length; j++) {
                const hqTitle = normalize(removeAccents(sheetData[j][0]));
                const hqFirst = normalize(removeAccents(sheetData[j][8]));
                const hqLast = normalize(removeAccents(sheetData[j][9]));

                if (
                    (hqTitle.includes(normalizedTitle) ||
                        normalizedTitle.includes(hqTitle) ||
                        isLooselyMatchingTitle(hqTitle, normalizedTitle)) &&
                    hqFirst === normalizedFirst &&
                    hqLast === normalizedLast
                ) {
                    workshopRowMap[i] = j + 1;
                    break;
                }
            }
            continue;
        }

        // New registrations → find first available slot by title
        const title = normalizedTitle;
        let startRow = nextAvailableRows[title] ? nextAvailableRows[title] + 1 : 1;

        for (let j = startRow - 1; j < sheetData.length; j++) {
            const ticketCell = (sheetData[j][5] || "").trim();
            const totalCell = sheetData[j][6];
            const bgColor = bgColors[j][0];
            const hqTitle = normalize(removeAccents(sheetData[j][0]));

            if (
                (hqTitle.includes(title) ||
                    title.includes(hqTitle) ||
                    isLooselyMatchingTitle(hqTitle, title)) &&
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

// --- Helper: Normalize & remove accents ---
function removeAccents(str) {
    return (str || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

// --- Helper: fuzzy title matching ---
function isLooselyMatchingTitle(a, b) {
    const cleanA = removeAccents(a).replace(/[^a-z0-9\s]/g, " ");
    const cleanB = removeAccents(b).replace(/[^a-z0-9\s]/g, " ");

    if (cleanA.includes(cleanB) || cleanB.includes(cleanA)) return true;

    const tokensA = cleanA.split(/\s+/).filter(t => t.length > 2);
    const tokensB = cleanB.split(/\s+/).filter(t => t.length > 2);
    let matchCount = 0;

    tokensA.forEach(token => {
        if (tokensB.includes(token)) matchCount++;
    });

    const similarity = matchCount / Math.max(tokensA.length, tokensB.length);

    if (similarity > 0.4) return true;

    const mainWordA = extractMainLocation(cleanA);
    const mainWordB = extractMainLocation(cleanB);
    return mainWordA && mainWordB && mainWordA === mainWordB;
}

// --- Helper: extract main location word for fallback match ---
function extractMainLocation(text) {
    const words = text.split(/\s+/).filter(w =>
        w.length > 2 &&
        !["animal", "flow", "level", "l", "workshop", "training", "eng", "esp", "prt"].includes(w)
    );
    return words.length ? words[0] : null;
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
    const logs = moveAllRowsToWorkshop(true);
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

// === Writes actual updates to HQ sheet and colors row green ===
function applyWorkshopUpdate(sheet, rowIndex, registrationData) {
    const range = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn());
    const rowValues = range.getValues()[0];

    const isL1 = registrationData.title?.toUpperCase().includes("L1");
    const isL2 = registrationData.title?.toUpperCase().includes("L2");

    registrationData.totalCost = registrationData.title.toUpperCase().includes("L2") || registrationData.title.toUpperCase().includes("L3") ? 795 : 695;

    if (registrationData.isBalanceUpdate) {
        const paidCol = 19;
        const currentPaid = parseFloat(rowValues[paidCol - 1]) || 0;
        const newPaid = currentPaid + parseFloat(registrationData.amountPaid || 0);
        sheet.getRange(rowIndex, paidCol).setValue(newPaid);
    } else {
        const fieldMap = {
            title: 1, ticket: 6, transactionNumber: 8, firstName: 9, lastName: 10,
            email: 11, date: 13, totalCost: 14, coupon: 15, couponCode: 16,
            affiliate: 17, customerNotes: 18, amountPaid: 19, balance: 20,
            waiver: 23, resides: 32, phoneNumber: 34, metaData: 38,
            isEmailSame: 40, fullName: 41, attendeeEmail: 42
        };

        for (const key in fieldMap) {
            if (registrationData[key] !== "" && registrationData[key] != null) {
                sheet.getRange(rowIndex, fieldMap[key]).setValue(registrationData[key]);
            }
        }

        if (registrationData.metaData && registrationData.title) {
            if (isL1) {
                const meta = parseMetaData(registrationData.metaData);
                if (meta) {
                    const metaFieldMap = { travel: 33, source: 35, certPlan: 36, certCategory: 37 };
                    for (const key in metaFieldMap) {
                        if (meta[key] && meta[key] !== "") sheet.getRange(rowIndex, metaFieldMap[key]).setValue(meta[key]);
                    }
                }
            } else if (isL2) {
                const metaL2 = parseMetaDataL2(registrationData.metaData);
                if (metaL2) {
                    const metaFieldMapL2 = { travel: 33, locationDate: 32, certStatus: 35 };
                    for (const key in metaFieldMapL2) {
                        if (metaL2[key] && metaL2[key] !== "") sheet.getRange(rowIndex, metaFieldMapL2[key]).setValue(metaL2[key]);
                    }
                }
            }
        }

        sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).setBackground('#c6efce');
        sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).setFontSize(10).setFontColor('#000000');
    }
}

function updateSelectedRows(selectedIndexes) {
    const cache = CacheService.getScriptCache();
    const ssUrl = 'https://docs.google.com/spreadsheets/d/16AOaXrNDgAIwfVaXsl3dIs35pkHwtd1exNYomT8ryqs/edit';
    const ss = SpreadsheetApp.openByUrl(ssUrl);
    const hqSheet = ss.getSheetByName("HQ Workshops 2025");

    const logs = moveAllRowsToWorkshop(false);
    const selectedLogs = selectedIndexes.map(i => logs[i]);

    const backupRows = selectedIndexes.map(i => logs[i].row).filter(r => r && r !== 'N/A');
    const backupData = backupRows.map(r => ({
        row: r,
        values: hqSheet.getRange(r, 1, 1, hqSheet.getLastColumn()).getValues()[0]
    }));

    cache.put('backupData', JSON.stringify(backupData), 3600);

    selectedLogs.forEach(entry => {
        if (entry.row && entry.row !== 'N/A') {
            const rowIndex = entry.row;
            hqSheet.getRange(rowIndex, 1, 1, hqSheet.getLastColumn()).setBackground('#ccffcc');
            hqSheet.getRange(rowIndex, hqSheet.getLastColumn()).setNote(`Updated ${new Date().toLocaleString()}`);
        }
    });

    return `${selectedLogs.length} rows updated successfully. You can revert within 1 hour.`;
}

function parseMetaData(metaString) {
    if (!metaString) return null;
    const parts = metaString.split(/,(?=(?:[^"]*"[^"]*")*[^"]*$)/).map(s => s.trim());
    let certPlan = (parts[3] || "").split(":")[0].trim();
    return { source: parts[2] || "", certPlan, certCategory: parts[4] || "", travel: parts[parts.length - 1] || "" };
}

function parseMetaDataL2(metaString) {
    if (!metaString) return null;
    const parts = metaString.split(/,(?=(?:[^"]*"[^"]*")*[^"]*$)/).map(s => s.trim());
    if (parts.length < 5) return null;
    const certStatus = parts[parts.length - 2] || "";
    const travel = parts[parts.length - 1] || "";
    const locationDate = parts.slice(2, parts.length - 2).join(", ").trim();
    return { travel, locationDate, certStatus };
}

function revertLastChanges() {
    const cache = CacheService.getScriptCache();
    const backup = cache.get('backupData');
    if (!backup) return 'No recent backup found — nothing to revert.';

    const data = JSON.parse(backup);
    const ssUrl = 'https://docs.google.com/spreadsheets/d/16AOaXrNDgAIwfVaXsl3dIs35pkHwtd1exNYomT8ryqs/edit';
    const ss = SpreadsheetApp.openByUrl(ssUrl);
    const hqSheet = ss.getSheetByName("HQ Workshops 2025");

    data.forEach(item => {
        hqSheet.getRange(item.row, 1, 1, item.values.length).setValues([item.values]);
    });

    cache.remove('backupData');
    return 'Selected rows have been restored from backup successfully.';
}
