function moveAllRowsToWorkshop() {
    const ssUrl = 'https://docs.google.com/spreadsheets/d/16AOaXrNDgAIwfVaXsl3dIs35pkHwtd1exNYomT8ryqs/edit';
    const ss = SpreadsheetApp.openByUrl(ssUrl);

    const newHQSheet = ss.getSheetByName("NEW HQ");
    const hqWorkshopSheet = ss.getSheetByName("HQ Workshops 2025");

    const allData = newHQSheet.getDataRange().getValues();
    const hqData = hqWorkshopSheet.getDataRange().getValues();
    const bgColors = hqWorkshopSheet.getRange(1, 1, hqData.length, 1).getBackgrounds(); // ✅ prefetch all backgrounds

    const targetRows = findWorkshopRow(allData, hqData, bgColors);

    for (let i = 0; i < allData.length; i++) {
        const registrationData = extractWorkshopInfo(allData[i]);
        if (!registrationData) continue;

        const rowIndex = targetRows[i]; // now it's tied to the specific entry
        if (rowIndex) {
            simulateWorkshopUpdate(hqWorkshopSheet, rowIndex, registrationData);
        }
    }
}


function extractWorkshopInfo(workshopData) {
    if (!workshopData || !workshopData[0]) return null;

    const titleText = workshopData[0];
    let level = null;
    let location = null;

    // Case 1: Matches "Animal Flow Level/Nivel <num> <Location>"
    let match = titleText.match(/Animal Flow\s+(?:Level|Nivel)\s*(\d+)\s+(.+?)\s*\(/i);

    if (match) {
        level = parseInt(match[1], 10);
        location = match[2].trim();
    } else {
        // Case 2: Matches "<Location> L<num>" OR "Virtual L<num>"
        match = titleText.match(/(.+?)\s+L(\d+)/i);
        if (match) {
            location = match[1].trim();
            level = parseInt(match[2], 10);
        }
    }

    const ticketValue = (workshopData[5] || "").trim();
    if (ticketValue === "Ticket") return null; // skip header or invalid rows

    const isBalanceUpdate = ticketValue.toUpperCase() === "BALANCE"; // ignore whitespace / case
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


function findWorkshopRow(newWorkshopData, sheetData, bgColors) {
    const nextAvailableRows = {}; // store next available row per title
    const workshopRowMap = {};    // store all row indexes to return

    for (let i = 0; i < newWorkshopData.length; i++) {
        const workshopInfo = extractWorkshopInfo(newWorkshopData[i]);
        if (!workshopInfo) continue;

        // If this is a balance update, find the row that matches title + first + last name
        if (workshopInfo.isBalanceUpdate) {
            for (let j = 0; j < sheetData.length; j++) {
                const sheetTitle = (sheetData[j][0] || "").toString().trim().toLowerCase();
                const sheetFirst = (sheetData[j][8] || "").toString().trim().toLowerCase();
                const sheetLast  = (sheetData[j][9] || "").toString().trim().toLowerCase();

                const regTitle = workshopInfo.title.trim().toLowerCase();
                const regFirst = (workshopInfo.firstName || "").trim().toLowerCase();
                const regLast  = (workshopInfo.lastName || "").trim().toLowerCase();

                if (
                    sheetTitle === regTitle &&
                    sheetFirst === regFirst &&
                    sheetLast  === regLast
                ) {
                    workshopRowMap[i] = j + 1;
                    Logger.log(`BALANCE row found for index ${i} at sheet row ${j + 1}`);
                    break;
                }
            }
            continue; // skip normal row search
        }

        // Normal workflow: find first available empty row
        const title = workshopInfo.title;
        let startRow = nextAvailableRows[title] ? nextAvailableRows[title] + 1 : 1;

        for (let j = startRow - 1; j < sheetData.length; j++) {
            const ticketCell = sheetData[j][5];
            const totalCell = sheetData[j][6];
            const bgColor = bgColors[j][0];

            if (
                title === sheetData[j][0] &&
                ticketCell === "" &&
                totalCell != 25 &&
                (bgColor === "#ffffff" || bgColor.toLowerCase() === "white")
            ) {
                workshopRowMap[i] = j + 1; // map *this registration* to that specific row
                nextAvailableRows[title] = j + 1; // update pointer for next same workshop
                break;
            }
        }
    }

    return workshopRowMap;
}

function simulateWorkshopUpdate(sheet, rowIndex, registrationData) {
    const row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

    if (registrationData.isBalanceUpdate) {
        const existingPaid = parseFloat(row[18]) || 0; // column 19 = amountPaid
        const newPaid = parseFloat(registrationData.amountPaid) || 0;
        const updatedPaid = existingPaid + newPaid;

        const updates = {
            amountPaid: updatedPaid
        };

        Logger.log(
            `Row ${rowIndex} (BALANCE update) would be updated with: ${JSON.stringify(updates)}`
        );
    } else {
        const updates = {};
        for (const key in registrationData) {
            if (registrationData[key] !== "" && registrationData[key] !== null) {
                updates[key] = registrationData[key];
            }
        }

        Logger.log(
            `Row ${rowIndex} (NEW registration) would be updated with: ${JSON.stringify(updates)}`
        );
    }
}
