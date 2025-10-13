//MAIN FUNCTION
function moveAllRowsToWorkshop() {
    const ssUrl = 'https://docs.google.com/spreadsheets/d/16AOaXrNDgAIwfVaXsl3dIs35pkHwtd1exNYomT8ryqs/edit';
    const ss = SpreadsheetApp.openByUrl(ssUrl);

    //SPREADSHEETS
    const newHQSheet = ss.getSheetByName("NEW HQ");
    const hqWorkshopSheet = ss.getSheetByName("HQ Workshops 2025");
    const partnerSheet = ss.getSheetByName("Partners");

    if (!newHQSheet || !hqWorkshopSheet) {
        Logger.log("One of the sheets was not found!");
        return;
    }

    ``//CHECK FOR NEW HQ DATA
    const allData = newHQSheet.getDataRange().getValues();
    if(!allData) {
        Logger.log("No data found for New HQ Sheet");
    }

    //CHECK TO SEE IF HQ DATA IS PRESENT
    const hqData = hqWorkshopSheet.getDataRange().getValues();
    if(!hqData) {
        Logger.log("No data found for HQ Workshop Sheet");
    }

    //CHECK TO SEE IF PARTNER DATA IS PRESENT
    const partnerData = partnerSheet.getDataRange().getValues();
    if(!partnerData) {
        Logger.log("No data found for Partners Sheet");
    }

    if (allData.length === 0 || (allData.length === 1 && allData[0].every(cell => cell === ""))) {
        Logger.log("No data to move.");
        return;
    }

    findWorkshopRow(allData, hqData);
}

//////////////////////
// HELPER FUNCTIONS //
//////////////////////

function findWorkshopRow(newWorkshopData, sheetData) {
    const firstAvailableRows = {};

    for (let i = 0; i < newWorkshopData.length; i++) {
        const workshopName = extractWorkshopInfo(newWorkshopData[i]);
        Logger.log(workshopName)
        Logger.log(firstAvailableRows)
        // const workshopName = extractWorkshopInfo(newWorkshopData[i][0]);
        if(firstAvailableRows[workshopName]) continue;
        for (let j = 0; j < sheetData.length; j++) {
            const ticketCell = sheetData[j][5];
            const totalCell = sheetData[j][6];
            // Logger.log(sheetData[j][0]) - this data is correct
            if (workshopName == sheetData[j][0] && ticketCell == "" && totalCell != 25) {
                firstAvailableRows[workshopName] = j + 1
                break;
            }
        }
    }
    Logger.log(firstAvailableRows);
    return firstAvailableRows;
}

function extractWorkshopInfo(workshopData) {
    if (!workshopData || !workshopData[0]) {
        return {};
    }

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
    if (workshopData[5] == "Ticket") {
        return null;
    }
    if (workshopData[5] == "BALANCE") {
        // Logger.log(workshopData[19]);
        return workshopData[19];
    }

    const workshopTitle = `${location} L${level}`;

    const registrationData = {
        title: workshopTitle,
        ticket: workshopData[5],
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
    }
    return JSON.stringify(registrationData);
}



