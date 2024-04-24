import puppeteer from 'puppeteer-extra';
import StealthPlugin from 'puppeteer-extra-plugin-stealth'
puppeteer.use(StealthPlugin());
import {writeFile} from 'fs';
import XLSX from 'xlsx'
import os from 'os';
import { time } from 'console';

async function getPostalCodes(filePath) {
    try {
        const workbook = XLSX.readFile(filePath);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const pins = XLSX.utils.sheet_to_json(worksheet, { raw: false });
        return pins.map(row => ({
			originalPostalCode: row['Origin Postal Code'],
			destinationPostalCode: row['Destination Postal Code'],
			deliveryTime: row['Business Days'],
			deliveryStandard: row['Delivery Standard']
		}));
    } catch (error) {
        console.error('Error reading Excel file:', error);
        return [];
    }
}

// Function to convert full month name to abbreviated form
function getAbbreviatedMonth(fullMonth) {
    return fullMonth.substring(0, 3) + ".";
}

function getBusinessDays(startDate, endDate) {
    // Define weekend days (Saturday = 6, Sunday = 0)
    const weekendDays = [0, 6];
    
    // Convert string dates to Date objects
    startDate = new Date(startDate);
    endDate = new Date(endDate);
    
    // Calculate business days
    let businessDays = 0;
    while (startDate < endDate) {
        if (!weekendDays.includes(startDate.getDay())) {
            businessDays++;
        }
        startDate.setDate(startDate.getDate() + 1);
    }
    
    return businessDays;
}

(async() => {

	console.log("Starting script...");
	const platform = os.platform();
	const startDate = new Date(); // Start time for the script
    const filename = process.argv.length >= 3 ? process.argv[2] : 'PuroMar5.xlsx'; // Change this to the name of your Excel file
    const sourceFile = `./data/${filename}`; // Change this to the path of your Excel file
    const result = await getPostalCodes(sourceFile);

	const dataObject = {};
	let amountPins = 0
	const amountPinsTotal = result.length;
	let totalIterationTime = 0;
	let timeout;
	let browser;

	// Also add to save the file more often, like every 100 pin.

	if (platform === 'linux') {
		browser = await puppeteer.launch({executablePath: '/usr/bin/chromium-browser'});
		timeout = 25000;
	}
	else if (platform === 'win32') {
		browser = await puppeteer.launch({});
		timeout = 5000;
	}
	else {
		console.error('Unsupported platform:', platform);
		browser = await puppeteer.launch();
		timeout = 5000;
	}
	
	for (const pinObj of result) {
		const originPostalCode = pinObj.originalPostalCode;
		const destinationPostalCode = pinObj.destinationPostalCode;
		const deliveryTime = pinObj.deliveryTime;
		let deliveryStandard = pinObj.deliveryStandard;
		amountPins++;
		console.log(`Processing, ${amountPins} of ${amountPinsTotal}`);
        console.log(originPostalCode, destinationPostalCode);

		if (originPostalCode === undefined || destinationPostalCode === undefined) {
			console.log('Skipping PIN due to missing postalCode.');
			continue;
		}

        let dateString;
        const website = 'https://eshiponline.purolator.com/ShipOnline/estimates/estimate.aspx?lang=E'
		
		if (deliveryTime === '1') {
			console.log('Skipping PIN due to delivery time.');
			continue;
		} else if (deliveryStandard !== undefined && deliveryStandard !== null && deliveryStandard !== '') {
			console.log('Skipping PIN due to delivery standard already written.');
			continue;
		}

		const page = await browser.newPage();
		await page.setViewport({
			width: 1600,
			height: 1000,
			isMobile: false,
			isLandscape: true,
			hasTouch: false,
			deviceScaleFactor: 1});

		try {
			await page.goto(website, { timeout: `${timeout}` });
		} catch (error) {
			console.log(`Skipping PIN due to timeout.`);
			await page.close();
			continue;
		}
        try {
            await Promise.all([
                page.waitForSelector('#ctl00_CPPC_lblFrom', { timeout: timeout }),
                page.waitForSelector('#ctl00_CPPC_lblTo', { timeout: timeout })
            ])
        } catch (error) {
			console.log(`Skipping PIN due to missing elements.`);
            await page.close();
            continue;
		}
        try { await page.type('#ctl00_CPPC_ctrlShipFrom_txtPostalZipCode', originPostalCode, { delay: 100 });
        await page.type('#ctl00_CPPC_ctrlShipTo_txtPostalZipCode', destinationPostalCode, { delay: 100 });
        await page.click('#ctl00_CPPC_btnEstimate');
		} catch (error) {
			console.log(`Skipping PIN due to typing error.`);
			await page.close();
			continue;
		}
	

        try { await page.waitForSelector('#ctl00_CPPC_ctrlEstimateDisplay_gvEstimates', { timeout: timeout });
		} catch (error) {
			console.log(`Skipping PIN due to missing elements.`);
			await page.close();
			continue;
		}

        // Get all the table rows
		try {
			const rows = await page.$$('#ctl00_CPPC_ctrlEstimateDisplay_gvEstimates > tbody > tr');

        	for (let i = rows.length - 1; i >= 0; i--){
				const row = rows[i];
				const secondCellText = await page.evaluate(row=> {
					const spanElement = row.querySelector('td:nth-child(2) > span');
					const txtContent = spanElement ? spanElement.innerText : '' ;
					const extractedText = txtContent.split('\n')[0];

					return extractedText;
				}, row)
				
				if (secondCellText === 'Purolator Express') {
					dateString = await page.evaluate(row => row.querySelector('td:first-child > span').textContent.trim(), row);
					break;
				}


        	}
		if (dateString === undefined) {
			console.log('Skipping PIN due to missing date.');
			await page.click('#ctl00_CPPC_btnModifyStep1');
			await page.close();
			continue;
		}	
		} catch (error) {
			console.log(`Skipping PIN due to missing elements`);
			await page.close();
			continue;
		}

		// Extracting the date portion "April 26"
		var datePortion = dateString.substring(dateString.indexOf(",") + 2, dateString.indexOf("End of day")).trim()

		// Get today's date
		var today = new Date();

		// Define month names array
		// Define month names array
		var monthNames = ["January", "February", "March", "April", "May", "June",
							"July", "August", "September", "October", "November", "December"];

		// Get the current year, month, and day
		var currentYear = today.getFullYear();
		var currentMonth = monthNames[today.getMonth()];
		var currentDay = today.getDate();
		var formattedMonth = getAbbreviatedMonth(currentMonth);

		// Format today's date
		var formattedToday = currentYear + " " + formattedMonth + " " + currentDay;

		// Extract the month and day from the date portion
		var comparisonMonthIndex = monthNames.indexOf(datePortion.split(" ")[0]);
		var comparisonMonth = getAbbreviatedMonth(monthNames[comparisonMonthIndex]);
		var comparisonDay = parseInt(datePortion.split(" ")[1]);

		// Format the comparison date
		var formattedComparisonDate = currentYear + " " + comparisonMonth + " " + comparisonDay;

		deliveryStandard = getBusinessDays(formattedToday,formattedComparisonDate);

		// console.log("Formatted today's date: " + formattedToday);
		// console.log("Formatted comparison date: " + formattedComparisonDate);

        await page.click('#ctl00_CPPC_btnModifyStep1');
        
        await page.close();
        console.log('-------------------------------------------');

		dataObject[amountPins] = deliveryStandard;

    
    }
    await browser.close();

		// Load the existing workbook
	const workbook = XLSX.readFile(sourceFile);
	const worksheet = workbook.Sheets[workbook.SheetNames[0]];
	
	// Construct the data array to be added to the worksheet
	const dataToAdd = [['Delivery Standard']];
	for (let rowNum = 2; rowNum <= amountPinsTotal + 1; rowNum++) {
		const deliveryStandard = dataObject[rowNum - 1] || '';
		dataToAdd.push([deliveryStandard]);
	}

	// Specify the starting cell where to add the data (G1)
	const startingCell = { r: 0, c: 6 };

	// Loop through each row in the data array
	dataToAdd.forEach((rowData, rowIndex) => {
		const cellValue = worksheet[XLSX.utils.encode_cell({ r: rowIndex + startingCell.r, c: startingCell.c })];
		// Check if the cell is empty or undefined
		if (!cellValue || !cellValue.v) {
			// Add the data to the worksheet starting from the specified cell
			XLSX.utils.sheet_add_aoa(worksheet, [rowData], { origin: { r: rowIndex + startingCell.r, c: startingCell.c } });
		}
	});

	// Save the modified workbook back to the file
	XLSX.writeFile(workbook, sourceFile);

	console.log('Delivery standards have been written to the Excel file.');

})();