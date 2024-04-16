import puppeteer from 'puppeteer-extra';
import StealthPlugin from 'puppeteer-extra-plugin-stealth'
puppeteer.use(StealthPlugin());
import {writeFile} from 'fs';
import XLSX from 'xlsx'

async function getPinNumbersFromExcel(filePath) {
    try {
        const workbook = XLSX.readFile(filePath);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const pins = XLSX.utils.sheet_to_json(worksheet, { raw: false });
        return pins.map(row => row['Tracking #']);
    } catch (error) {
        console.error('Error reading Excel file:', error);
        return [];
    }
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

function formatDateString(dateString) {
    // Split the string by space
    const parts = dateString.split(" ");
    // Extract Month and Day parts
    const month = parts[1];
    const day = parts[2].slice(0, -1); // Remove comma
    return `2024 ${month} ${day}`;
}

(async() => {

	const startDate = new Date(); // Start time for the script

	const filename = process.argv.length >= 3 ? process.argv[2] : 'PuroMar5.xlsx'; // Change this to the name of your Excel file
	const pinFilePath = `./data/${filename}`; // Change this to the path of your Excel file
    const pinNumbers = await getPinNumbersFromExcel(pinFilePath);

    const data = [];

	let amountPins = 0
	const amountPinsTotal = pinNumbers.length;
	let totalIterationTime = 0;
	const browser = await puppeteer.launch();
	for (const pin of pinNumbers) {
		let dateString;
		const iterationStartTime = new Date(); // Start time for the iteration
		amountPins++;
		console.log(`Processing PIN: ${pin}, ${amountPins} of ${amountPinsTotal}`);
		const page = await browser.newPage();
		await page.setViewport({
			width: 1600,
			height: 1000,
			isMobile: false,
			isLandscape: true,
			hasTouch: false,
			deviceScaleFactor: 1});

		await page.goto(`https://www.purolator.com/en/shipping/tracker?pin=${pin}`);

		try { 
			await Promise.all([
				page.waitForSelector('#tracking-detail > div.detailed-view.DEL > div:nth-child(5) > div.col-12.col-sm-7 > p', { timeout: 7000 }),
				page.waitForSelector('#tracking-detail > div.detailed-view.DEL > div.row.border-top.pt-2 > div.col-12.col-sm-4.col-md-4.col-lg-4.pl-sm-0.order-3 > div:nth-child(3) > div.col-7.col-sm-12.col-md-7', { timeout: 7000 })
			])
		} catch (error) {
			console.log(`Skipping PIN ${pin} due to missing or empty date information.`);
            await page.close();
            continue;
		}
		
		// Get all the table rows
		const rows = await page.$$('tbody tr');

		// Iterate through each row starting from the last one
		for (let i = rows.length - 1; i >= 0; i--) {
    		const row = rows[i];
			// Get the last cell in the row
			const lastCellText = await page.evaluate(row => row.querySelector('td:last-child').textContent.trim(), row);
			// Check if the last cell contains the text "Picked up by Purolator"
			if (lastCellText === "Picked up by Purolator") {
				// If it does, get the date string from the first cell in the same row
				dateString = await page.evaluate(row => row.querySelector('td:first-child').textContent.trim(), row);

				// Optionally, you can break the loop here if you only want to find the first occurrence
				break;
			}
		}

		if (!dateString) {
			dateString = await page.$eval('#tracking-detail > div.detailed-view.DEL > div.row.border-top.pt-2 > div.col-12.col-sm-4.col-md-4.col-lg-4.pl-sm-0.order-3 > div:nth-child(3) > div.col-7.col-sm-12.col-md-7', (el) => el.innerText);
		}

		const deliveryDateStr = await page.$eval('#tracking-detail > div.detailed-view.DEL > div:nth-child(5) > div.col-12.col-sm-7 > p', (el) => el.innerText);
	
		await page.close()
		const deliveryDate = formatDateString(deliveryDateStr);
		const shippingDate = formatDateString(dateString);

		// Calculate business days
		const businessDays = getBusinessDays(shippingDate, deliveryDate);
		console.log("Business days between shipping and delivery:", businessDays);

		data.push([pin, shippingDate, deliveryDate, businessDays]);
		const iterationEndTime = new Date(); // End time for the iteration
        const iterationDuration = (iterationEndTime - iterationStartTime) / 1000; // Duration of the iteration in seconds
        console.log(`Iteration took ${iterationDuration.toFixed(2)} seconds.`);
		totalIterationTime += iterationDuration;
		const averageIterationTime = totalIterationTime / amountPins;
    	const estimatedTotalTime = (pinNumbers.length - amountPins) * averageIterationTime; // Estimate remaining time
    	console.log(`Estimated total runtime: ${estimatedTotalTime.toFixed(2)} seconds.`);
		console.log('-------------------------------------------');
	}
	await browser.close();

	// Write data to Excel file
    const ws = XLSX.utils.aoa_to_sheet([['PIN', 'Shipping Date', 'Delivery Date', 'Business Days'], ...data]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Tracking Data');
    writeFile(`./data/results${filename}`, XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }), (err) => {
        if (err) throw err;
        console.log(`Data has been written to results${filename}`);
    });
	const endDate = new Date(); // End time for the script
	const scriptDuration = (endDate - startDate) / 1000; // Duration of the script in seconds
	print(`Script took ${scriptDuration.toFixed(2)} seconds.`);
})();

