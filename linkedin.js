const playwright = require('playwright');
const XLSX = require('xlsx');

async function readExcel(filePath) {
  const workbook = XLSX.readFile(filePath);
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

  const profiles = [];
  for (let i = 1; i < data.length; i++) {
    const [profileUrl, message, status] = data[i];
    profiles.push({ profileUrl, message, status, rowNumber: i + 1 });
  }
  
  return { profiles, worksheet, workbook };
}

async function updateExcel(worksheet, rowNumber, status, workbook, filePath) {
  const statusCell = worksheet[`C${rowNumber}`];
  if (statusCell) {
    statusCell.v = status;
    XLSX.writeFile(workbook, filePath);
    console.log(`Updated status for row ${rowNumber}: ${status}`);
  }
}

(async () => {
  const browser = await playwright.chromium.launch({ headless: false, slowMo: 3000 });
  const context = await browser.newContext();
  const page = await context.newPage();

  console.log("Logging into LinkedIn...");
  await page.goto('https://www.linkedin.com/login', { waitUntil: 'domcontentloaded' });


// PUT YOUR MAIL AND PASSWORD HERE

  await page.fill('#username', 'EMAIL');
  await page.fill('#password', 'PASSWORD');  
  await page.click('[type="submit"]');
  await page.waitForTimeout(5000);

  console.log("Reading data from Excel...");
  const { profiles, worksheet, workbook } = await readExcel('LinkedInProfiles.xlsx');
  console.log("Starting to process profiles...");

  for (const { profileUrl, message, status, rowNumber } of profiles) {
    if (status === 'Ok') {
      console.log(`Skipping ${profileUrl}, already processed.`);
      continue;
    }

    console.log(`Processing: ${profileUrl}`);
    try {
      await page.goto(profileUrl, { waitUntil: 'domcontentloaded' });
      console.log(`Navigated to: ${profileUrl}`);

      const connectButton = await page.locator('[data-test-icon="connect-small"]').nth(1);
      const moreButton = await page.locator('[aria-label="More actions"]').nth(1);

    try {
  await connectButton.waitFor({ state: 'attached', timeout: 5000 });
  try {
    await connectButton.click();
    console.log("Clicked 'Connect' button.");
  } catch (error) {
    console.log("Connect button not visible, clicking 'More' button.");
    await moreButton.click();
    await page.waitForTimeout(1000);
    await connectButton.click();
    console.log("Clicked 'Connect' button after opening the 'More' menu.");
  }
  const addNoteButton = await page.locator('button:has-text("Add a note")');
  await addNoteButton.click();
  console.log("Clicked 'Add a note' button.");

  const messageBox = await page.locator('textarea[name="message"]');
  await messageBox.fill(message);
  console.log(`Filled message: ${message}`);

  const sendButton = await page.locator('button:has-text("Send")');
  await sendButton.click();
  console.log("Clicked 'Send' button.");

  await updateExcel(worksheet, rowNumber, 'Ok', workbook, 'LinkedInProfiles.xlsx');
  console.log(`Connection request sent to: ${profileUrl}`);
    } catch (error) {
      console.error("Error while trying to click the Connect button or Add a note button:", error);
      await updateExcel(worksheet, rowNumber, 'Not ok', workbook, 'LinkedInProfiles.xlsx');
    }


    } catch (error) {
      console.error(`Error processing ${profileUrl}:`, error);
      await updateExcel(worksheet, rowNumber, 'Failed, Error occurred', workbook, 'LinkedInProfiles.xlsx');
    }
  }

  console.log("All profiles processed.");
  await browser.close();
})();
