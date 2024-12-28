const playwright = require("playwright");
const XLSX = require("xlsx");

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
  const browser = await playwright.chromium.launch({
    headless: false,
    slowMo: 3000,
  });
  const context = await browser.newContext();
  const page = await context.newPage();

  console.log("Logging into LinkedIn...");
  await page.goto("https://www.linkedin.com/login", {
    waitUntil: "domcontentloaded",
  });

  // PUT YOUR MAIL AND PASSWORD HERE

  await page.fill("#username", "");
  await page.fill("#password", "");
  await page.click('[type="submit"]');
  await page.waitForTimeout(5000);

  console.log("Reading data from Excel...");
  const { profiles, worksheet, workbook } = await readExcel(
    "LinkedInProfiles.xlsx"
  );
  console.log("Starting to process profiles...");

  for (const { profileUrl, message, status, rowNumber } of profiles) {
    if (status === "Ok") {
      console.log(`Skipping ${profileUrl}, already processed.`);
      continue;
    }

    console.log(`Processing: ${profileUrl}`);
    try {
      await page.goto(profileUrl, { waitUntil: "domcontentloaded" });
      console.log(`Navigated to: ${profileUrl}`);

      const connectButton = await page
        .locator('[data-test-icon="connect-small"]')
        .nth(1); // Connect button
      const moreButton = await page
        .locator('[aria-label="More actions"]')
        .nth(1); // More button

      if ((await moreButton.count()) > 0) {
        await moreButton.click(); 
        console.log("Clicked 'More' button.");

        const dropdown = await page.locator('.artdeco-dropdown__content-inner');
        const connectFromDropdown = dropdown.locator('div[aria-label*="Invite"][role="button"]');

        if (await connectFromDropdown.count() > 0) {
          await connectFromDropdown.click();
          console.log("Clicked 'Connect' button from the dropdown.");
        } else {
          console.log("'Connect' button not found in the dropdown. Looking beside the 'More' button.");

          if ((await connectButton.count()) > 0) {
            await connectButton.click();
            console.log("Clicked 'Connect' button beside 'More'.");
          } else {
            console.error(
              "Connect button not found beside 'More' button or in dropdown."
            );
            await updateExcel(
              worksheet,
              rowNumber,
              "Failed - Connect button not found",
              workbook,
              "LinkedInProfiles.xlsx"
            );
            return; 
          }
        }
      } else {
        console.error("'More' button not found.");
        await updateExcel(
          worksheet,
          rowNumber,
          "Failed - More button not found",
          workbook,
          "LinkedInProfiles.xlsx"
        );
        return; 
      }

      // ADD NOTE FROM HERE
      const addNoteButton = await page.locator('button:has-text("Add a note")');
      if ((await addNoteButton.count()) > 0) {
        await addNoteButton.click();
        console.log("Clicked 'Add a note' button.");

        const messageBox = await page.locator('textarea[name="message"]');
        await messageBox.fill(message);
        console.log(`Filled message: ${message}`);

        const sendButton = await page.locator('button:has-text("Send")');
        await sendButton.click();
        console.log("Clicked 'Send' button.");

        await updateExcel(
          worksheet,
          rowNumber,
          "Ok",
          workbook,
          "LinkedInProfiles.xlsx"
        );
        console.log(`Connection request sent to: ${profileUrl}`);
      } else {
        console.error(`"Add a note" button not visible for ${profileUrl}`);
        await updateExcel(
          worksheet,
          rowNumber,
          "Failed - Add note not visible",
          workbook,
          "LinkedInProfiles.xlsx"
        );
      }
    } catch (error) {
      console.error(`Error processing ${profileUrl}:`, error);
      await updateExcel(
        worksheet,
        rowNumber,
        "Failed, Error occurred",
        workbook,
        "LinkedInProfiles.xlsx"
      );
    }
  }

  console.log("All profiles processed.");
  await browser.close();
})();
