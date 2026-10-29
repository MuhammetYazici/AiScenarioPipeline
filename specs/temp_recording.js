const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({
    headless: false
  });
  const context = await browser.newContext();
  const page = await context.newPage();
  await page.goto('https://aistager.ai/tr/register');
  await page.locator('.absolute.inset-0.bg-gradient-to-t').click();
  await page.close();

  // ---------------------
  await context.close();
  await browser.close();
})();