import { chromium } from 'playwright-core';

async function checkRemote() {
  let browser;
  try {
    browser = await chromium.connectOverCDP('http://127.0.0.1:9222');
    const context = browser.contexts()[0];
    if (!context) {
      console.log('Connected to Edge, but no context found.');
      return;
    }

    const pages = context.pages();
    console.log(`Connected to Edge on port 9222. Tabs: ${pages.length}`);

    for (let i = 0; i < pages.length; i++) {
      const page = pages[i];
      const title = await page.title();
      const url = page.url();
      console.log(`[Tab ${i + 1}] ${title} (${url})`);

      if (url.includes('x.com') || url.includes('twitter.com')) {
        const cookies = await context.cookies();
        const loggedIn = cookies.some(c => c.name === 'auth_token');
        console.log(`Status: ${loggedIn ? 'Logged in' : 'Not logged in'}`);
      }
    }
  } catch (err) {
    if (err.message.includes('ECONNREFUSED')) {
      console.log('Edge is not running on port 9222. Run `node launch.js` first.');
    } else {
      console.error(`Error: ${err.message}`);
    }
  } finally {
    if (browser) {
      await browser.close();
    }
  }
}

checkRemote();
