// move-to-shared-drive.js
const puppeteer = require("puppeteer");

function arg(name, def = "") {
  const a = process.argv.find(x => x.startsWith(`--${name}=`));
  return a ? a.slice(name.length + 3) : def;
}

const sheetUrl  = arg("url");
const profile   = arg("profile", "C:\\PPRChrome");
const driveName = arg("drive", "PPR Incoming Folder");
const pathStr   = arg("path", "Testing");

if (!sheetUrl) {
  console.error('Usage: node move-to-shared-drive.js --url="<sheet url>" --profile="C:\\PPRChrome" --drive="PPR Incoming Folder" --path="Testing"');
  process.exit(1);
}

(async () => {
  const browser = await puppeteer.launch({
    headless: false,
    executablePath: "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
    userDataDir: profile,
    defaultViewport: { width: 1400, height: 900 },
    args: ["--no-sandbox","--disable-features=BlockThirdPartyCookies","--window-size=1400,900"]
  });

  const page = await browser.newPage();
  await page.goto(sheetUrl, { waitUntil: "domcontentloaded", timeout: 60000 });

  // Wait for Sheets app shell, then open File menu
  await page.waitForSelector('div[aria-label="File"]', { timeout: 30000 });
  await page.click('div[aria-label="File"]');
  // Wait for menu and click "Move"
  await page.waitForSelector('div[role="menuitem"][aria-label^="Move"]', { timeout: 15000 });
  await page.click('div[role="menuitem"][aria-label^="Move"]');

  // Wait for Move dialog
  await page.waitForSelector('div[role="dialog"][aria-label*="Move"]', { timeout: 20000 });

  // Helper to click an item in the left tree or list by its aria-label/text
  async function clickItemByLabel(label) {
    const ok = await page.evaluate((lbl) => {
      const norm = s => (s||"").replace(/\s+/g," ").trim().toLowerCase();
      const L = norm(lbl);
      const all = Array.from(document.querySelectorAll('div[role="dialog"] [role="treeitem"], div[role="dialog"] [aria-label], div[role="dialog"] [role="option"]'));
      const hit = all.find(n => {
        const t = norm(n.getAttribute("aria-label") || n.textContent);
        return t === L || t.includes(L);
      });
      if (hit) { hit.click(); return true; }
      return false;
    }, label);
    return ok;
  }

  // Select "Shared drives"
  await clickItemByLabel("Shared drives");

  // Select the specific Shared Drive
  await page.waitForTimeout(400);
  await clickItemByLabel(driveName);

  // Walk into the nested folder path
  const parts = pathStr.split("/").map(s => s.trim()).filter(Boolean);
  for (const part of parts) {
    await page.waitForTimeout(300);
    await clickItemByLabel(part);
  }

  // Click "Move here" (or "Move")
  const moved = await page.evaluate(() => {
    const norm = s => (s||"").trim().toLowerCase();
    const labels = ["move here", "move"];
    const btns = Array.from(document.querySelectorAll('div[role="dialog"] button, div[role="dialog"] [role="button"]'));
    const b = btns.find(x => labels.some(l => (norm(x.textContent || x.getAttribute("aria-label")) || "").includes(l)));
    if (b) { b.click(); return true; }
    return false;
  });

  if (!moved) {
    console.error("Failed to click Move here/Move button.");
  }

  await page.waitForTimeout(1200);
  await browser.close();
})().catch(e => {
  console.error(e);
  process.exit(1);
});
