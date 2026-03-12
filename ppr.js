const puppeteer = require("puppeteer");
const fs = require("fs");
const path = require("path");
const AR_EMAIL = process.env.AR_EMAIL;
const AR_PASSWORD = process.env.AR_PASSWORD;

// --------- Load client list from clients.csv ---------
function parseCsvLine(line) {
  const cells = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === "," && !inQuotes) {
      cells.push(current);
      current = "";
    } else {
      current += ch;
    }
  }
  cells.push(current);
  return cells.map(c => c.trim());
}

function loadClientsFromCsv() {
  const csvPath = path.join(__dirname, "clients.csv");
  if (!fs.existsSync(csvPath)) {
    throw new Error("clients.csv not found next to ppr.js");
  }

  const text = fs.readFileSync(csvPath, "utf8");
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  if (lines.length < 2) {
    throw new Error("clients.csv has no data rows");
  }

  const result = [];
  for (let i = 1; i < lines.length; i++) {
    const row = parseCsvLine(lines[i]);
    if (!row.length) continue;
    const [name, idStr, enabledRaw] = row;
    if (!name || !idStr) continue;
    const id = Number(idStr);
    if (!Number.isFinite(id)) continue;

    let enabled = (enabledRaw || "true").toString().toLowerCase();
    enabled = enabled === "true";

    if (!enabled) continue; // skip disabled rows
    result.push({ name, id });
  }
  return result;
}

/* =========================
   Env / Config
   ========================= */
const EMAIL = process.env.AR_EMAIL; // required

// Load all enabled clients from clients.csv
let CLIENTS = loadClientsFromCsv();

// Optional: run only one client if AR_CLIENT_ID is set (for testing)
const SINGLE_CLIENT = process.env.AR_CLIENT_ID ? Number(process.env.AR_CLIENT_ID) : null;
const hasSingleOverride = (SINGLE_CLIENT !== null && !Number.isNaN(SINGLE_CLIENT));

if (hasSingleOverride) {
  CLIENTS = CLIENTS.filter(c => c.id === SINGLE_CLIENT);
} else {
  // No single override: apply weekday scheduling logic

  // IDs for clients that should run on Monday/Wednesday/Friday only
  const MWF_IDS = new Set([
    40,   // Advance Local
    1354, // Gray Television
    1140, // McClatchy AdAuto & Custom
    1033, // McClatchy Interactive
    1075, // Sproutloud
  ]);

  const today = new Date();
  const dow = today.getDay(); // 0=Sun, 1=Mon, 2=Tue, 3=Wed, 4=Thu, 5=Fri, 6=Sat

  const isMWF = (dow === 1 || dow === 3 || dow === 5); // Mon, Wed, Fri
  const isTTh = (dow === 2 || dow === 4);              // Tue, Thu

  if (isMWF) {
    // Keep only MWF clients
    CLIENTS = CLIENTS.filter(c => MWF_IDS.has(c.id));
    console.log("Weekday group: MWF – running only MWF clients.");
  } else if (isTTh) {
    // Keep only non-MWF clients (T/Th group)
    CLIENTS = CLIENTS.filter(c => !MWF_IDS.has(c.id));
    console.log("Weekday group: TTh – running all clients except MWF group.");
  } else {
    console.log("Today is weekend (Sat/Sun) or unsupported; no clients scheduled to run.");
    CLIENTS = [];
  }
}

if (!CLIENTS.length) {
  console.error("No clients to run for today's schedule. Exiting.");
  process.exit(0); // 0 so Task Scheduler doesn't treat it as an error storm
}

let VERSION = process.env.AR_VERSION || ""; // auto-detected if empty
const PROFILE = process.env.AR_CHROME_PROFILE || "C:\\PPRChrome";
const CHROME  = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe";

const ADV_PASSWORD    = process.env.AR_PASSWORD || "";           // if Advantage fields not autofilled
const GOOGLE_PASSWORD = process.env.AR_GOOGLE_PASSWORD || "";    // if Google asks for a password
const HEADLESS        = process.env.AR_HEADLESS === "1";         // default minimized; set AR_HEADLESS=1 for true headless

const ADV   = "https://advantage.advertiserreports.com";
const SUPER = "https://superusers.advertiserreports.com";
const ROAR  = "https://roar.advertiserreports.com";

const ADV_HOME   = `${ADV}/rip/launch/home?localeCode=EN_US`;
const SUPER_HOME = `${SUPER}/rip/launch/home?localeCode=EN_US`;

const TOKEN_URL   = `${SUPER}/rip/common/json/getAccessToken`;
const CHANGE_URL  = `${SUPER}/rip/login/json/changeClient`;
const RUN_URL     = `${ROAR}/rip/roar/json/runCustomReport`;
const STATUS_URL  = `${SUPER}/rip/report/json/isGeneratingCustomReport`;

const REPORT_NAME = "com.matchcraft.rip.report.PremierPartnershipReport";

const SNAP_MODE = (process.env.AR_SNAP_MODE || "none").toLowerCase(); // "none" | "errors" | "all"

/* =========================
   Summary Table helpers
   ========================= */
function yn(v) { return v ? "Yes" : "No"; }

function padRight(str, len) {
  str = String(str ?? "");
  if (str.length >= len) return str;
  return str + " ".repeat(len - str.length);
}

function printSummaryTable(rows) {
  const headers = ["Client", "Report downloaded", "Report uploaded to drive"];

  const widths = [
    Math.max(headers[0].length, ...rows.map(r => (r.client || "").length)),
    Math.max(headers[1].length, ...rows.map(r => yn(r.downloaded).length)),
    Math.max(headers[2].length, ...rows.map(r => yn(r.uploaded).length)),
  ];

  const line = (cols) =>
    `${padRight(cols[0], widths[0])} | ${padRight(cols[1], widths[1])} | ${padRight(cols[2], widths[2])}`;

  console.log("\n==================== SUMMARY ====================");
  console.log(line(headers));
  console.log(`${"-".repeat(widths[0])}-+-${"-".repeat(widths[1])}-+-${"-".repeat(widths[2])}`);

  for (const r of rows) {
    console.log(
      line([
        r.client || "",
        yn(r.downloaded),
        yn(r.uploaded),
      ])
    );
  }
  console.log("=================================================\n");
}

/* =========================
   Screenshot Debug (context-safe)
   ========================= */
const SCREEN_DIR = path.resolve(process.cwd(), "screenshots");
if (!fs.existsSync(SCREEN_DIR)) fs.mkdirSync(SCREEN_DIR, { recursive: true });
let shotSeq = 0;

async function snap(page, label) {
  try {
    if (SNAP_MODE === "none") return;              // <- disables all screenshots
    if (!page || page.isClosed()) return;
    const num = String(++shotSeq).padStart(3, "0");
    const fname = `${num}_${String(label || "").replace(/[^\w.-]+/g, "_")}.png`;
    const full = path.join(SCREEN_DIR, fname);
    await page.screenshot({ path: full, fullPage: true }).catch(() => {});
  } catch {}
}

/* =========================
   Utilities (context-safe waits)
   ========================= */
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

async function safeWaitForSelector(page, sel, opts = {}) {
  try {
    if (!page || page.isClosed()) return null;
    return await page.waitForSelector(sel, opts);
  } catch {
    return null;
  }
}

async function safeWaitForNavigation(page, opts = {}) {
  try {
    if (!page || page.isClosed()) return null;
    await page.waitForNavigation(opts);
  } catch {}
  return page;
}

function istDates() {
  const IST_OFFSET_MS = 5.5 * 60 * 60 * 1000;
  const ONE_DAY_MS = 24 * 60 * 60 * 1000;

  const nowUtcMs = Date.now();

  // "Now" in IST, as a shifted view of UTC
  const istNow = new Date(nowUtcMs + IST_OFFSET_MS);

  // Weekday in IST: 0=Sun, 1=Mon, ..., 6=Sat
  const dow = istNow.getUTCDay();

  // Base = midnight (00:00) of *today in IST*, represented in UTC millis
  const baseMs = Date.UTC(
    istNow.getUTCFullYear(),
    istNow.getUTCMonth(),
    istNow.getUTCDate()
  );

  // Offsets from today's IST date, excluding today itself.
  let startOffset;
  let endOffset;

  if (dow === 1) {
    // Monday run covers Fri, Sat, Sun
    startOffset = -3;
    endOffset = -1;
  } else if (dow === 2) {
    // Tuesday run covers Thu, Fri, Sat, Sun, Mon
    startOffset = -5;
    endOffset = -1;
  } else if (dow === 3) {
    // Wednesday run covers Mon, Tue
    startOffset = -2;
    endOffset = -1;
  } else if (dow === 4) {
    // Thursday run covers Tue, Wed
    startOffset = -2;
    endOffset = -1;
  } else if (dow === 5) {
    // Friday run covers Wed, Thu
    startOffset = -2;
    endOffset = -1;
  } else {
    // Weekend or unexpected: default to "yesterday only"
    startOffset = -1;
    endOffset = -1;
  }

  function fmtYyyymmdd(ms) {
    const d = new Date(ms);
    const y = d.getUTCFullYear();
    const m = String(d.getUTCMonth() + 1).padStart(2, "0");
    const day = String(d.getUTCDate()).padStart(2, "0");
    return Number(`${y}${m}${day}`);
  }

  const startDate = fmtYyyymmdd(baseMs + startOffset * ONE_DAY_MS);
  const endDate = fmtYyyymmdd(baseMs + endOffset * ONE_DAY_MS);

  console.log("istDates range (IST):", {
    dow,
    startOffset,
    endOffset,
    startDate,
    endDate,
  });

  return { startDate, endDate };
}

async function ensureMCID(page) {
  const cookies = await page.cookies(SUPER); // may throw if page closed
  const names = cookies.map((c) => c.name);
  console.log("SuperUsers cookies:", names.join(", ") || "(none)");
  return cookies.some((c) => c.name === "MCID");
}

// Fetch inside page so cookies/policies apply
async function postInside(page, url, body, headers, credOpt) {
  if (!page || page.isClosed()) {
    return { status: -1, text: "Page is closed", json: null };
  }

  try {
    return await page.evaluate(
      async ({ url, body, headers, credOpt }) => {
        try {
          const r = await fetch(url, {
            method: "POST",
            headers: Object.assign(
              {
                "Content-Type": "application/json;charset=UTF-8",
                Accept: "application/json, text/plain, */*",
                "X-Requested-With": "XMLHttpRequest",
              },
              headers || {}
            ),
            body: JSON.stringify(body),
            credentials: credOpt || "include",
            referrerPolicy: "no-referrer-when-downgrade",
          });
          const text = await r.text();
          let json = null;
          try { json = JSON.parse(text); } catch {}
          return { status: r.status, text, json };
        } catch (e) {
          return { status: -1, text: String(e), json: null };
        }
      },
      { url, body, headers, credOpt }
    );
  } catch (e) {
    // This is where "Execution context was destroyed" was coming from.
    console.log("postInside outer error:", e?.message || String(e));
    return { status: -1, text: String(e), json: null };
  }
}

async function detectVersion(page) {
  // Each evaluate is wrapped to survive "execution context destroyed" from navigations
  try {
    const v1 = await page.evaluate(() => {
      const tryvals = [];
      tryvals.push(window.__APP_CONFIG__ && window.__APP_CONFIG__.version);
      tryvals.push(window.appVersion);
      tryvals.push(window.version);

      const mMeta = document.querySelector("meta[name='app:version']")?.content;
      const mData = document.querySelector("[data-version]")?.getAttribute("data-version");
      if (mMeta) tryvals.push(mMeta);
      if (mData) tryvals.push(mData);

      const re = /\b\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2}:\d{2}\s+[A-Z]{2,5}\b/;
      for (const s of Array.from(document.scripts)) {
        const t = s.textContent || "";
        const m = t.match(re);
        if (m) { tryvals.push(m[0]); break; }
      }
      const html = document.documentElement.outerHTML;
      const m2 = html.match(re);
      if (m2) tryvals.push(m2[0]);
      return tryvals.filter(Boolean)[0] || "";
    });
    if (v1) return v1;
  } catch (e) {
    console.log("detectVersion v1 evaluate failed (navigation?):", e.message);
  }

  try {
    const v2 = await page.evaluate(async (home) => {
      try {
        const r = await fetch(home, { credentials: "include" });
        const h = await r.text();
        const re = /\b\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2}:\d{2}\s+[A-Z]{2,5}\b/;
        const m = h.match(re);
        if (m) return m[0];
        const m2 = h.match(/"version"\s*:\s*"([^"]+)"/);
        return (m2 && m2[1]) || "";
      } catch { return ""; }
    }, `${SUPER}/`);
    if (v2) return v2;
  } catch (e) {
    console.log("detectVersion v2 evaluate failed (navigation?):", e.message);
  }

  try {
    const v3 = await page.evaluate(async (home) => {
      try {
        const r = await fetch(home, { credentials: "include" });
        const h = await r.text();
        const re = /\b\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2}:\d{2}\s+[A-Z]{2,5}\b/;
        const m = h.match(re);
        if (m) return m[0];
        const m2 = h.match(/"version"\s*:\s*"([^"]+)"/);
        return (m2 && m2[1]) || "";
      } catch { return ""; }
    }, SUPER_HOME);
    if (v3) return v3;
  } catch (e) {
    console.log("detectVersion v3 evaluate failed (navigation?):", e.message);
  }

  return "";
}

/* =========================
   App version auto-repair helpers
   ========================= */
async function detectVersionFresh(page) {
  // Try normal detection
  let v = await detectVersion(page);
  if (v) return v;

  // Force a network reload of SuperUsers and launch, bypassing cache
  const bust = `&_=${Date.now()}`;
  try {
    await page.goto(`${SUPER}/?nocache=1${bust}`, { waitUntil: "domcontentloaded", timeout: 60000 }).catch(()=>{});
    await page.goto(`${SUPER_HOME}&nocache=1${bust}`, { waitUntil: "domcontentloaded", timeout: 60000 }).catch(()=>{});
  } catch {}

  // Let the page settle after navigation before evaluating
  await sleep(2000);

  try {
    const v2 = await page.evaluate(async (home) => {
      try {
        const r = await fetch(home + (home.includes('?') ? '&' : '?') + '_=' + Date.now(), { credentials: "include", cache: "reload" });
        const h = await r.text();
        const re = /\b\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2}:\d{2}\s+[A-Z]{2,5}\b/;
        const m = h.match(re);
        if (m) return m[0];
        const m2 = h.match(/"version"\s*:\s*"([^"]+)"/);
        return (m2 && m2[1]) || "";
      } catch { return ""; }
    }, SUPER_HOME);
    if (v2) return v2;
  } catch (e) {
    console.log("detectVersionFresh evaluate failed (navigation?):", e.message);
  }

  // Final fallback: wait longer, try detectVersion one more time
  await sleep(3000);
  return await detectVersion(page);
}

async function forceAppRefresh(page, snapName) {
  const bust = `&_=${Date.now()}`;
  try {
    await page.goto(`${SUPER}/?refresh=1${bust}`, { waitUntil: "domcontentloaded", timeout: 60000 }).catch(()=>{});
  } catch {}
  try {
    await page.goto(`${SUPER_HOME}&refresh=1${bust}`, { waitUntil: "domcontentloaded", timeout: 60000 }).catch(()=>{});
  } catch {}
  if (snapName) { try { await snap(page, snapName); } catch {} }
}

async function changeClientRobust(page, email, clientId) {
  // 1) Decide which version to use
  let version = VERSION;

  // If we don't have a cached version yet, detect it once for this login session
  if (!version) {
    const maxTries = 5;
    const delayMs = 4000;
    for (let attempt = 1; attempt <= maxTries; attempt++) {
      version = await detectVersionFresh(page);
      console.log(`detectVersionFresh attempt ${attempt}:`, version || "(none)");
      if (version) {
        VERSION = version; // cache for rest of session
        break;
      }
      if (attempt < maxTries) {
        console.log("No version yet, waiting before retry…");
        await sleep(delayMs);
      }
    }

    if (!version) {
      console.error("changeClientRobust: could not detect app version after retries.");
      return {
        ok: false,
        version: "",
        res: {
          status: 500,
          json: { error: { message: "Could not detect app version (fresh)" } }
        }
      };
    }
  }

  console.log("changeClient(version):", version || "(none)");

  // 2) Try up to 5 attempts, auto-repair on -32003 and transient -1 errors
  for (let attempt = 1; attempt <= 5; attempt++) {
    const res = await postInside(page, CHANGE_URL, {
      currentLoginName: email,
      version,
      data: { id: clientId },
    });
    const short = (res.text || "").slice(0, 200);
    console.log(`changeClient attempt ${attempt}:`, res.status, short);
    await snap(page, `104_change_client_${res.status}_attempt${attempt}`);

    const code = res?.json?.error?.code;
    if (res.status === 200 && !code) {
      // Success, keep VERSION as-is
      return { ok: true, version, res };
    }

    if (res.status === -1) {
      console.log(`>> changeClient: transient error (context destroyed / navigation), waiting before retry…`);
      await sleep(3000);
      try {
        await page.goto(SUPER_HOME, { waitUntil: "networkidle2", timeout: 60000 }).catch(()=>{});
        await sleep(1000);
      } catch {}
      continue;
    }

    if (code === -32003) {
      console.log(">> changeClient: app version stale; forcing app refresh and re-detecting…");
      await forceAppRefresh(page, `093_super_session_refreshed_attempt${attempt}`);

      let newVersion = "";
      const maxVerTries = 3;
      const verDelayMs = 4000;
      for (let vTry = 1; vTry <= maxVerTries; vTry++) {
        newVersion = await detectVersionFresh(page);
        console.log(`detectVersionFresh (stale) attempt ${vTry}:`, newVersion || "(none)");
        if (newVersion) break;
        if (vTry < maxVerTries) {
          console.log("No version yet after refresh, waiting before retry…");
          await sleep(verDelayMs);
        }
      }

      if (newVersion) {
        version = newVersion;
        VERSION = newVersion;
        console.log(">> changeClient: re-detected version:", newVersion);
        continue;
      }

      console.error("changeClientRobust: still no version after refresh retries.");
      return {
        ok: false,
        version: "",
        res: {
          status: 500,
          json: { error: { message: "Could not detect app version after refresh" } }
        }
      };
    }

    // Some other non-retryable error
    return { ok: false, version, res };
  }

  return {
    ok: false,
    version,
    res: {
      status: 200,
      json: { error: { code: -32003, message: "Version still stale after retries" } }
    }
  };
}

/* =========================
   DOM helpers
   ========================= */
async function firstExistingSelector(page, selectors) {
  for (const sel of selectors) {
    try { if (await page.$(sel)) return sel; } catch {}
  }
  return null;
}

async function clickSubmitByHeuristics(page) {
  const primarySelectors = [
    "button[type='submit']",
    "input[type='submit']",
    "button#signIn",
    "button[name='login']",
    "button[id*='sign']",
    "button[name*='sign']",
    "button[id*='login']",
    "button[name*='login']",
  ];
  for (const sel of primarySelectors) {
    try {
      const el = await page.$(sel);
      if (el) { await el.click().catch(()=>{}); return true; }
    } catch {}
  }
  const xpaths = [
    "//button[normalize-space()='Sign in']",
    "//button[normalize-space()='Sign In']",
    "//button[contains(translate(normalize-space(.),'SIGN','sign'),'sign in')]",
    "//*[self::button or self::a][contains(translate(normalize-space(.),'SIGN','sign'),'sign in')]",
    "//*[self::button or self::input][contains(translate(normalize-space(.),'LOGIN','login'),'login')]",
  ];
  for (const xp of xpaths) {
    try {
      const [el] = await page.$x(xp);
      if (el) { await el.click().catch(()=>{}); return true; }
    } catch {}
  }
  return false;
}

async function forceType(page, selector, value) {
  try {
    const el = await page.$(selector);
    if (!el) return false;

    // Always clear ALL prefilled values
    await el.click({ clickCount: 3 }).catch(() => {});
    await page.keyboard.press("Backspace").catch(() => {});
    await page.keyboard.press("Delete").catch(() => {});

    // Now type correct value
    await el.type(value, { delay: 30 }).catch(() => {});
    return true;

  } catch {
    return false;
  }
}

/* =========================
   Navigation handoff helper
   ========================= */
async function waitForNextPageOrNav(currentPage, browser, {
  sameTabWaitUntil = "networkidle2",
  timeout = 30000,
  urlMustMatch = null,
} = {}) {
  if (!currentPage || currentPage.isClosed()) {
    const pages = await browser.pages();
    return pages[0] || null;
  }
  const startUrl = currentPage.url();
  function urlMatches(u) { return urlMustMatch ? urlMustMatch.test(u) : true; }

  const sameTabNav = currentPage
    .waitForNavigation({ waitUntil: sameTabWaitUntil, timeout })
    .then(() => currentPage)
    .catch(() => null);

  const newTarget = browser
    .waitForTarget(t => {
      const u = t.url();
      if (!u || u === "about:blank") return false;
      if (u === startUrl) return false;
      return urlMatches(u);
    }, { timeout })
    .then(t => t.page().catch(() => null))
    .catch(() => null);

  const closedThenFallback = (async () => {
    try { await currentPage.waitForEvent?.("close", { timeout }); } catch { return null; }
    const pages = await browser.pages();
    for (const p of pages) { try { if (urlMatches(p.url())) return p; } catch {} }
    return pages[0] || null;
  })();

  const winner = await Promise.race([sameTabNav, newTarget, closedThenFallback]);
  if (!winner) {
    const pages = await browser.pages();
    for (const p of pages) { const u = p.url(); if (urlMatches(u)) return p; }
    return currentPage;
  }
  return winner;
}

/* =========================
   Dialog auto-accept + modal OK sweeper
   ========================= */
async function wireDialogAutoAccept(browser) {
  const hook = async (p) => {
    try {
      p.on('dialog', async (d) => { try { await d.accept(); } catch {} });
    } catch {}
  };
  for (const p of await browser.pages()) await hook(p);
  browser.on('targetcreated', async (t) => {
    const p = await t.page().catch(() => null);
    if (p) await hook(p);
  });
}

async function clickModalOk(page, {
  labels = ["OK","Ok","Okay","Confirm","Continue","Proceed","Yes"],
  timeoutMs = 8000
} = {}) {
  if (!page || page.isClosed()) return false;
  const t0 = Date.now();

  async function tryClickInFrame(frame) {
    return frame.evaluate((labels) => {
      const norm = (s) => (s || "").replace(/\s+/g, " ").trim().toLowerCase();
      const wanted = labels.map(l => l.toLowerCase());
      const candidates = Array.from(document.querySelectorAll("button, [role='button'], input[type='button'], input[type='submit']"));
      for (const el of candidates) {
        const txt = norm(el.innerText || el.value || el.textContent || el.getAttribute('aria-label'));
        if (!txt) continue;
        if (wanted.some(w => txt === w || txt.includes(w))) {
          const rect = el.getBoundingClientRect();
          const style = window.getComputedStyle(el);
          const visible = rect.width > 1 && rect.height > 1 && style.visibility !== "hidden" && style.display !== "none";
          if (visible) { (el instanceof HTMLElement) && el.click(); return true; }
        }
      }
      const common = document.querySelector(
        ".modal-footer .btn-primary, .modal .btn-primary, .swal2-confirm, .mdc-button--raised, .btn.btn-primary"
      );
      if (common) { (common instanceof HTMLElement) && common.click(); return true; }
      return false;
    }, labels).catch(() => false);
  }

  while (Date.now() - t0 < timeoutMs) {
    if (await tryClickInFrame(page.mainFrame())) return true;
    for (const f of page.frames()) {
      if (f === page.mainFrame()) continue;
      if (await tryClickInFrame(f)) return true;
    }
    await new Promise(r => setTimeout(r, 250));
  }
  return false;
}

/* =========================
   Login flows
   ========================= */
async function handleAdvantageCredentials(page) {
  console.log(">> Advantage login form detected — using autofill");

  // Universal cross-version delay (3 seconds)
  const delay = ms => new Promise(res => setTimeout(res, ms));
  await delay(3000);   // allow browser autofill

  // Try all submit buttons
  const clickTargets = [
    "button[type='submit']",
    "input[type='submit']",
    "button#submit",
    "button:has-text('Sign In')",
    "button:has-text('Login')",
    "input[value='Sign In']"
  ];

  let clicked = false;

  for (const sel of clickTargets) {
    try {
      const el = await page.$(sel);
      if (el) {
        await el.click();
        clicked = true;
        break;
      }
    } catch {}
  }

  // Fallback — press Enter
  if (!clicked) {
    try { await page.keyboard.press("Enter"); } catch {}
  }

  // Wait for page to update
  try {
    await page.waitForLoadState("domcontentloaded", { timeout: 20000 });
  } catch {}

  console.log(">> Advantage login submitted (autofill mode)");
}

async function handleGoogleSSO(page, browser) {
  await snap(page, "010_before_google_sso");

  let clickedBtn = false;
  const btn = await safeWaitForSelector(page, "button.gsi-material-button", {
    visible: true,
    timeout: 12000,
  });
  if (btn) {
    await snap(page, "011_google_btn_found");
    try {
      await page.$eval("button.gsi-material-button", (b) => {
        const evOpts = { bubbles: true, cancelable: true, composed: true };
        b.dispatchEvent(new MouseEvent("pointerdown", evOpts));
        b.dispatchEvent(new MouseEvent("mousedown", evOpts));
        b.dispatchEvent(new MouseEvent("mouseup", evOpts));
        b.dispatchEvent(new MouseEvent("click", evOpts));
      });
      clickedBtn = true;
    } catch {}
  }
  if (!clickedBtn) {
    await snap(page, "011_google_btn_not_found_try_text");
    const viaText = await page
      .evaluate(() => {
        const norm = (s) =>
          (s || "").replace(/\s+/g, " ").trim().toLowerCase();
        const spans = Array.from(
          document.querySelectorAll(
            "span.gsi-material-button-contents, button, [role='button']"
          )
        );
        const el = spans.find((n) =>
          norm(n.textContent).includes("sign in with google")
        );
        const btn = el ? el.closest("button") || el : null;
        if (btn) {
          const evOpts = { bubbles: true, cancelable: true, composed: true };
          btn.dispatchEvent(new MouseEvent("pointerdown", evOpts));
          btn.dispatchEvent(new MouseEvent("mousedown", evOpts));
          btn.dispatchEvent(new MouseEvent("mouseup", evOpts));
          btn.dispatchEvent(new MouseEvent("click", evOpts));
          return true;
        }
        return false;
      })
      .catch(() => false);
    await snap(
      page,
      viaText ? "012_google_btn_clicked_via_text" : "012_google_btn_click_failed"
    );
  }

  await sleep(600);
  let googlePage = null;
  try {
    const target = await browser.waitForTarget(
      (t) => /accounts\.google\.com/.test(t.url()),
      { timeout: 15000 }
    );
    googlePage = await target.page();
  } catch {
    if (/accounts\.google\.com/.test(page.url())) googlePage = page;
  }
  if (!googlePage || googlePage.isClosed()) {
    await snap(page, "013_no_google_accounts_detected");
    console.log(">> Google SSO: accounts.google.com did not appear after click.");
    return;
  }
  await snap(googlePage, "014_google_accounts_page");

  let chooserShown = false;
  try {
    const chooser = await safeWaitForSelector(
      googlePage,
      "div[data-identifier], div[data-email], div[role='button'][data-identifier]",
      { timeout: 8000 }
    );
    chooserShown = !!chooser;
  } catch {
    chooserShown = false;
  }

  if (chooserShown) {
    await snap(googlePage, "015_google_chooser_ready");
    const matched = await googlePage
      .$$eval(
        "div[data-identifier], div[data-email], div[role='button'][data-identifier]",
        (nodes, email) => {
          const norm = (s) => String(s || "").trim().toLowerCase();
          const target = norm(email);
          const m = nodes.find((n) => {
            const de =
              n.getAttribute("data-email") ||
              n.getAttribute("data-identifier") ||
              n.textContent;
            return (
              norm(de) === target ||
              norm(n.textContent || "").includes(target)
            );
          });
          if (m) {
            m.click();
            return true;
          }
          return false;
        },
        EMAIL
      )
      .catch(() => false);

    if (matched) {
      const nextPage = await waitForNextPageOrNav(googlePage, browser, {
        sameTabWaitUntil: "domcontentloaded",
        timeout: 40000,
        urlMustMatch:
          /advertiserreports\.com|vendasta|superusers|advantage|roar|accounts\.google\.com/i,
      });
      if (nextPage && !nextPage.isClosed()) {
        await snap(nextPage, "016_after_tile_nav");
        if (!/accounts\.google\.com/i.test(nextPage.url())) return;
        googlePage = nextPage;
      } else {
        return;
      }
    }
  }

  console.log(">> Google SSO: using email/password fallback.");
  if (!googlePage || googlePage.isClosed()) return;

  // NOTE: typeIfEmpty isn't defined in your paste; leaving your original logic untouched.
  // If you actually rely on this path, you should implement typeIfEmpty similarly to forceType.
  try {
    if (!/accounts\.google\.com/i.test(googlePage.url())) return;

    const emailField = await safeWaitForSelector(
      googlePage,
      "input[type='email'], input#identifierId",
      { timeout: 10000 }
    );
    if (emailField) {
      const emailSel = (await googlePage
        .$(`input[type='email']`)
        .catch(() => null))
        ? "input[type='email']"
        : "input#identifierId";
      // fallback to forceType since typeIfEmpty isn't present
      await forceType(googlePage, emailSel, EMAIL);
      await snap(googlePage, "017_email_entered");
      await Promise.race([
        googlePage.click("#identifierNext").catch(() => {}),
        googlePage.click("button[type='submit']").catch(() => {}),
        googlePage.click("div[role='button']#identifierNext").catch(() => {}),
      ]);
      await safeWaitForNavigation(googlePage, {
        waitUntil: "domcontentloaded",
        timeout: 15000,
      });
      if (googlePage && !googlePage.isClosed())
        await snap(googlePage, "018_after_identifier_next");
    }
  } catch {}

  try {
    if (!googlePage || googlePage.isClosed()) return;
    if (!/accounts\.google\.com/i.test(googlePage.url())) return;

    const pwField = await safeWaitForSelector(
      googlePage,
      "input[type='password']",
      { timeout: 15000 }
    );
    if (pwField) {
      if (GOOGLE_PASSWORD) {
        await forceType(googlePage, "input[type='password']", GOOGLE_PASSWORD);
      }
      await snap(googlePage, "019_password_entered");
      await Promise.race([
        googlePage.click("#passwordNext").catch(() => {}),
        googlePage.click("button[type='submit']").catch(() => {}),
        googlePage.click("div[role='button']#passwordNext").catch(() => {}),
      ]);
      await safeWaitForNavigation(googlePage, {
        waitUntil: "networkidle2",
        timeout: 20000,
      });
      if (googlePage && !googlePage.isClosed())
        await snap(googlePage, "020_after_password_next");
    }
  } catch {}
}

async function settleHandsFreeLogin(page, browser) {
  console.log(">> Hands-free login starting…");
  await snap(page, "000_landed_on_advantage");

  const start = Date.now();
  const maxMs = 60 * 1000;

  while (Date.now() - start < maxMs) {
    const url = page.url();

    // Already inside app shell?
    if (/superusers\.advertiserreports\.com|roar\.advertiserreports\.com|advantage\.advertiserreports\.com/.test(url)) {
      const hasAppShell = await page.evaluate(() => {
        return !!(document.querySelector("app-root, app-shell, nav .user, [data-role='app-container']"));
      }).catch(() => false);
      if (hasAppShell) { await snap(page, "020_app_shell_detected"); break; }
    }

    // Look for SSO or form cues
    let seen = null;
    try {
      seen = await Promise.race([
        safeWaitForSelector(page, "button.gsi-material-button",          { timeout: 3000, visible: true }).then(()=> "google"),
        safeWaitForSelector(page, "[aria-label*='Sign in with Google']", { timeout: 3000, visible: true }).then(()=> "google"),
        safeWaitForSelector(page, "button[data-provider='google']",      { timeout: 3000, visible: true }).then(()=> "google"),
        safeWaitForSelector(page, "a[data-provider='google']",           { timeout: 3000, visible: true }).then(()=> "google"),
        safeWaitForSelector(page, "input[type='email']",                 { timeout: 3000, visible: true }).then(()=> "adv"),
        safeWaitForSelector(page, "input[name='email']",                 { timeout: 3000, visible: true }).then(()=> "adv"),
        safeWaitForSelector(page, "input[type='password']",              { timeout: 3000, visible: true }).then(()=> "adv"),
        safeWaitForSelector(page, "button[type='submit']",               { timeout: 3000, visible: true }).then(()=> "adv"),
      ]);
    } catch {}

    if (seen) {
      if (seen === "google") {
        await snap(page, "030_detected_google_option");
        await handleGoogleSSO(page, browser);
      }
      const hasAdvField = await page.$("input[type='email'], input[name='email'], input[type='password']").catch(()=>null);
      if (hasAdvField) {
        console.log(">> Detected Advantage username/password form; attempting login.");
        await handleAdvantageCredentials(page);
      }
    } else {
      await sleep(800);
      await snap(page, "031_waiting_for_login_cues");
    }
  }
}

/* =========================
   Main
   ========================= */
(async () => {
  if (!EMAIL) {
    console.error("Missing AR_EMAIL");
    process.exit(1);
  }

  const browser = await puppeteer.launch({
    headless: HEADLESS ? "new" : false,     // set AR_HEADLESS=1 for true headless; otherwise minimized
    executablePath: CHROME,
    userDataDir: PROFILE,
    defaultViewport: { width: 1400, height: 900 },
    args: [
      "--no-sandbox",
      "--disable-features=BlockThirdPartyCookies",
      "--window-size=1400,900",
      "--start-minimized",
    ],
  });

  // Auto-accept any JS alerts/confirms
  await wireDialogAutoAccept(browser);

  // Collect results for summary output
  const summaryRows = [];

  // 1) Advantage -> hands-free login
  const page = await browser.newPage();
  await page.goto(ADV_HOME, { waitUntil: "domcontentloaded" }).catch(() => {});
  await settleHandsFreeLogin(page, browser);
  // In case a post-login sheet appears
  await clickModalOk(page, { timeoutMs: 4000 });

  // 2) Poll for MCID (resilient)
  let poll = await browser.newPage();
  const maxMs = 10 * 60 * 1000;
  const start = Date.now();
  let has = false;

  while (Date.now() - start < maxMs) {
    try {
      await poll.goto(`${SUPER}/`, { waitUntil: "domcontentloaded", timeout: 60000 }).catch(() => {});
      await poll.goto(SUPER_HOME, { waitUntil: "domcontentloaded", timeout: 60000 }).catch(() => {});
      await snap(poll, "100_poll_super_home");
      has = await ensureMCID(poll);
      if (has) {
        console.log(">> MCID detected. Continuing…");
        await snap(poll, "101_mcid_detected");
        break;
      }
    } catch (e) {
      console.log("Poll error (will retry):", e.message || e);
      try { await poll.close().catch(()=>{}); } catch {}
      poll = await browser.newPage();
      await snap(poll, "099_poll_reopened");
    }
    const remain = Math.ceil((maxMs - (Date.now() - start)) / 1000);
    console.log(`Waiting for SSO to finish… (${remain}s left)`);
    await sleep(3000);
  }
  if (!has) {
    await snap(poll, "102_mcid_timeout");
    console.error("Timed out waiting for MCID.");
    await browser.close();
    process.exit(1);
  }

  // Before running any reports, dismiss any confirmation sheet already open
  await poll.bringToFront().catch(() => {});
  await clickModalOk(poll, { timeoutMs: 4000 });

  // Run reports for each client from clients.csv
  for (const client of CLIENTS) {
    console.log("==================================================");
    console.log(`Starting client ${client.name} [${client.id}]`);

    // Per-client status tracking for summary
    let reportDownloaded = false;
    let reportUploaded = false;

    // 4) changeClient with auto version repair (also gets a fresh version)
    const cc = await changeClientRobust(poll, EMAIL, client.id);
    if (!cc.ok) {
      console.error(
        `changeClient failed for clientId=${client.id}. Raw:`,
        JSON.stringify(cc.res?.json || cc.res || {}, null, 2)
      );
      await snap(poll, `104_change_client_failed_${client.id}`);

      summaryRows.push({
        client: client.name,
        downloaded: false,
        uploaded: false,
      });

      // Skip this client; continue to the next
      continue;
    }
    // VERSION is now correct for THIS client/session
    VERSION = cc.version;

    // Sweep any modal confirmations (e.g., "are you sure?" dialogs)
    await clickModalOk(poll, { timeoutMs: 4000 });

    // 5) getAccessToken using the NEW cookie/session for this client
    const tok = await postInside(
      poll,
      TOKEN_URL,
      {
        currentClientId: client.id,
        currentLoginName: EMAIL,
        version: VERSION,
        data: {},
      }
    );
    await snap(poll, `105_get_access_token_${client.id}_${tok.status}`);
    if (
      tok.status !== 200 ||
      !tok.json ||
      !tok.json.result ||
      !tok.json.result.accessToken
    ) {
      console.error(
        `getAccessToken failed for clientId=${client.id}:`,
        tok.status,
        (tok.text || "").slice(0, 400)
      );

      summaryRows.push({
        client: client.name,
        downloaded: false,
        uploaded: false,
      });

      // Don’t kill the whole job; go to next client
      continue;
    }
    const accessToken = tok.json.result.accessToken;
    console.log(
      `AccessToken OK for clientId=${client.id}. length:`,
      accessToken.length
    );

    // 6) Run report FROM ROAR origin for this client
    const { startDate, endDate } = istDates();
    const roarPage = await browser.newPage();
    try {
      await roarPage.goto(`${ROAR}/`, { waitUntil: "domcontentloaded" });
      await roarPage
        .goto(`${ROAR}/rip/launch/home?localeCode=EN_US`, {
          waitUntil: "domcontentloaded",
        })
        .catch(() => {});
      await snap(roarPage, `200_roar_home_loaded_${client.id}`);

      const runResp = await postInside(
        roarPage,
        `${RUN_URL}?export=google_sheets&access_token=${encodeURIComponent(
          accessToken
        )}`,
        {
          currentClientId: client.id,
          currentLoginName: EMAIL,
          version: VERSION,
          data: { startDate, endDate, reportName: REPORT_NAME },
        },
        {
          Origin: SUPER,
          Referer: `${SUPER}/`,
        },
        "omit"
      );
      console.log(
        `runCustomReport for clientId=${client.id}:`,
        runResp.status,
        (runResp.text || "").slice(0, 300)
      );
      await snap(
        roarPage,
        `201_run_report_${client.id}_${runResp.status}`
      );

      // After runCustomReport: call Apps Script to move only this file
      const sheetUrl = runResp?.json?.result?.url || "";
      const m = sheetUrl.match(/\/spreadsheets\/d\/([A-Za-z0-9_-]+)/);
      const fileId = m ? m[1] : null;
      console.log("Report Sheet URL:", sheetUrl);
      console.log("Report File ID:", fileId || "(none)");

      // Downloaded = we successfully got a sheet URL + ID
      reportDownloaded = (runResp.status === 200 && !!sheetUrl && !!fileId);

      const APPS_SCRIPT_WEBHOOK = process.env.AR_APPS_SCRIPT_URL || "";
      if (fileId && APPS_SCRIPT_WEBHOOK) {
        try {
          const hookUrl = `${APPS_SCRIPT_WEBHOOK}${
            APPS_SCRIPT_WEBHOOK.includes("?") ? "&" : "?"
          }fileId=${encodeURIComponent(fileId)}&clientId=${encodeURIComponent(
            String(client.id)
          )}&ts=${Date.now()}`;

          const hook = await browser.newPage();
          const resp = await hook.goto(hookUrl, {
            waitUntil: "domcontentloaded",
            timeout: 20000,
          });
          await snap(
            hook,
            `210_move_webhook_loaded_${client.id}`
          );

          // Some Apps Script web apps legitimately return an empty body.
          // So consider "uploaded" if:
          // - reportDownloaded is true
          // - hook navigation returned HTTP 2xx (or at least ok())
          // - no exception thrown
          const status = resp ? resp.status() : 0;
          const ok = resp ? resp.ok() : false;

          const msg = await hook.evaluate(
            () => ((document.body && document.body.innerText) || "").slice(0, 300)
          ).catch(() => "");

          console.log("Move webhook HTTP status:", status || "(unknown)");
          console.log("Move webhook page says:", msg || "(no body)");

          // Primary success signal = HTTP ok
          reportUploaded = reportDownloaded && (ok || (status >= 200 && status < 300));

          await hook.close().catch(() => {});
        } catch (e) {
          console.log(
            `Move webhook navigation failed for clientId=${client.id} (non-fatal):`,
            e?.message || e
          );
          reportUploaded = false;
        }
      } else {
        console.log(
          "Skipping move webhook (no fileId or AR_APPS_SCRIPT_URL not set)."
        );
        reportUploaded = false;
      }
    } finally {
      await roarPage.close().catch(() => {});
    }

    // 7) Optional status from SuperUsers tab for this client
    const st = await postInside(
      poll,
      STATUS_URL,
      {
        currentClientId: client.id,
        currentLoginName: EMAIL,
        version: VERSION,
        data: { reportClassName: REPORT_NAME, userEmail: EMAIL },
      }
    );
    console.log(
      `status for clientId=${client.id}:`,
      st.status,
      (st.text || "").slice(0, 220)
    );
    await snap(poll, `202_status_${client.id}_${st.status}`);

    // Record summary row
    summaryRows.push({
      client: client.name,
      downloaded: reportDownloaded,
      uploaded: reportUploaded,
    });
  }

  // After all clients are processed
  printSummaryTable(summaryRows);
  await browser.close();
})().catch(async (e) => {

  console.error(e);
  try {
    const fallback = path.join(SCREEN_DIR, "zzz_unhandled_error.txt");
    fs.writeFileSync(fallback, (e && e.stack) ? e.stack : String(e));
    console.log("Saved error stack:", fallback);
  } catch {}
  process.exit(1);
});
