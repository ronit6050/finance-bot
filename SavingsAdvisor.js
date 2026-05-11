/* ============================================
   SAVINGS & INVESTMENT ADVISOR
   /save — log a saving
   /invest — log an investment
   /wishlist — manage wish list
   Dashboard button — full overview
   Sunday reminder — weekly nudge
============================================ */

const MONTHLY_SAVE_GOAL = 1000; // ₹1,000 fixed monthly goal


/* =======================================
   SAVINGS DASHBOARD
   Shows savings + wish list progress
======================================= */
function sendSavingsDashboard() {

  try {

    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const savSheet    = ss.getSheetByName("Savings");
    const wishSheet   = ss.getSheetByName("WishList");
    const investSheet = ss.getSheetByName("Investments");

    if (!savSheet)    { sendMessage("❌ Savings sheet not found");    return; }
    if (!wishSheet)   { sendMessage("❌ WishList sheet not found");   return; }
    if (!investSheet) { sendMessage("❌ Investments sheet not found"); return; }

    const fmt = (n) => Math.round(n).toLocaleString('en-IN');

    // ── Total savings ──
    const savData    = savSheet.getDataRange().getValues();
    let totalSavings = 0;
    let thisMonthSav = 0;
    let lastMonthSav = 0;

    const today     = new Date();
    const thisMonth = today.getMonth();
    const thisYear  = today.getFullYear();
    const lastMonth = thisMonth === 0 ? 11 : thisMonth - 1;
    const lastYear  = thisMonth === 0 ? thisYear - 1 : thisYear;

    for (let i = 1; i < savData.length; i++) {
      const rawDate = savData[i][0];
      const amount  = Number(savData[i][1]) || 0;
      if (!rawDate) continue;
      const d = new Date(rawDate);
      totalSavings += amount;
      if (d.getFullYear() === thisYear  && d.getMonth() === thisMonth) thisMonthSav += amount;
      if (d.getFullYear() === lastYear  && d.getMonth() === lastMonth) lastMonthSav += amount;
    }

    // ── Monthly average savings ──
    const months    = getMonthlySavingsHistory(savData);
    const avgSaving = months.length > 0
      ? Math.round(months.reduce((a, b) => a + b, 0) / months.length)
      : MONTHLY_SAVE_GOAL;

    // ── Wish list progress ──
    const wishData = wishSheet.getDataRange().getValues();
    let wishText   = "";
    let soonest    = null;

    for (let i = 1; i < wishData.length; i++) {
      const item     = (wishData[i][0] || "").toString().trim();
      const price    = Number(wishData[i][1]) || 0;
      const priority = (wishData[i][2] || "Medium").toString().trim();
      const status   = (wishData[i][3] || "Active").toString().trim();

      if (!item || status !== "Active") continue;

      const remaining    = Math.max(price - totalSavings, 0);
      const monthsNeeded = avgSaving > 0
        ? Math.ceil(remaining / avgSaving)
        : "∞";

      const canAfford = totalSavings >= price;

      let line = `\n• ${item} — ₹${fmt(price)} [${priority}]`;

      if (canAfford) {
        line += `\n  ✅ You can afford this now!`;
      } else {
        line += `\n  Saved: ₹${fmt(totalSavings)} | Still need: ₹${fmt(remaining)}`;
        line += `\n  At ₹${fmt(avgSaving)}/month → ${monthsNeeded} months away`;
      }

      wishText += line;

      // Track soonest affordable item
      if (typeof monthsNeeded === "number") {
        if (!soonest || monthsNeeded < soonest.months) {
          soonest = { item, months: monthsNeeded, price, remaining };
        }
      }
    }

    // ── This month vs last month ──
    const savingsTrend = thisMonthSav >= lastMonthSav
      ? `📈 Up from ₹${fmt(lastMonthSav)} last month`
      : `📉 Down from ₹${fmt(lastMonthSav)} last month`;

    // ── Goal progress this month ──
    const goalPct  = Math.min(Math.round((thisMonthSav / MONTHLY_SAVE_GOAL) * 100), 100);
    const goalBar  = buildSavingsBar(goalPct);

    const message =
`💰 Savings Dashboard

📊 This month: ₹${fmt(thisMonthSav)} / ₹${fmt(MONTHLY_SAVE_GOAL)} goal
${goalBar}
${savingsTrend}

💵 Total saved: ₹${fmt(totalSavings)}
📅 Avg monthly saving: ₹${fmt(avgSaving)}

🎯 Wish List:${wishText || "\nNo active items — add one with /wishlist"}`;

    sendMessage(message);

    // ── Soonest affordable item highlight ──
    if (soonest && soonest.months <= 6) {
      sendMessage(
`🏆 Closest goal: ${soonest.item}

You need ₹${fmt(soonest.remaining)} more.
At your current pace — just ${soonest.months} month${soonest.months === 1 ? "" : "s"} away!

Keep it up! 💪`
      );
    }

    // ── AI suggestion ──
    sendSavingsAITip(totalSavings, avgSaving, thisMonthSav, lastMonthSav, wishData);

  } catch (err) {
    logAI("SAVINGS_DASHBOARD_ERROR", err.toString());
    sendMessage("❌ Error: " + err.message);
  }
}


/* =======================================
   INVESTMENT DASHBOARD
   Shows current investments breakdown
======================================= */
function sendInvestmentDashboard() {

  try {

    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const investSheet = ss.getSheetByName("Investments");

    if (!investSheet) { sendMessage("❌ Investments sheet not found"); return; }

    const data = investSheet.getDataRange().getValues();
    const fmt  = (n) => Math.round(n).toLocaleString('en-IN');

    // ── Group by type ──
    let typeTotals = {};
    let total      = 0;

    for (let i = 1; i < data.length; i++) {
      const rawDate = data[i][0];
      const type    = (data[i][1] || "Other").toString().trim();
      const amount  = Number(data[i][2]) || 0;
      if (!rawDate || !amount) continue;
      typeTotals[type] = (typeTotals[type] || 0) + amount;
      total += amount;
    }

    // ── Build breakdown ──
    let breakdown = "";
    Object.entries(typeTotals)
      .sort((a, b) => b[1] - a[1])
      .forEach(([type, amt]) => {
        const pct = total > 0 ? ((amt / total) * 100).toFixed(1) : 0;
        breakdown += `\n• ${type}: ₹${fmt(amt)} (${pct}%)`;
      });

    // ── Recent entries ──
    const recent = data.slice(1)
      .filter(r => r[0] && r[2])
      .sort((a, b) => new Date(b[0]) - new Date(a[0]))
      .slice(0, 3)
      .map(r => {
        const d    = new Date(r[0]);
        const date = `${d.getDate()} ${d.toLocaleString('default', {month:'short'})}`;
        return `• ${date} — ${r[1]}: ₹${fmt(Number(r[2]))}${r[3] ? " (" + r[3] + ")" : ""}`;
      })
      .join("\n");

    const message =
`📈 Investment Dashboard

💰 Total invested: ₹${fmt(total)}

📊 Breakdown:${breakdown || "\nNo investments yet"}

🕐 Recent entries:
${recent || "None yet"}

💡 To update a value, use:
/invest 500 mutual fund update`;

    sendMessage(message);

  } catch (err) {
    logAI("INVEST_DASHBOARD_ERROR", err.toString());
    sendMessage("❌ Error: " + err.message);
  }
}


/* =======================================
   LOG A SAVING
   /save 1000 bank
   /save 500 mutual fund
======================================= */
function logSaving(text) {

  try {

    const parts  = text.trim().split(" ");
    const amount = Number(parts[0]);

    if (!amount || amount <= 0) {
      sendMessage(
`❌ Invalid format.

Correct format:
Amount Type Note

Examples:
1000 bank monthly goal
500 mutual fund SIP
200 cash leftover`
      );
      return;
    }

    // ── Parse type and note ──
    // Second word = type, rest = note
    const type = parts[1] || "General";
    const note = parts.slice(2).join(" ") || "";

    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const savSheet = ss.getSheetByName("Savings");

    if (!savSheet) { sendMessage("❌ Savings sheet not found"); return; }

    const today = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"
    );

    savSheet.appendRow([today, amount, type, note]);

    const fmt = (n) => Math.round(n).toLocaleString('en-IN');

    // ── Get new total ──
    const data = savSheet.getDataRange().getValues();
    let total  = 0;
    for (let i = 1; i < data.length; i++) {
      total += Number(data[i][1]) || 0;
    }

    // ── Check wish list affordability ──
    checkWishListAffordability(total);

    sendMessage(
`✅ Saving logged!

💵 Amount: ₹${fmt(amount)}
🏦 Type: ${type}
📝 Note: ${note || "—"}
💰 Total savings: ₹${fmt(total)}`
    );

  } catch (err) {
    logAI("LOG_SAVING_ERROR", err.toString());
    sendMessage("❌ Error: " + err.message);
  }
}


/* =======================================
   LOG AN INVESTMENT
   /invest 500 mutual fund
   /invest 200 gold
======================================= */
function logInvestment(text) {

  try {

    const parts  = text.trim().split(" ");
    const amount = Number(parts[0]);

    if (!amount || amount <= 0) {
      sendMessage("❌ Invalid format.\n\nTry:\n500 Mutual Fund\n200 Gold\n1000 Stocks");
      return;
    }

    const type = parts.slice(1).join(" ") || "General";

    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const investSheet = ss.getSheetByName("Investments");

    if (!investSheet) { sendMessage("❌ Investments sheet not found"); return; }

    const today = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"
    );

    investSheet.appendRow([today, type, amount, ""]);

    const fmt = (n) => Math.round(n).toLocaleString('en-IN');

    // ── Get new total for this type ──
    const data      = investSheet.getDataRange().getValues();
    let typeTotal   = 0;
    let grandTotal  = 0;

    for (let i = 1; i < data.length; i++) {
      const rowType = (data[i][1] || "").toString().trim().toLowerCase();
      const rowAmt  = Number(data[i][2]) || 0;
      grandTotal   += rowAmt;
      if (rowType === type.toLowerCase()) typeTotal += rowAmt;
    }

    sendMessage(
`✅ Investment logged!

📈 Type: ${type}
💵 Amount added: ₹${fmt(amount)}
📊 Total in ${type}: ₹${fmt(typeTotal)}
💰 Total invested: ₹${fmt(grandTotal)}`
    );

  } catch (err) {
    logAI("LOG_INVEST_ERROR", err.toString());
    sendMessage("❌ Error: " + err.message);
  }
}


/* =======================================
   ADD TO WISH LIST
   /wishlist iPhone 16 80000 high
======================================= */
function addWishListItem(text) {

  try {

    const { GEMINI_KEY } = getConfig();

    const prompt = `
Extract wish list item from this text written by an Indian user: "${text}"

The text contains: item name, price (a number), and optionally a priority (high/medium/low).

Return ONLY a JSON object:
{
  "item": "full item name",
  "price": number,
  "priority": "High or Medium or Low"
}

Rules:
- item: the full product or experience name (can be multiple words)
- price: extract the number — it is always a standalone number in the text
- priority: if mentioned use it, otherwise default to "Medium"
- Return ONLY JSON. No explanation, no markdown.

Examples:
"Body Massager 2000 High" → {"item":"Body Massager","price":2000,"priority":"High"}
"iPhone 16 80000" → {"item":"iPhone 16","price":80000,"priority":"Medium"}
"Goa trip 25000 medium" → {"item":"Goa Trip","price":25000,"priority":"Medium"}
"new shoes 3000 low" → {"item":"New Shoes","price":3000,"priority":"Low"}
`;

    const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent?key=" + GEMINI_KEY;

    const response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { maxOutputTokens: 100 }
      }),
      muteHttpExceptions: true
    });

    const json    = JSON.parse(response.getContentText());
    let rawText   = json?.candidates?.[0]?.content?.parts?.[0]?.text || "";
    rawText       = rawText.replace(/```json|```/g, "").trim();

    let parsed;
    try{
      parsed = JSON.parse(rawText);
    }catch(e){
      sendMessage("❌ Couldn't parse that.\n\nTry:\nBody Massager 2000 high\niPhone 16 80000\nGoa trip 25000 medium");
      return;
    }

    if (!parsed.item || !parsed.price) {
      sendMessage("❌ Couldn't find item name or price.\n\nTry:\nBody Massager 2000 high\niPhone 16 80000");
      return;
    }

    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const wishSheet = ss.getSheetByName("WishList");

    if (!wishSheet) { sendMessage("❌ WishList sheet not found"); return; }

    const today = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"
    );

    wishSheet.appendRow([
      parsed.item,
      parsed.price,
      parsed.priority,
      "Active",
      today
    ]);

    const fmt = (n) => Math.round(n).toLocaleString('en-IN');

    sendMessage(
`✅ Added to Wish List!

🎯 ${parsed.item}
💰 Price: ₹${fmt(parsed.price)}
⭐ Priority: ${parsed.priority}

Check Savings Dashboard to see how long until you can afford it.`
    );

  } catch (err) {
    logAI("WISHLIST_ERROR", err.toString());
    sendMessage("❌ Error: " + err.message);
  }
}


/* =======================================
   CHECK IF WISH LIST ITEM NOW AFFORDABLE
   Called after every saving log
======================================= */
function checkWishListAffordability(totalSavings) {

  try {

    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const wishSheet = ss.getSheetByName("WishList");
    if (!wishSheet) return;

    const data = wishSheet.getDataRange().getValues();
    const fmt  = (n) => Math.round(n).toLocaleString('en-IN');

    for (let i = 1; i < data.length; i++) {
      const item   = (data[i][0] || "").toString().trim();
      const price  = Number(data[i][1]) || 0;
      const status = (data[i][3] || "Active").toString().trim();

      if (!item || status !== "Active") continue;

      // Check if just became affordable
      if (totalSavings >= price) {
        sendMessage(
`🎉 Wish List Alert!

You can now afford: ${item}
Price: ₹${fmt(price)}
Your savings: ₹${fmt(totalSavings)}

Time to treat yourself? 🛍️`
        );
      }
    }

  } catch (err) {
    logAI("WISHLIST_AFFORD_ERROR", err.toString());
  }
}


/* =======================================
   GEMINI SAVINGS TIP
======================================= */
function sendSavingsAITip(total, avgSaving, thisMonth, lastMonth, wishData) {

  try {

    const { GEMINI_KEY } = getConfig();
    const fmt = (n) => Math.round(n).toLocaleString('en-IN');

    // Build wish list summary for Gemini
    const wishSummary = wishData.slice(1)
      .filter(r => r[0] && r[3] === "Active")
      .map(r => `${r[0]}: ₹${fmt(Number(r[1]))} [${r[2]}]`)
      .join(", ") || "No items yet";

    const prompt = `
You are a savings advisor for an Indian user.

Their savings data:
- Total saved so far: ₹${fmt(total)}
- Monthly saving goal: ₹${fmt(MONTHLY_SAVE_GOAL)}
- This month's saving: ₹${fmt(thisMonth)}
- Last month's saving: ₹${fmt(lastMonth)}
- Average monthly saving: ₹${fmt(avgSaving)}
- Wish list items: ${wishSummary}

Write 2-3 sentences of specific, actionable savings advice:
- Which wish list item should they prioritise and why?
- If avg saving is low, suggest a realistic increase amount
- If they saved more this month, acknowledge it
- Be specific — use real numbers and item names
- Use ₹ and Indian formatting
- No generic advice like "spend less"
- End with one motivating line
- Write ONLY the advice. No labels.
`;

    const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + GEMINI_KEY;

    const response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
      muteHttpExceptions: true
    });

    const json = JSON.parse(response.getContentText());
    const tip  = json?.candidates?.[0]?.content?.parts?.[0]?.text;

    if (tip && tip.trim()) {
      sendMessage("💡 " + tip.trim());
    }

  } catch (err) {
    logAI("SAVINGS_TIP_ERROR", err.toString());
  }
}


/* =======================================
   SUNDAY SAVINGS REMINDER
   Runs every Sunday via trigger
======================================= */
function sendSundaySavingsReminder() {

  try {

    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const savSheet = ss.getSheetByName("Savings");
    if (!savSheet) return;

    const data      = savSheet.getDataRange().getValues();
    const today     = new Date();
    const thisMonth = today.getMonth();
    const thisYear  = today.getFullYear();
    const fmt       = (n) => Math.round(n).toLocaleString('en-IN');

    let thisMonthSav = 0;
    let total        = 0;

    for (let i = 1; i < data.length; i++) {
      const rawDate = data[i][0];
      const amount  = Number(data[i][1]) || 0;
      if (!rawDate) continue;
      const d = new Date(rawDate);
      total += amount;
      if (d.getFullYear() === thisYear && d.getMonth() === thisMonth) {
        thisMonthSav += amount;
      }
    }

    const remaining = Math.max(MONTHLY_SAVE_GOAL - thisMonthSav, 0);
    const onTrack   = thisMonthSav >= MONTHLY_SAVE_GOAL;

    let message = `💰 Weekly Savings Check-in\n\n`;

    if (onTrack) {
      message += `✅ Monthly goal hit! ₹${fmt(thisMonthSav)} saved this month.\n`;
      message += `💵 Total savings: ₹${fmt(total)}\n\n`;
      message += `Any extra to add? Use /save amount type`;
    } else {
      message += `📊 This month: ₹${fmt(thisMonthSav)} of ₹${fmt(MONTHLY_SAVE_GOAL)} goal\n`;
      message += `⏳ Still need: ₹${fmt(remaining)} to hit goal\n`;
      message += `💵 Total savings: ₹${fmt(total)}\n\n`;
      message += `Log a saving with /save amount type\nExample: /save 500 bank`;
    }

    sendMessage(message);

  } catch (err) {
    logAI("SUNDAY_REMINDER_ERROR", err.toString());
  }
}


/* =======================================
   HELPER: Get monthly savings history
   Returns array of monthly totals
======================================= */
function getMonthlySavingsHistory(data) {

  const monthly = {};

  for (let i = 1; i < data.length; i++) {
    const rawDate = data[i][0];
    const amount  = Number(data[i][1]) || 0;
    if (!rawDate) continue;
    const d   = new Date(rawDate);
    const key = `${d.getFullYear()}-${d.getMonth()}`;
    monthly[key] = (monthly[key] || 0) + amount;
  }

  return Object.values(monthly);
}


/* =======================================
   PROGRESS BAR
======================================= */
function buildSavingsBar(pct) {
  const filled = Math.min(Math.round(pct / 10), 10);
  const empty  = 10 - filled;
  return "▓".repeat(filled) + "░".repeat(empty) + ` ${pct}%`;
}