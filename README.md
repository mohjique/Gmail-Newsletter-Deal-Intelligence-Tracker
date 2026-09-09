# VC Newsletter Intelligence Pipeline

A Google Apps Script that reads your VC and startup newsletter emails, uses Google Gemini AI to extract structured deal data, and logs everything into a Google Sheet — automatically, every morning.

**What it captures:** VC funding rounds, M&A deals, IPOs, notable company moves, layoffs, and leadership changes.

**Cost:** Free. Uses Gemini 2.0 Flash via Google AI Studio's free tier (1,500 requests/day — far beyond any newsletter volume).

---

## How it works

1. Each morning, Gmail is queried for emails from your configured newsletter senders
2. Each email is cleaned, compressed, and checked for deal-relevant keywords
3. Emails with no relevant signals are dropped before touching the AI
4. Batches of 10 emails are sent to Gemini 2.0 Flash for structured extraction
5. Each extracted deal is deduplicated via MD5 hash and written to Google Sheets
6. Processed thread IDs are cached so nothing is ever scanned twice

---

## What you need

- A Google account
- A Gmail inbox with newsletter subscriptions
- A free Gemini API key from [Google AI Studio](https://aistudio.google.com/app/apikey)

---

## Step 1 — Copy the Sheet template

Open the [Deal Intelligence Tracker template](https://docs.google.com/spreadsheets/d/1IUqxJYCTUpwSPpmHu-OKGRLTpM1HfwyXE2JgqrAD6C4/edit?usp=sharing) and click **File → Make a copy**. This gives you your own copy with the correct tabs and column headers already set up.

Once copied, grab your **Sheet ID** from the URL — it is the long string between `/d/` and `/edit`:

```
https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID_HERE/edit
```

You will need this later.

> **Do not delete column K (Dedup Hash).** The script uses it to prevent duplicate rows. It can be hidden but must stay in place.

---

## Step 2 — Create your Apps Script project

1. Go to [script.google.com](https://script.google.com) and click **New project**
2. Delete the placeholder code in the editor
3. Paste the full contents of `pipeline.gs` from this repo
4. Click **Save** (Ctrl+S / Cmd+S)

---

## Step 3 — Enable the Gmail advanced service

The script uses Gmail's advanced API for efficient inbox search. Without this it will not run.

1. In the left sidebar click **Services** (the + icon)
2. Find **Gmail API** in the list
3. Click **Add**

---

## Step 4 — Configure your newsletter senders

In the `CONFIG` block near the top of the script, update `NEWSLETTER_QUERY` to match your actual newsletter senders:

```javascript
NEWSLETTER_QUERY: [
  'from:your-newsletter@example.com',
  'from:another@newsletter.com',
].join(' OR '),
```

Each entry must be in the format `from:email@domain.com`. Add as many senders as you like.

---

## Step 5 — Set Script Properties

Script Properties store your API key and Sheet ID securely — they are never written into the code itself.

1. In Apps Script, click the **gear icon** (Project Settings) in the left sidebar
2. Scroll down to **Script Properties**
3. Click **Add property** and add both of the following:

| Property name | Value |
|---------------|-------|
| `GEMINI_API_KEY` | Your API key from [aistudio.google.com](https://aistudio.google.com/app/apikey) |
| `SHEET_ID` | The Sheet ID you copied in Step 1 |

> Property names are **case-sensitive** — copy them exactly as shown above.

---

## Step 6 — Verify the setup

Select **`showConfig`** from the function dropdown and click **Run**.

Google will ask you to authorize access to Gmail and Google Sheets — approve everything. If you see a warning that the app is unverified, click **Advanced** then **Go to [project name] (unsafe)** — this is normal for personal scripts that have not been submitted to Google for review.

Check the Execution log to confirm your sender list looks correct.

---

## Step 7 — Activate daily automation

Select **`installDailyTrigger`** from the function dropdown and click **Run**.

This installs a time-driven trigger that runs `dailyScan()` every morning at 7am. From this point on the pipeline is fully automatic — new newsletter emails will be processed and written to your sheet each day without any manual action.

To turn off automation at any time, run **`removeDailyTrigger`**.

---

## Sheet structure

| Column | Field | Notes |
|--------|-------|-------|
| A | Date | Date of the email |
| B | Category | vc_funding / ma / ipo / notable / layoffs / leadership |
| C | Company | Primary company name |
| D | Deal Type | Fixed list: Seed / Series A–E+ / Growth Round / Acquisition / Merger / IPO / Funding Round / Layoff / Leadership Change / Other |
| E | Amount (USD, millions) | Plain number, no $ sign — e.g. `500` = $500M, `1300` = $1.3B, `0.5` = $500K; blank if not mentioned |
| F | Investors / Acquirer / Parties | External parties only |
| G | Stage | Seed / Series A–E+ / Growth / Public / blank |
| H | Sector | Fixed list — AI/ML, Fintech, Biotech, Healthcare, Cybersecurity, Defense, Robotics, Semiconductors, Energy & Climate, Space & Aerospace, SaaS / Enterprise Software, Crypto / Web3, Logistics & Supply Chain, Media & Entertainment, Quantum Computing, AgTech, Deep Tech / Hardware, Real Estate & PropTech, EdTech, Legal Tech, Data Infrastructure, Automotive & Mobility, Consumer & Retail, GovTech, Venture Capital, Telecommunications, Technology, Other |
| I | Summary | One sentence, max 20 words |
| J | Source Newsletter | Fixed list: StrictlyVC / TechCrunch / Newcomer / Cheddar / Opening Bell / Axios / Tom Tunguz / VC Deals |
| K | Dedup Hash | **Do not delete** — used for deduplication |

---

## Troubleshooting

**No rows appearing after the first run**
- Check the Execution log for error messages
- Confirm `SHEET_ID` and `GEMINI_API_KEY` are set in Script Properties (names are case-sensitive)
- Make sure the Gmail API advanced service is enabled (Step 3)

**`Invalid argument: id` error**
- `SHEET_ID` is missing or misspelled in Script Properties

**Gemini API error 400 or 403**
- Confirm your API key is valid at [aistudio.google.com](https://aistudio.google.com)
- Make sure the Gemini API is enabled for your key

**Authorization error on first run**
- Click through the Google authorization flow — you need to grant access to both Gmail and Google Sheets
- If prompted about an unverified app, click Advanced > Go to [project name] (unsafe)

**Duplicate rows appearing**
- Do not delete or move column K — the dedup check depends on it being in position 11

---

## Repo structure

```
├── pipeline.gs   # Full Apps Script source
└── README.md     # This file
```

---

## Extending the pipeline

**Add more newsletter senders** — update `NEWSLETTER_QUERY` in CONFIG and re-save the script.

**Change the scan window** — `dailyScan()` looks back 1 day by default. Change `yesterday.setDate(yesterday.getDate() - 1)` to a larger number to catch up if the trigger missed a day.

**Add more deal categories** — update the `CATEGORIES` array in CONFIG and adjust the extraction prompt in `callGeminiForDeals()` accordingly.
