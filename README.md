# HK Gold Price - web version

A small website that shows the client's "HK" gold price. It reads the USD-&gt;PHP rate
from XE and the gold Bid from Kitco, applies the same formula as the Windows program,
and recalculates every 15 minutes while the page is open (plus a "Calculate now" button).

It works in any browser, including Safari on an iPhone - no install, no sideloading.

## How it is built

- `index.html` - the page itself (what the client sees). Plain HTML/JavaScript, no libraries.
- `api/price.py` - a tiny server function that reads XE + Kitco and returns the finished
  numbers as JSON. A browser cannot read those sites directly (they block cross-site
  requests), so this does it on the server.
- `api/core.py` - the exact same logic file as the Windows program (formula, parsing,
  rounding, the client's text format). One source of truth.
- `requirements.txt` - lists `tzdata` so time zones work on the server.

History is saved **in the browser** (per device), not on the server. The client's phone
builds up its own history as they use it. "Export CSV" downloads it with the same columns
as the desktop `records.csv`.

## Deploy it for free (Vercel, from GitHub)

Vercel's free "Hobby" plan runs both the page and the small Python function, and it
deploys straight from a GitHub repo.

1. **Put these files in a GitHub repo.** Create a new repo (e.g. `hk-gold-web`), and add:
   ```
   index.html
   requirements.txt
   api/price.py
   api/core.py
   ```
   Keep the `api/` folder exactly as-is.
2. **Sign up at vercel.com** with your GitHub account (free).
3. Click **Add New -> Project**, pick your `hk-gold-web` repo, and press **Deploy**.
   No settings to change - Vercel sees `index.html` and serves it, and sees `api/price.py`
   and runs it at `/api/price`.
4. You get a URL like `https://hk-gold-web.yourname.vercel.app`. Open it on the iPhone.
   Every time you push to GitHub, Vercel redeploys automatically.

That is still "deploying with GitHub" - the repo lives on GitHub, Vercel just runs it.

## The one thing to test first

I could not test the live fetch from here. The real risk with ANY free host is that XE or
Kitco may block or throttle requests coming from a data-center IP (Vercel's servers), even
though they work from a home PC. So right after deploying:

- Open the URL and press **Calculate now**.
- If you see a price, it works.
- If you see an error like "Could not reach XE" or "not in the expected layout", the host's
  IP is likely being blocked or served different HTML. Tell me and we will look at options
  (different host, or adjusting how it fetches).

## Pure-GitHub alternative (no Vercel)

You can keep everything inside GitHub using **GitHub Pages + a GitHub Actions workflow**
that fetches the price on a schedule and commits it as a file the page reads. Downsides:
the schedule is only roughly every 15 minutes (often delayed), "Calculate now" can't fetch
a fresh price on demand, and the schedule auto-disables after 60 days of no repo activity.
Say the word if you want this version instead.
