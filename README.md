# HK Gold Price - web version

A small website that shows the client's "HK" gold price. It reads the USD-&gt;PHP rate
from XE and the gold Bid from Kitco, applies the same formula as the Windows program,
and recalculates every 15 minutes while the page is open (plus a "Calculate now" button).

It works in any browser, including Safari on an iPhone - no install, no sideloading.

## The files (flat layout - no subfolders)

```
app.py             the whole program for Vercel (serves the page + reads XE/Kitco)
core.py            the SAME logic file as the Windows program (formula, parsing, format)
index.html         the page the client sees (plain HTML/JavaScript, no libraries)
requirements.txt   lists Flask (the web framework) and tzdata
```

`app.py` uses Flask because Vercel's Python runtime loads one app object. It defines just
two addresses: `/` serves the page, and `/api/price` reads XE + Kitco and returns the
finished numbers as JSON. A browser can't read those sites directly (they block cross-site
requests), so the server does it. All the real logic stays in `core.py`, so there is one
source of truth shared with the Windows program.

History is saved **in the browser** (per device), not on the server. The client's phone
builds up its own history as they use it. "Export CSV" downloads it with the same columns
as the desktop `records.csv`.

## Deploy it for free (Vercel, from GitHub)

Vercel's free "Hobby" plan runs the Flask app, and it deploys straight from a GitHub repo.

1. **Put these four files in a GitHub repo**, all at the top level (no `api/` folder):
   ```
   app.py
   core.py
   index.html
   requirements.txt
   ```
2. **Sign up at vercel.com** with your GitHub account (free).
3. Click **Add New -> Project**, pick your repo, and press **Deploy**. Vercel sees Flask in
   `requirements.txt`, loads `app`, and runs it. No settings to change.
4. You get a URL like `https://hk-gold-web.yourname.vercel.app`. Open it on the iPhone.
   Every push to GitHub redeploys automatically.

## The one thing to test first

I could not test the live fetch from here. The real risk with ANY free host is that XE or
Kitco may block or throttle requests coming from a data-center IP (Vercel's servers), even
though they work from a home PC. So right after deploying:

- Open the URL and press **Calculate now**.
- A price appears -> it works.
- A red error like "Could not reach XE" or "not in the expected layout" -> the host's IP is
  likely blocked or served different HTML. Tell me and we will look at options.

## Pure-GitHub alternative (no Vercel)

You can keep everything inside GitHub using **GitHub Pages + a GitHub Actions workflow**
that fetches the price on a schedule and commits it as a file the page reads. Downsides:
the schedule is only roughly every 15 minutes (often delayed), "Calculate now" can't fetch
a fresh price on demand, and the schedule auto-disables after 60 days of no repo activity.
Say the word if you want this version instead.
