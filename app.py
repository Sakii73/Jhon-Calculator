"""
app.py  -  the whole web program for Vercel: it serves the page AND does the
server-side price read, in one small Flask app.

Why Flask: Vercel's Python runtime loads one app object named `app`. Flask lets us
define exactly two addresses - "/" for the page and "/api/price" for the price - so
there is no guessing about routing.

A browser cannot read XE and Kitco directly (those sites block cross-site requests),
so "/api/price" does it on the server and hands back the finished numbers as JSON.

All the real logic - fetching, the formula, rounding, the client's text format - stays
in core.py, the SAME file the Windows program and its tests use. One source of truth.
"""
import os
import sys
from datetime import datetime

from flask import Flask, Response, jsonify

# Make sure "import core" finds core.py next to this file, whatever folder Vercel runs us from.
BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
import core  # noqa: E402  (must come after the line above)

app = Flask(__name__)


def build_payload():
    """Do one full calculation. Returns a dict the web page can show directly.
    Raises core.FetchError if a source can't be read or is too old - exactly like
    the desktop program, which refuses to compute with missing or stale values."""
    raw_rate, xe_time = core.get_us_rate()
    us = core.round_us(raw_rate)
    gold, kitco_time, source = core.get_gold_bid()
    result = core.compute_hk(us, gold)

    # The "as of" clock is Philippine time. The server may sit in another country,
    # so convert "now" to Manila rather than trusting the server's own clock zone.
    now_ph = datetime.now(core.UTC).astimezone(core.MANILA)
    xe_ph = xe_time.astimezone(core.MANILA)
    kitco_ph = kitco_time.astimezone(core.MANILA)

    return {
        "ok": True,
        "block": core.client_block(result, now_ph),         # the 8:14 / Us / Gold / Hk lines
        "sentence": core.client_sentence(result, now_ph),   # "Thus, the current HK GOLD PRICE..."
        "steps": core.step_lines(result),                   # the computation, one line each
        "sources": core.source_text(raw_rate, us, xe_time, gold, kitco_time, source),
        "hk": result["hk"],
        "record": {                                         # one history row (desktop CSV columns)
            "date": now_ph.strftime("%Y-%m-%d"),
            "time": now_ph.strftime("%H:%M:%S"),
            "us_raw": format(raw_rate, "f"),
            "us": core.plain_2(us),
            "gold": core.plain_2(gold),
            "hk": result["hk"],
            "xe_time": xe_ph.strftime("%Y-%m-%d %H:%M:%S"),
            "kitco_time": kitco_ph.strftime("%Y-%m-%d %H:%M:%S"),
        },
    }


@app.get("/api/price")
def api_price():
    try:
        payload = build_payload()
    except core.FetchError as error:        # a source failed or was stale: a clean message
        payload = {"ok": False, "error": str(error)}
    except Exception as error:              # anything else: still answer, never a blank 500
        payload = {"ok": False, "error": f"Unexpected problem: {error}"}
    response = jsonify(payload)
    response.headers["Cache-Control"] = "no-store"   # never show a cached price
    return response


@app.get("/")
def home():
    """Serve the page itself."""
    with open(os.path.join(BASE, "index.html"), encoding="utf-8") as page:
        return Response(page.read(), mimetype="text/html")
