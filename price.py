"""
api/price.py  -  the one server-side job: read XE + Kitco and compute the HK price.

A web browser CANNOT read XE and Kitco directly (those sites block cross-site
requests from page JavaScript - "CORS"), so this small function does it on the
server and hands the finished numbers back to the page as JSON.

On Vercel this file becomes the address  /api/price  automatically (a .py file in
the api/ folder is served at its own path). The page calls that address.

All the real logic - fetching, the formula, rounding, the client's text format - is
in core.py, the SAME file the Windows program and its tests use. This function only
wires it together and turns the result into JSON, so there is one source of truth.
"""
import json
import os
import sys
from datetime import datetime
from http.server import BaseHTTPRequestHandler

# Make sure "import core" finds api/core.py no matter what folder Vercel runs us from.
sys.path.insert(0, os.path.dirname(__file__))
import core  # noqa: E402  (must come after the line above)


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


class handler(BaseHTTPRequestHandler):
    """Vercel looks for a class named exactly 'handler' here."""

    def do_GET(self):
        try:
            payload = build_payload()
        except core.FetchError as error:        # a source failed or was stale: a clean message
            payload = {"ok": False, "error": str(error)}
        except Exception as error:              # anything else: still answer, never a blank 500
            payload = {"ok": False, "error": f"Unexpected problem: {error}"}

        body = json.dumps(payload).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")   # harmless; lets you test from anywhere
        self.send_header("Cache-Control", "no-store")          # never show a cached price
        self.end_headers()
        self.wfile.write(body)
