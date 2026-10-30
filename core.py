"""
core.py - the logic of the HK Gold Price Calculator. No GUI in this file.

What the GUI will call:
    get_us_rate()         -> (raw USD->PHP rate, time of that rate)
    get_gold_bid()        -> (gold Bid, time of that price, which source it came from)
    round_us(raw_rate)    -> the dollar rate cut to 2 decimals (extra decimals dropped)
    compute_hk(us, gold)  -> dict with every step of the client's formula

Plus text helpers at the bottom (client_block, client_sentence, step_lines, ...)
that build the exact text the window shows, so they can be tested without a window.

Needs Python 3.9+ (zoneinfo). On Windows also:  py -m pip install tzdata
All money math uses Decimal, never float.
"""
import html
import json
import re
import time
import urllib.request
from datetime import datetime, timedelta, timezone
from decimal import (Decimal, InvalidOperation,
                     ROUND_CEILING, ROUND_DOWN, ROUND_FLOOR, ROUND_HALF_UP)
from zoneinfo import ZoneInfo

# ------------------------------------------------------------------ settings ---
XE_URL = "https://www.xe.com/currencyconverter/convert/?Amount=1&From=USD&To=PHP"
KITCO_JSON_URL = "https://api.kitco.com/api/v1/precious-metals"
KITCO_HTML_URL = "https://www.kitco.com/charts/gold"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)
TIMEOUT_SECONDS = 25

# The ONE place that decides how Us is rounded to 2 decimals.
# Client's rule (decided 2026-10-30): CUT OFF the extra decimals, never round up:
#     63.70390875 -> 63.70,  63.709 -> 63.70,  63.951 -> 63.95
# Other choices, if the client ever changes his mind:
#     ROUND_CEILING = always up (63.7039 -> 63.71)   ROUND_HALF_UP = normal (63.706 -> 63.71)
US_ROUNDING = ROUND_DOWN

MAX_AGE_MINUTES = 20  # a source time older than this is treated as a failure

UTC = timezone.utc
NEW_YORK = ZoneInfo("America/New_York")  # Kitco's times
MANILA = ZoneInfo("Asia/Manila")         # for showing times to the client

# The client's formula, as named numbers so the code reads like the steps.
US_DEDUCTION = Decimal("0.25")
TAKE_OFF = Decimal("1")
FACTOR = Decimal("0.9965")
DIVISOR = Decimal("31.1")
TWO_DECIMALS = Decimal("0.01")

# Names of the two gold sources. get_gold_bid() returns one of these.
SOURCE_JSON = "Kitco JSON feed"
SOURCE_HTML = "Kitco web page"


class FetchError(Exception):
    """A price could not be read, or could not be trusted.
    The message is written so the GUI can show it to the user as-is."""


# ------------------------------------------------------------------- helpers ---
def fetch_text(url):
    """Download a URL and return its text. Any network problem becomes a FetchError."""
    request = urllib.request.Request(
        url, headers={"User-Agent": USER_AGENT, "Accept-Language": "en-US,en;q=0.9"}
    )
    try:
        with urllib.request.urlopen(request, timeout=TIMEOUT_SECONDS) as response:
            return response.read().decode("utf-8", errors="replace")
    except Exception as error:  # timeout, no internet, HTTP 403, ... all handled the same
        raise FetchError(f"Could not reach {url} ({error})") from error


def unique_address(url):
    """The same address plus '?nocache=<milliseconds>', different on every read.
    XE and Kitco both sit behind caches that hand out an OLD copy of a plain address
    (XE: for hours; Kitco's gold feed: about a minute, although Kitco itself updates
    every 10 seconds). A brand-new address is never in any cache, so we get today's
    real numbers. (Found with test_xe_live.py, and for Kitco by watching its charts page.)"""
    joiner = "&" if "?" in url else "?"
    return f"{url}{joiner}nocache={int(time.time() * 1000)}"


def page_texts(raw):
    """Two plain-text versions of a web page: tags removed with and without a space.
    (A site can split one number across several tags, so we try both.)"""
    raw = re.sub(r"(?is)<(script|style).*?</\1>", " ", raw)
    joined = html.unescape(re.sub(r"(?s)<[^>]+>", "", raw))
    spaced = html.unescape(re.sub(r"(?s)<[^>]+>", " ", raw))
    return [re.sub(r"\s+", " ", joined), re.sub(r"\s+", " ", spaced)]


def first_match(texts, pattern):
    for text in texts:
        match = re.search(pattern, text)
        if match:
            return match
    return None


def to_decimal(text):
    """'4,013.90' -> Decimal('4013.90')"""
    return Decimal(text.replace(",", "").replace(" ", ""))


def two_decimals(value):
    """Normal rounding to 2 decimals (4011.2 -> 4011.20). Used for Gold and for
    showing step e. NOT used for Us (that rounds up) or Hk (that truncates)."""
    return Decimal(value).quantize(TWO_DECIMALS, rounding=ROUND_HALF_UP)


def check_fresh(source_name, moment):
    """Raise FetchError if `moment` (a timezone-aware datetime) is older than the limit."""
    age = (datetime.now(UTC) - moment).total_seconds() / 60
    if age > MAX_AGE_MINUTES:
        # `moment` is the time stamped on the source's own data, not when we downloaded it.
        raise FetchError(f"The {source_name} reading is time-stamped {age:.0f} minutes ago "
                         f"(limit is {MAX_AGE_MINUTES} minutes), so I won't use it.")


# ------------------------------------------------------------------------ XE ---
def parse_xe(raw):
    """From the XE page text, return (full-precision rate, time of the rate in UTC)."""
    texts = page_texts(raw)
    rate = first_match(texts, r"1\.00\s*USD\s*=\s*([\d,]+\.[\d ]+?)\s*PHP")
    stamp = first_match(texts, r"Mid-market rate at (\d{1,2}):(\d{2}) UTC")
    # No low-precision backup pattern on purpose: a rate already rounded to
    # 63.95 would be rounded UP differently than the real 63.9507...
    if rate is None:
        raise FetchError("XE: could not find the USD to PHP rate on the page.")
    if stamp is None:
        raise FetchError("XE: could not find the 'Mid-market rate at' time, "
                         "so I can't tell whether the rate is fresh.")
    # XE only gives HH:MM, no date. Use today's date; if that lands in the
    # future, the time was from before midnight UTC, so use yesterday.
    now = datetime.now(UTC)
    xe_time = now.replace(hour=int(stamp.group(1)), minute=int(stamp.group(2)),
                          second=0, microsecond=0)
    if xe_time > now + timedelta(minutes=5):
        xe_time -= timedelta(days=1)
    return to_decimal(rate.group(1)), xe_time


def get_us_rate():
    """Read XE. Returns (raw rate as Decimal, time of the rate as a UTC datetime)."""
    # XE's CDN (CloudFront) keeps its page cached for hours, so the plain address can give
    # a copy that is more than an hour old. See unique_address().
    rate, xe_time = parse_xe(fetch_text(unique_address(XE_URL)))
    check_fresh("XE", xe_time)
    return rate, xe_time


def round_us(raw_rate):
    """The XE rate with 2 decimals, using US_ROUNDING (now: cut off): 63.70390875 -> 63.70."""
    return Decimal(raw_rate).quantize(TWO_DECIMALS, rounding=US_ROUNDING)


# --------------------------------------------------------------------- Kitco ---
def parse_kitco_json(raw):
    """From the JSON feed text, return (gold Bid with 2 decimals, time in New York)."""
    try:
        metals = json.loads(raw)["PreciousMetals"]["PM"]
    except (ValueError, KeyError, TypeError):
        raise FetchError("Kitco JSON feed: the data is not in the expected layout.")
    for metal in metals:
        if metal.get("Symbol") == "AU":  # AU = gold
            try:
                bid = Decimal(metal["current_bid"])
                naive = datetime.strptime(metal["timestamp"], "%Y-%m-%d %H:%M:%S")
            except (KeyError, ValueError, TypeError, InvalidOperation):
                raise FetchError("Kitco JSON feed: the gold entry has no readable bid or time.")
            if bid <= 0:
                raise FetchError(f"Kitco JSON feed: gold bid is {bid}, which can't be right.")
            return two_decimals(bid), naive.replace(tzinfo=NEW_YORK)
    raise FetchError("Kitco JSON feed: no gold (AU) entry found.")


def read_kitco_json():
    # Kitco's feed changes every 10 seconds, but the plain address is cached for about a
    # minute - that is why fast reads used to repeat the same gold price. See unique_address().
    return parse_kitco_json(fetch_text(unique_address(KITCO_JSON_URL)))


def parse_kitco_html(raw):
    """From the Kitco chart page text, return (gold Bid with 2 decimals, time in New York)."""
    texts = page_texts(raw)
    bid = first_match(texts, r"Bid\s*([\d,]+\.\d{2})")
    sentence = first_match(
        texts,
        r"current price of gold is \$([\d,]+\.\d{2}) per troy ounce as of "
        r"([A-Z][a-z]+ \d{1,2}, \d{4}, \d{1,2}:\d{2} [ap]\.m\.) ET",
    )
    if bid is None or sentence is None:
        raise FetchError("Kitco web page: could not find the Bid and the 'as of' time.")
    if to_decimal(bid.group(1)) != to_decimal(sentence.group(1)):
        raise FetchError("Kitco web page: the Bid and the 'current price' sentence differ, "
                         "so I can't tell which time belongs to the Bid.")
    # "October 26, 2026, 9:10 a.m."  ->  "October 26, 2026, 9:10 AM"  ->  datetime
    text = sentence.group(2).replace("a.m.", "AM").replace("p.m.", "PM")
    try:
        naive = datetime.strptime(text, "%B %d, %Y, %I:%M %p")
    except ValueError:
        raise FetchError(f"Kitco web page: can't read the time '{sentence.group(2)}'.")
    return two_decimals(to_decimal(bid.group(1))), naive.replace(tzinfo=NEW_YORK)


def read_kitco_html():
    return parse_kitco_html(fetch_text(unique_address(KITCO_HTML_URL)))


def get_gold_bid():
    """Gold Bid in USD per troy ounce. Tries the JSON feed first, then the web page.
    Returns (bid, time of the price, name of the source that worked)."""
    problems = []
    for source_name, reader in ((SOURCE_JSON, read_kitco_json),
                                (SOURCE_HTML, read_kitco_html)):
        try:
            bid, kitco_time = reader()
            check_fresh("Kitco", kitco_time)  # an old price counts as a failure too
            return bid, kitco_time, source_name
        except FetchError as error:
            problems.append(f"{source_name}: {error}")
    raise FetchError("Could not get a fresh Kitco Bid.\n" + "\n".join(problems))


# ------------------------------------------------------------------- formula ---
def truncate_hk(e):
    """Drop the decimals, never round up: 8202.90 -> 8202. (Hk is always positive,
    so rounding down with ROUND_FLOOR is the same as chopping off the decimals.)"""
    return int(Decimal(e).to_integral_value(rounding=ROUND_FLOOR))


def compute_hk(us, gold):
    """The client's formula. Pass Decimals or strings (never floats).
    Returns every step, exact and unrounded, in a dict:
        a = Us - 0.25      b = a * Gold      c = b - 1
        d = c * 0.9965     e = d / 31.1      hk = e with decimals dropped
    """
    us = Decimal(us)
    gold = Decimal(gold)
    a = us - US_DEDUCTION
    b = a * gold
    c = b - TAKE_OFF
    d = c * FACTOR
    e = d / DIVISOR
    return {"us": us, "gold": gold, "a": a, "b": b, "c": c, "d": d, "e": e,
            "hk": truncate_hk(e)}


# -------------------------------------------------------------- display text ---
# Pure functions that build the text the window shows. They live here (not in the
# GUI) so test_core.py can check them without opening a window.

def plain_2(value):
    """Always 2 decimals, no commas: 4016 -> '4016.00'  (the client's Us and Gold lines)."""
    return format(two_decimals(value), "f")


def with_commas(value):
    """Exact value with commas; trailing zeros trimmed but at least 2 decimals kept:
    Decimal('255082.91406000') -> '255,082.91406'   (the client's step numbers)"""
    whole, _, decimals = format(Decimal(value), "f").partition(".")
    return f"{int(whole):,}.{decimals.rstrip('0').ljust(2, '0')}"


def clock_12h(moment, seconds=False, suffix=True):
    """12-hour clock without a leading zero: '8:14 P.M.' (suffix=False -> '8:14')."""
    hour = moment.hour % 12 or 12          # 0 -> 12, 13 -> 1, 12 stays 12
    text = f"{hour}:{moment.minute:02d}"
    if seconds:
        text += f":{moment.second:02d}"
    if suffix:
        text += " A.M." if moment.hour < 12 else " P.M."
    return text


def ph_clock(moment, seconds=False):
    """Show a source time (XE's UTC or Kitco's New York) as Philippine time: '9:19:18 P.M.'"""
    return clock_12h(moment.astimezone(MANILA), seconds=seconds)


def client_block(result, moment):
    """The client's 4-line format. `moment` is the PC's local time of the calculation."""
    return (f"{clock_12h(moment, suffix=False)}\n"
            f"Us:{plain_2(result['us'])}\n"
            f"Gold:{plain_2(result['gold'])}\n"
            f"Hk={result['hk']}")


def client_sentence(result, moment):
    return (f"Thus, the current HK GOLD PRICE as of {clock_12h(moment)} "
            f"is {result['hk']:,} pesos per gram.")


def step_lines(result):
    """The computation steps, one per line, showing the client's numbers in order."""
    r = result
    return [
        f"{plain_2(r['us'])} - {US_DEDUCTION} = {with_commas(r['a'])}",
        f"{with_commas(r['a'])} × {plain_2(r['gold'])} = {with_commas(r['b'])}",
        f"{with_commas(r['b'])} - {TAKE_OFF} = {with_commas(r['c'])}",
        f"{with_commas(r['c'])} × {FACTOR} = {with_commas(r['d'])}",
        f"{with_commas(r['d'])} ÷ {DIVISOR} = {with_commas(two_decimals(r['e']))}",
        f"Hk = {r['hk']:,}  (decimals dropped)",
    ]


def source_text(raw_rate, us, xe_time, gold, kitco_time, source):
    """Where the two numbers came from, with source times in Philippine time.
    The JSON feed has seconds; XE and the Kitco web page only have minutes."""
    rounded = {ROUND_CEILING: "rounded up", ROUND_DOWN: "cut"}.get(US_ROUNDING, "rounded")
    return (f"XE USD to PHP: {format(raw_rate, 'f')}, {rounded} to {plain_2(us)}\n"
            f"   rate time: {ph_clock(xe_time)} (Philippine time)\n"
            f"Kitco gold Bid: {plain_2(gold)} USD per troy ounce ({source})\n"
            f"   price time: {ph_clock(kitco_time, seconds=(source == SOURCE_JSON))} "
            f"(Philippine time)")
