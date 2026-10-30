"""
gui.py - the window of the HK Gold Price Calculator.   Run:  py gui.py
(The packaged .exe starts from start.py instead, which adds a safety net around this file.)

How it works (read this first):
  * There are two "feeds": the XE rate and the Kitco gold Bid. Both are read in the
    background every N seconds/minutes/hours - ONE setting for both, chosen in the window.
  * ONE clock ticks every second (tick). Each tick asks "is something due?":
        - time to read XE / Kitco again?           -> start_fetch(feed)
        - time for the automatic calculation?      -> start_calculation()
    then it refreshes the countdown and the age of the readings on screen.
  * Anything slow (downloading, writing the Excel file) runs in a BACKGROUND thread so the
    window never freezes. A background thread never touches the widgets. When it finishes it
    calls root.after(0, ...), which makes the window's own thread run fetch_done / calc_done /
    export_done.
  * The calculation uses the XE rate and the Kitco Bid that are already on screen (if they
    are still fresh). A source with nothing fresh on screen is read on the spot.
  * Every successful calculation is added to records.csv, and every calculation ATTEMPT
    (worked or failed) to log.txt - see records.py. "Export to Excel" turns records.csv into
    a formatted .xlsx.
  * The look: the build_* functions make the boxes ("cards"), setup_styles() holds all the
    colours. Change a colour there and the whole window follows.
"""
import json
import math
import sys
import threading
import time
import tkinter as tk
import traceback
from collections import deque
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, ttk

import core
import icon_data
import records

# ------------------------------------------------------------------ settings ---
AUTO_CALC_SECONDS = 15 * 60          # one calculation is SAVED to records.csv every 15 minutes
                                     # (the price on screen follows every new reading live)

UNITS = {"seconds": 1, "minutes": 60, "hours": 3600}        # unit name -> seconds
INTERVAL_MIN, INTERVAL_MAX = 1, 24 * 3600                   # allowed range, in seconds
DEFAULT_INTERVAL = (10, "seconds")                          # how often the gold price is read
FAST_SECONDS = 10                    # Kitco's gold price changes every 10 s; faster reads repeat it
XE_MIN_SECONDS = 60                  # XE is read at most once a minute, however fast gold is read:
                                     # its rate is only time-stamped to the minute, and reading
                                     # it every few seconds gets the PC's internet address blocked
ACTIVITY_LINES = 8                   # how many log lines the "Activity log" tab shows
ACTIVITY_CHARS = 84                  # longer lines are cut (the full text is in log.txt)

# colours: window background, white cards, their border, the dark blue header, gold accent
BG, CARD, BORDER = "#eef1f6", "#ffffff", "#d7deea"
NAVY, NAVY_HOVER, GOLD = "#14213d", "#26396a", "#d4af37"
GREEN, ORANGE, RED, GREY, INK = "#15803d", "#b45309", "#b91c1c", "#6b7280", "#1c2433"
# the status bar gets a pale background in the colour of its message
TINTS = {GREEN: "#e3f3e8", ORANGE: "#fdf0d8", RED: "#fde6e6", GREY: "#e8ecf3"}
UI_FONT = "Segoe UI"
MONO_FONT = "Consolas"


def app_folder():
    """Folder where the program lives, so records.csv / log.txt / settings.json sit
    next to it. The program can run three ways and each finds that folder differently:
      * Nuitka build (the "hard to read" .exe): Nuitka defines __compiled__ in every
        module, and __file__ then points to a temporary unpack folder - so we use
        sys.argv[0], which is the real .exe the user double-clicked.
      * PyInstaller build: sys.frozen is set and sys.executable is the real .exe.
      * plain "py gui.py" while developing: use this file's own folder.
    """
    if "__compiled__" in globals():            # Nuitka marks every compiled module
        return Path(sys.argv[0]).resolve().parent
    if getattr(sys, "frozen", False):          # PyInstaller sets sys.frozen
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


SETTINGS_FILE = app_folder() / "settings.json"
RECORDS_FILE = app_folder() / "records.csv"
LOG_FILE = app_folder() / "log.txt"


def plural(count, word):
    """plural(1, 'minute') -> '1 minute',  plural(2, 'minute') -> '2 minutes'"""
    return f"{count} {word}" if count == 1 else f"{count} {word}s"


def parse_interval(amount_text, unit):
    """('5', 'minutes') -> 5.  Returns None unless the amount is a whole number and the
    total time is between 1 second and 24 hours."""
    try:
        amount = int(amount_text.strip())
    except ValueError:
        return None
    if unit not in UNITS or not (INTERVAL_MIN <= amount * UNITS[unit] <= INTERVAL_MAX):
        return None
    return amount


def load_settings():
    """The saved reading interval as (amount, unit), e.g. (60, 'seconds').
    A missing, damaged or old-format file falls back to the default."""
    try:
        saved = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        unit = saved["unit"]
        amount = parse_interval(str(saved["amount"]), unit)
    except (OSError, ValueError, KeyError, TypeError):
        return DEFAULT_INTERVAL
    return (amount, unit) if amount else DEFAULT_INTERVAL


def save_settings(amount, unit):
    """Write the interval to settings.json. Raises OSError if the folder is not writable."""
    SETTINGS_FILE.write_text(json.dumps({"amount": amount, "unit": unit}, indent=2),
                             encoding="utf-8")


def age_text(seconds):
    seconds = max(0, int(seconds))
    return f"{seconds} s old" if seconds < 120 else f"{seconds // 60} min old"


def activity_text(line):
    """A log line as shown in the Activity tab: no year, and cut if it is very long."""
    text = line[5:]                                  # '2026-10-26 20:14:05 ...' -> '10-26 20:14:05 ...'
    return text if len(text) <= ACTIVITY_CHARS else text[:ACTIVITY_CHARS - 1] + "…"


def activity_colour(line):
    """Failures red, program start/stop grey, everything else normal."""
    what = line[21:]                                 # after 'YYYY-MM-DD HH:MM:SS  '
    if what.startswith(("START", "STOP")):
        return GREY
    if what.startswith("ERROR") or "FAILED" in what[:16]:
        return RED
    return INK


def read_xe():
    """core.get_us_rate() gives (rate, time); feeds want (value, time, source name)."""
    rate, xe_time = core.get_us_rate()
    return rate, xe_time, "XE"


def fresh_or_read(latest, name, read):
    """Use the reading already on screen if it is still fresh, otherwise read the source now."""
    if latest is not None:
        try:
            core.check_fresh(name, latest[1])
            return latest
        except core.FetchError:
            pass                 # too old: fall through and read the source now
    return read()


class Feed:
    """Everything the window needs to know about one data source (XE or Kitco)."""

    def __init__(self, key, title, read):
        self.key = key              # "xe" or "kitco"
        self.title = title          # shown above its reading
        self.read = read            # function that reads the source -> (value, time, source name)
        self.latest = None          # newest good reading: (value, time, source name)
        self.error = ""             # message if the newest read failed
        self.running = False        # True while a background read is in progress
        self.next_at = time.monotonic()    # when the next read is due (due now at startup)
        # the widgets of its card are stored here by App.build_feed_card:
        self.value_label = self.info_label = self.error_label = None


# -------------------------------------------------------------------- the app ---
class App:
    def __init__(self, root):
        self.root = root
        self.amount, self.unit = load_settings()        # ONE interval for both feeds
        self.xe = Feed("xe", "XE  ·  USD to PHP", read_xe)
        self.kitco = Feed("kitco", "Kitco  ·  gold Bid", core.get_gold_bid)
        self.feeds = [self.xe, self.kitco]
        self.last_data = None      # everything about the last successful calculation
        self.calc_running = False
        self.calc_trigger = "auto"              # who started the running calculation: auto / manual
        self.next_calc_at = time.monotonic()    # due immediately at startup

        self.saved_count = 0       # rows in records.csv
        self.unsaved_rows = []     # rows waiting because records.csv could not be written
        self.records_problem = ""  # red text about records.csv, "" when all is well
        self.export_message = ("", GREY)        # (text, colour) of the last Excel export
        self.export_running = False
        self.activity = deque(maxlen=ACTIVITY_LINES)    # newest log lines, oldest first
        self.load_history()

        root.report_callback_exception = self.callback_failed   # a bug is logged, not silent
        self.build_window()
        self.show_result()
        self.refresh_records_panel()
        self.set_status("Starting... reading XE and Kitco.", ORANGE)
        self.write_log(records.log_text(datetime.now(core.MANILA), "START", "Program opened"))
        self.tick()
        self.fix_window_size()

    def fix_window_size(self):
        """Fix the window size once, from what the content needs plus some spare room for
        longer values and error messages. Without this the window would grow and shrink
        whenever a long message appears. (Measuring, instead of typing a number, also copes
        with Windows display scaling of 125% / 150%.)"""
        self.root.update_idletasks()
        width = max(self.root.winfo_reqwidth() + 80, 1050)
        self.root.geometry(f"{width}x{self.root.winfo_reqheight() + 80}")

    def interval_seconds(self):
        """The reading interval in seconds."""
        return self.amount * UNITS[self.unit]

    def load_history(self):
        """Count what is already in records.csv and show the end of log.txt, so what
        happened in earlier runs is still visible."""
        try:
            rows, _ = records.read_rows(RECORDS_FILE)
            self.saved_count = len(rows)
        except OSError as error:
            self.records_problem = f"Could not read records.csv: {error}"
        try:
            self.activity.extend(records.read_log_tail(LOG_FILE, ACTIVITY_LINES))
        except OSError:
            pass

    # ---------------------------------------------------------------- layout ---
    def setup_styles(self):
        """All the colours and fonts of the buttons, boxes and tabs live here.
        'clam' is the one built-in theme that lets us set colours and looks the same on
        every PC, so we use it instead of the plain Windows look."""
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(".", background=BG, foreground=INK, font=(UI_FONT, 10))
        style.configure("Card.TFrame", background=CARD)
        style.configure("Card.TLabel", background=CARD, foreground=INK)
        # main button: dark blue, white text
        style.configure("Primary.TButton", font=(UI_FONT, 12, "bold"), padding=10,
                        background=NAVY, foreground="white", relief="flat", borderwidth=0,
                        bordercolor=NAVY, lightcolor=NAVY, darkcolor=NAVY, focuscolor=NAVY)
        style.map("Primary.TButton",
                  background=[("disabled", "#aab3c5"), ("active", NAVY_HOVER)],
                  foreground=[("disabled", "#f2f4f8")])
        # normal buttons: soft blue-grey
        style.configure("Soft.TButton", padding=(12, 5), background="#e3e9f4",
                        foreground=NAVY, relief="flat", borderwidth=0, bordercolor="#e3e9f4",
                        lightcolor="#e3e9f4", darkcolor="#e3e9f4", focuscolor="#e3e9f4")
        style.map("Soft.TButton",
                  background=[("disabled", "#f0f2f7"), ("active", "#d3dcee")],
                  foreground=[("disabled", "#a3abbb")])
        style.configure("TEntry", fieldbackground="white", bordercolor=BORDER,
                        lightcolor=BORDER, darkcolor=BORDER, padding=4)
        style.configure("TCombobox", fieldbackground="white", background="#e3e9f4",
                        bordercolor=BORDER, arrowcolor=NAVY, padding=4)
        style.map("TCombobox", fieldbackground=[("readonly", "white")],
                  foreground=[("readonly", INK)])
        style.configure("TNotebook", background=CARD, borderwidth=0)
        style.configure("TNotebook.Tab", padding=(16, 7), background="#e9edf5",
                        foreground=GREY, font=(UI_FONT, 10, "bold"), borderwidth=0)
        style.map("TNotebook.Tab", background=[("selected", CARD)],
                  foreground=[("selected", NAVY)])

    def load_icon(self):
        """The little gold coin in the title bar and the header (see icon_data.py)."""
        try:
            self.icon_big = tk.PhotoImage(data=icon_data.ICON_PNG_BASE64)
            self.root.iconphoto(True, self.icon_big)
            self.icon_small = self.icon_big.subsample(2)
        except tk.TclError:
            self.icon_big = self.icon_small = None      # no icon is not worth a crash

    def make_card(self, parent, accent=None):
        """A white box with a thin border (and optionally a coloured bar on top).
        Returns (outer, inner): place `outer` in the window, put the widgets in `inner`."""
        outer = tk.Frame(parent, background=BORDER, padx=1, pady=1)
        if accent:
            tk.Frame(outer, background=accent, height=4).pack(fill="x")
        inner = ttk.Frame(outer, style="Card.TFrame", padding=14)
        inner.pack(fill="both", expand=True)
        return outer, inner

    @staticmethod
    def card_title(parent, text):
        """The small grey heading at the top of a card."""
        return ttk.Label(parent, text=text.upper(), style="Card.TLabel", foreground=GREY,
                         font=(UI_FONT, 9, "bold"))

    def build_window(self):
        root = self.root
        root.title("HK Gold Price Calculator")
        root.configure(background=BG)
        self.setup_styles()
        self.load_icon()
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        main = ttk.Frame(root)
        main.grid(row=0, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        self.build_header(main).grid(row=0, column=0, sticky="ew")
        self.build_status_bar(main).grid(row=1, column=0, sticky="ew")

        body = ttk.Frame(main, padding=(18, 14, 18, 16))
        body.grid(row=2, column=0, sticky="nsew")
        body.columnconfigure(0, weight=0)          # left: the result, keeps its own width
        body.columnconfigure(1, weight=1)          # right: takes the rest
        left = ttk.Frame(body)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 16))
        self.build_result_card(left).grid(row=0, column=0, sticky="ew")
        self.build_records_card(left).grid(row=1, column=0, sticky="ew", pady=(14, 0))

        right = ttk.Frame(body)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1, uniform="feeds")     # XE and Kitco cards: same width
        right.columnconfigure(1, weight=1, uniform="feeds")
        self.build_feed_card(right, self.xe).grid(row=0, column=0, sticky="nsew", padx=(0, 7))
        self.build_feed_card(right, self.kitco).grid(row=0, column=1, sticky="nsew", padx=(7, 0))
        self.build_refresh_card(right).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(14, 0))
        self.build_tabs(right).grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(14, 0))
        root.minsize(960, 560)

    def build_header(self, parent):
        """The dark blue band at the top: coin, title, one-line description."""
        header = tk.Frame(parent, background=NAVY, padx=18, pady=8)
        header.columnconfigure(1, weight=1)
        if self.icon_small is not None:
            tk.Label(header, image=self.icon_small, background=NAVY).grid(
                row=0, column=0, rowspan=2, padx=(0, 12))
        tk.Label(header, text="HK Gold Price Calculator", background=NAVY, foreground="white",
                 font=(UI_FONT, 17, "bold")).grid(row=0, column=1, sticky="w")
        tk.Label(header, text="USD to PHP rate from XE  ·  gold Bid from Kitco  ·  "
                              "all times are Philippine time",
                 background=NAVY, foreground="#aeb9d3",
                 font=(UI_FONT, 9)).grid(row=1, column=1, sticky="w")
        return header

    def build_status_bar(self, parent):
        """One line under the header: coloured dot + message, on a pale background."""
        self.status_bar = tk.Frame(parent, background=TINTS[GREY], padx=18, pady=6)
        self.status_dot = tk.Label(self.status_bar, text="●", font=(UI_FONT, 12),
                                   foreground=GREY, background=TINTS[GREY])
        self.status_dot.grid(row=0, column=0, sticky="n", padx=(0, 8))
        self.status_text = tk.Label(self.status_bar, text="", wraplength=900, justify="left",
                                    foreground=INK, background=TINTS[GREY],
                                    font=(UI_FONT, 10))
        self.status_text.grid(row=0, column=1, sticky="w")
        return self.status_bar

    def build_result_card(self, parent):
        outer, box = self.make_card(parent, accent=GOLD)
        box.columnconfigure(0, weight=1)
        self.card_title(box, "HK Gold Price").grid(row=0, column=0, sticky="w")
        self.stale_label = ttk.Label(
            box, style="Card.TLabel", foreground=RED, justify="left",
            text="NOT CURRENT - the last attempt failed.\nShowing the last successful result.")
        self.stale_label.grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.stale_label.grid_remove()                      # hidden until needed
        self.block_labels = []                              # time, Us, Gold, Hk lines
        for row, (size, weight) in enumerate([(13, "normal"), (20, "normal"),
                                              (20, "normal"), (34, "bold")], start=2):
            label = ttk.Label(box, style="Card.TLabel", font=(MONO_FONT, size, weight))
            label.grid(row=row, column=0, sticky="w", pady=(8 if row == 2 else 0, 0))
            self.block_labels.append(label)
        self.sentence_label = ttk.Label(box, style="Card.TLabel", wraplength=330,
                                        justify="left", font=(UI_FONT, 11))
        self.sentence_label.grid(row=6, column=0, sticky="w", pady=(10, 10))
        self.copy_button = ttk.Button(box, text="Copy result", style="Soft.TButton",
                                      command=self.copy_result)
        self.copy_button.grid(row=7, column=0, sticky="w")
        self.calc_button = ttk.Button(box, text="Calculate now", style="Primary.TButton",
                                      command=lambda: self.start_calculation("manual"))
        self.calc_button.grid(row=8, column=0, sticky="ew", pady=(14, 6))
        self.countdown_label = ttk.Label(box, style="Card.TLabel", foreground=GREY)
        self.countdown_label.grid(row=9, column=0, sticky="w")
        return outer

    def build_records_card(self, parent):
        outer, box = self.make_card(parent)
        box.columnconfigure(0, weight=1)
        self.card_title(box, "History").grid(row=0, column=0, sticky="w")
        self.count_label = ttk.Label(box, style="Card.TLabel", font=(UI_FONT, 12, "bold"))
        self.count_label.grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.export_button = ttk.Button(box, text="Export to Excel", style="Soft.TButton",
                                        command=self.export_excel)
        self.export_button.grid(row=1, column=1, sticky="e")
        self.path_label = ttk.Label(box, style="Card.TLabel", wraplength=330, justify="left",
                                    foreground=GREY, font=(UI_FONT, 9),
                                    text=f"Saved in {RECORDS_FILE}")
        self.path_label.grid(row=2, column=0, columnspan=2, sticky="w", pady=(2, 0))
        self.problem_label = ttk.Label(box, style="Card.TLabel", wraplength=330,
                                       justify="left", foreground=RED)
        self.problem_label.grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.problem_label.grid_remove()                    # hidden until something goes wrong
        self.export_label = ttk.Label(box, style="Card.TLabel", wraplength=330, justify="left")
        self.export_label.grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.export_label.grid_remove()
        return outer

    def build_feed_card(self, parent, feed):
        """One live reading: title, big value, how old it is, and a red line if a read failed."""
        outer, box = self.make_card(parent, accent=NAVY)
        box.columnconfigure(0, weight=1)
        self.card_title(box, feed.title).grid(row=0, column=0, sticky="w")
        feed.value_label = ttk.Label(box, style="Card.TLabel", foreground=NAVY,
                                     font=(UI_FONT, 22, "bold"))
        feed.value_label.grid(row=1, column=0, sticky="w", pady=(4, 2))
        feed.info_label = ttk.Label(box, style="Card.TLabel", wraplength=250, justify="left",
                                    foreground=GREY, font=(UI_FONT, 9))
        feed.info_label.grid(row=2, column=0, sticky="w")
        feed.error_label = ttk.Label(box, style="Card.TLabel", wraplength=250, justify="left",
                                     foreground=RED, font=(UI_FONT, 9))
        feed.error_label.grid(row=3, column=0, sticky="w", pady=(6, 0))
        feed.error_label.grid_remove()                      # hidden until a read fails
        return outer

    def build_refresh_card(self, parent):
        """ONE setting and ONE button for both sources."""
        outer, box = self.make_card(parent)
        box.columnconfigure(5, weight=1)                    # pushes the button to the right
        self.card_title(box, "Refresh XE and Kitco").grid(row=0, column=0, columnspan=6, sticky="w")
        ttk.Label(box, text="Read gold every", style="Card.TLabel").grid(
            row=1, column=0, padx=(0, 8), pady=(8, 0))
        self.amount_var = tk.StringVar(value=str(self.amount))
        entry = ttk.Entry(box, textvariable=self.amount_var, width=6, justify="center")
        entry.grid(row=1, column=1, pady=(8, 0))
        entry.bind("<Return>", lambda event: self.apply_interval())         # Enter = Set
        self.unit_var = tk.StringVar(value=self.unit)
        ttk.Combobox(box, textvariable=self.unit_var, values=list(UNITS), state="readonly",
                     width=8).grid(row=1, column=2, padx=8, pady=(8, 0))
        ttk.Button(box, text="Set", style="Soft.TButton", width=5,
                   command=self.apply_interval).grid(row=1, column=3, pady=(8, 0))
        self.update_button = ttk.Button(box, text="Update both now", style="Soft.TButton",
                                        command=self.update_both)
        self.update_button.grid(row=1, column=5, sticky="e", pady=(8, 0))
        self.interval_message = ttk.Label(box, style="Card.TLabel", wraplength=540,
                                          justify="left", foreground=GREY, font=(UI_FONT, 9),
                                          text="Allowed: 1 second to 24 hours. Kitco's gold "
                                               "changes every 10 seconds; XE is read at most "
                                               "once a minute.")
        self.interval_message.grid(row=2, column=0, columnspan=6, sticky="w", pady=(6, 0))
        return outer

    def build_tabs(self, parent):
        """Two tabs: how the price was calculated, and the activity log."""
        outer = tk.Frame(parent, background=BORDER, padx=1, pady=1)
        self.tabs = tabs = ttk.Notebook(outer)
        tabs.pack(fill="both", expand=True)

        steps_tab = ttk.Frame(tabs, style="Card.TFrame", padding=14)
        tabs.add(steps_tab, text="How it was calculated")
        self.steps_label = ttk.Label(steps_tab, style="Card.TLabel", font=(MONO_FONT, 10),
                                     justify="left")
        self.steps_label.grid(row=0, column=0, sticky="w")
        self.sources_label = ttk.Label(steps_tab, style="Card.TLabel", wraplength=560,
                                       justify="left", foreground=GREY, font=(UI_FONT, 9))
        self.sources_label.grid(row=1, column=0, sticky="w", pady=(10, 0))

        log_tab = ttk.Frame(tabs, style="Card.TFrame", padding=14)
        tabs.add(log_tab, text="Activity log")
        self.activity_labels = []
        for row in range(ACTIVITY_LINES):
            label = ttk.Label(log_tab, style="Card.TLabel", font=(MONO_FONT, 9))
            label.grid(row=row, column=0, sticky="w")
            self.activity_labels.append(label)
        ttk.Label(log_tab, style="Card.TLabel", foreground=GREY, font=(UI_FONT, 9),
                  text=f"Newest first. The full history is in {LOG_FILE.name}.").grid(
            row=ACTIVITY_LINES, column=0, sticky="w", pady=(8, 0))
        self.refresh_activity()
        return outer

    # ------------------------------------------------------------ the clock ---
    def tick(self):
        self.root.after(1000, self.tick)      # schedule the NEXT tick first: one bad tick can't stop the clock
        now = time.monotonic()
        for feed in self.feeds:
            if now >= feed.next_at and not feed.running:
                self.start_fetch(feed)
        if now >= self.next_calc_at and not self.calc_running:
            # At startup, wait for the first readings instead of asking the sites twice.
            if not any(feed.running and feed.latest is None for feed in self.feeds):
                self.start_calculation("auto")
        self.refresh_labels()

    def refresh_labels(self):
        """Countdown + both live readings. Runs every second and after every result."""
        left = max(0, math.ceil(self.next_calc_at - time.monotonic()))
        self.countdown_label.config(
            text=f"Next save to records.csv in {left // 60}:{left % 60:02d}")
        for feed in self.feeds:
            self.refresh_feed_card(feed)

    def refresh_feed_card(self, feed):
        if feed.latest is None:
            feed.value_label.config(text="--")
            feed.info_label.config(text="Waiting for the first reading...", foreground=GREY)
        else:
            value, moment, source = feed.latest
            age = (datetime.now(core.UTC) - moment).total_seconds()
            if feed.key == "kitco":
                feed.value_label.config(text=f"{value:,.2f}")
                info = (f"USD per troy ounce (Bid)\nprice time "
                        f"{core.ph_clock(moment, seconds=(source == core.SOURCE_JSON))}"
                        f"  ·  {age_text(age)}\n{source}")
            else:
                feed.value_label.config(text=format(value, "f"))
                info = (f"PHP for 1 USD\nrate time {core.ph_clock(moment)}"
                        f"  ·  {age_text(age)}")
            feed.info_label.config(
                text=info,
                foreground=RED if age > core.MAX_AGE_MINUTES * 60 else GREY)
        self.show_text(feed.error_label, feed.error)

    def refresh_records_panel(self):
        """The History card: how many rows are saved, problems with the file, export result."""
        self.count_label.config(
            text="No calculations saved yet." if self.saved_count == 0
            else f"{plural(self.saved_count, 'calculation')} saved")
        self.show_text(self.problem_label, self.records_problem)
        text, colour = self.export_message
        self.export_label.config(foreground=colour)
        self.show_text(self.export_label, text)
        # the file location is only shown while there is nothing more important to say
        self.show_text(self.path_label,
                       "" if (self.records_problem or text) else f"Saved in {RECORDS_FILE}")
        self.export_button.state(["disabled" if self.export_running else "!disabled"])
        self.export_button.config(text="Exporting..." if self.export_running
                                  else "Export to Excel")

    def refresh_activity(self):
        """The Activity log tab: the newest log lines, newest first."""
        lines = list(self.activity)[::-1]
        for label, line in zip(self.activity_labels, lines + [""] * ACTIVITY_LINES):
            label.config(text=activity_text(line) if line else "",
                         foreground=activity_colour(line) if line else INK)
        if not lines:
            self.activity_labels[0].config(text="No activity yet.", foreground=GREY)

    @staticmethod
    def show_text(label, text):
        """Set a label's text, and hide the label completely when there is nothing to say."""
        label.config(text=text)
        if text:
            label.grid()
        else:
            label.grid_remove()

    def set_status(self, text, colour):
        self.status_colour = colour          # remembered so live_update can clear an old error
        tint = TINTS.get(colour, TINTS[GREY])
        self.status_bar.config(background=tint)
        self.status_dot.config(foreground=colour, background=tint)
        self.status_text.config(text=text, background=tint)

    # ------------------------------------------------------------- the log ---
    def write_log(self, line):
        """Add a line to log.txt and to the Activity tab. A problem with the log file must
        never stop the program, so write errors are ignored."""
        self.activity.append(line)
        try:
            records.append_log(LOG_FILE, line)
        except OSError:
            pass
        self.refresh_activity()

    def callback_failed(self, exc_type, value, tb):
        """Tkinter calls this when something inside the window code crashes (a bug).
        Without it a packaged .exe would swallow the error with no trace."""
        last = traceback.extract_tb(tb)[-1] if tb else None
        where = f" (in {Path(last.filename).name}, line {last.lineno}, {last.name})" if last else ""
        self.write_log(records.log_text(datetime.now(core.MANILA), "ERROR",
                                        f"{exc_type.__name__}: {value}{where}"))

    def close(self):
        """The window's X button: note it in the log, then close."""
        self.write_log(records.log_text(datetime.now(core.MANILA), "STOP", "Program closed"))
        self.root.destroy()

    # ----------------------------- reading XE / Kitco in the background ---
    def update_both(self):
        """The 'Update both now' button: read XE and Kitco right now."""
        for feed in self.feeds:
            self.start_fetch(feed)

    def start_fetch(self, feed):
        if feed.running:
            return
        feed.running = True
        self.refresh_update_button()
        threading.Thread(target=self.fetch_worker, args=(feed,), daemon=True).start()

    def refresh_update_button(self):
        """Grey out 'Update both now' while any source is being read."""
        busy = any(feed.running for feed in self.feeds)
        self.update_button.state(["disabled" if busy else "!disabled"])
        self.update_button.config(text="Updating..." if busy else "Update both now")

    def fetch_worker(self, feed):
        # BACKGROUND thread: only call core here, never touch widgets.
        try:
            outcome = ("ok", feed.read())
        except core.FetchError as error:
            outcome = ("error", str(error))
        except Exception as error:               # a bug: still report it, don't die silently
            outcome = ("error", f"Unexpected problem: {error}")
        self.send_to_window(self.fetch_done, feed, outcome)

    def send_to_window(self, function, *args):
        """Hand a result from a background thread to the window's own thread."""
        try:
            self.root.after(0, function, *args)
        except (RuntimeError, tk.TclError):
            pass    # the window was closed while the thread was still working

    def fetch_done(self, feed, outcome):          # runs on the window's thread
        feed.running = False
        self.refresh_update_button()
        # The next read is one interval after this one FINISHED, so reads never overlap.
        # XE never faster than once a minute (see XE_MIN_SECONDS); gold at the user's interval.
        wait = self.interval_seconds()
        if feed.key == "xe":
            wait = max(wait, XE_MIN_SECONDS)
        feed.next_at = time.monotonic() + wait
        kind, value = outcome
        if kind == "ok":
            feed.latest = value
            feed.error = ""
            self.live_update()
        else:
            feed.error = f"Last read failed: {value}"
        self.refresh_labels()

    def live_update(self):
        """Recompute Hk from the newest XE and Kitco readings and show it straight away, so the
        price on screen moves when Kitco's price moves. No internet here - both numbers are
        already on screen - and nothing is written to records.csv or log.txt (that still
        happens once every 15 minutes and on Calculate now, so the files don't fill up with a
        row every few seconds)."""
        if self.xe.latest is None or self.kitco.latest is None:
            return                       # still waiting for the first reading of one of them
        raw_rate, xe_time, _ = self.xe.latest
        bid, kitco_time, source = self.kitco.latest
        try:
            core.check_fresh("XE", xe_time)          # same rule as a saved calculation:
            core.check_fresh("Kitco", kitco_time)    # never show a price built on old numbers
        except core.FetchError:
            return
        us = core.round_us(raw_rate)
        self.last_data = {
            "calculated_at": datetime.now(core.MANILA),
            "raw_rate": raw_rate, "us": us, "xe_time": xe_time,
            "gold": bid, "kitco_time": kitco_time, "source": source,
            "result": core.compute_hk(us, bid),
        }
        self.show_result()
        if getattr(self, "status_colour", None) == RED:
            # an earlier failure is over: don't leave a red error next to a good, live price
            self.set_status(f"Live price at {core.clock_12h(self.last_data['calculated_at'])}"
                            f"  ·  Gold Bid from the {source}", GREEN)

    # ------------------------------------------------------- the calculation ---
    def start_calculation(self, trigger="auto"):
        """trigger is 'auto' (the 15-minute timer) or 'manual' (the button); it goes in the log."""
        if self.calc_running:
            return
        self.calc_running = True
        self.calc_trigger = trigger
        self.next_calc_at = time.monotonic() + AUTO_CALC_SECONDS   # restart the countdown
        self.calc_button.state(["disabled"])
        self.set_status("Calculating...", ORANGE)
        self.refresh_labels()
        threading.Thread(target=self.calc_worker, args=(self.xe.latest, self.kitco.latest),
                         daemon=True).start()

    def calc_worker(self, xe_latest, kitco_latest):
        # BACKGROUND thread: only call core here, never touch widgets.
        try:
            raw_rate, xe_time, _ = fresh_or_read(xe_latest, "XE", self.xe.read)
            bid, kitco_time, source = fresh_or_read(kitco_latest, "Kitco", self.kitco.read)
            us = core.round_us(raw_rate)
            outcome = ("ok", {
                "calculated_at": datetime.now(core.MANILA),     # Philippine time, whatever the PC's zone is
                "raw_rate": raw_rate, "us": us, "xe_time": xe_time,
                "gold": bid, "kitco_time": kitco_time, "source": source,
                "result": core.compute_hk(us, bid),
            })
        except core.FetchError as error:
            outcome = ("error", str(error))
        except Exception as error:
            outcome = ("error", f"Unexpected problem: {error}")
        self.send_to_window(self.calc_done, outcome)

    def calc_done(self, outcome):                 # runs on the window's thread
        self.calc_running = False
        self.calc_button.state(["!disabled"])
        kind, value = outcome
        if kind == "ok":
            self.last_data = value
            self.show_result()
            saved = self.save_record(value)
            line = records.log_ok(value, self.calc_trigger)
            self.write_log(line if saved else line + "  [not saved to records.csv yet]")
            self.set_status(f"Calculated at {core.clock_12h(value['calculated_at'])}  ·  "
                            f"Gold Bid from the {value['source']}"
                            + ("  ·  Saved to records.csv" if saved else ""), GREEN)
        else:
            self.write_log(records.log_failed(datetime.now(core.MANILA), self.calc_trigger, value))
            self.show_result(stale=True)
            self.set_status(f"{value}\nNo new price was calculated. It will try again at "
                            f"the next automatic calculation (or press Calculate now).", RED)
        self.refresh_labels()

    def show_result(self, stale=False):
        """Fill the result area from the last successful calculation.
        stale=True greys it out and shows a warning, so an old price can't pass as current."""
        data = self.last_data
        if data is None:
            for label, text in zip(self.block_labels, ("--:--", "Us:--", "Gold:--", "Hk=--")):
                label.config(text=text, foreground=GREY)
            self.sentence_label.config(text="No price calculated yet.")
            self.steps_label.config(text="No calculation yet.")
            self.sources_label.config(text="")
            self.copy_button.state(["disabled"])
            return
        result, moment = data["result"], data["calculated_at"]
        colour = GREY if stale else INK
        for label, text in zip(self.block_labels,
                               core.client_block(result, moment).split("\n")):
            label.config(text=text, foreground=colour)
        self.sentence_label.config(text=core.client_sentence(result, moment),
                                   foreground=colour)
        self.steps_label.config(text="\n".join(core.step_lines(result)))
        self.sources_label.config(text=core.source_text(
            data["raw_rate"], data["us"], data["xe_time"],
            data["gold"], data["kitco_time"], data["source"]))
        if stale:
            self.stale_label.grid()
            self.copy_button.state(["disabled"])
        else:
            self.stale_label.grid_remove()
            self.copy_button.state(["!disabled"])

    # ------------------------------------------------------ records.csv ---
    def save_record(self, data):
        """Add this calculation to records.csv. Returns True when it is safely on disk.
        If the file can't be written (most often: it is open in Excel) the row is kept in
        memory and saved together with the next calculation - nothing is lost."""
        self.unsaved_rows.append(records.make_row(data))
        return self.flush_unsaved_rows()

    def flush_unsaved_rows(self):
        if not self.unsaved_rows:
            return True
        try:
            records.append_rows(RECORDS_FILE, self.unsaved_rows)
        except OSError:
            waiting = len(self.unsaved_rows)
            self.records_problem = (
                f"Could not save to records.csv. Close it if it is open in Excel (or move the "
                f"program out of Program Files). {plural(waiting, 'calculation')} "
                f"{'is' if waiting == 1 else 'are'} waiting and will be saved after the next "
                f"calculation.")
            saved = False
        else:
            self.saved_count += len(self.unsaved_rows)
            self.unsaved_rows = []
            self.records_problem = ""
            saved = True
        self.refresh_records_panel()
        return saved

    # ------------------------------------------------------- Excel export ---
    def export_excel(self):
        """The 'Export to Excel' button: ask where to save, then write the file in a thread."""
        if self.export_running:
            return
        self.flush_unsaved_rows()               # try to get any waiting rows into the file first
        if self.saved_count == 0:
            self.export_message = ("There is nothing to export yet - no calculation has "
                                   "been saved.", RED)
            self.refresh_records_panel()
            return
        path = filedialog.asksaveasfilename(
            parent=self.root, title="Export to Excel", defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
            initialfile=f"HK Gold Records {datetime.now(core.MANILA):%Y-%m-%d}.xlsx")
        if not path:                            # the user pressed Cancel
            return
        self.export_running = True
        self.export_message = ("", GREY)
        self.refresh_records_panel()
        threading.Thread(target=self.export_worker, args=(Path(path),), daemon=True).start()

    def export_worker(self, target):
        # BACKGROUND thread: only write the file, never touch widgets.
        try:
            outcome = ("ok", records.export_to_excel(RECORDS_FILE, target))
        except ValueError as error:              # nothing to export
            outcome = ("error", str(error))
        except ImportError as error:             # openpyxl missing
            outcome = ("error", str(error))
        except PermissionError:
            outcome = ("error", f"Could not save {target.name}. If it is open in Excel, "
                                f"close it and try again, or pick another name or folder.")
        except OSError as error:
            outcome = ("error", f"Could not save the file: {error}")
        except Exception as error:
            outcome = ("error", f"Unexpected problem: {error}")
        self.send_to_window(self.export_done, target, outcome)

    def export_done(self, target, outcome):       # runs on the window's thread
        self.export_running = False
        kind, value = outcome
        if kind == "ok":
            count, skipped = value
            text = f"Exported {plural(count, 'calculation')} to {target}"
            if skipped:
                text += f" ({plural(skipped, 'damaged line')} in records.csv left out)"
            self.export_message = (text, GREEN)
        else:
            self.export_message = (value, RED)
        self.refresh_records_panel()

    # --------------------------------------------------------------- buttons ---
    def apply_interval(self):
        """The Set button (or Enter): check the input, save it, read both sources right now."""
        unit = self.unit_var.get()
        amount = parse_interval(self.amount_var.get(), unit)
        if amount is None:
            self.interval_message.config(
                text="Type a whole number. Allowed: 1 second to 24 hours.", foreground=RED)
            return
        self.amount, self.unit = amount, unit
        self.amount_var.set(str(amount))                       # tidy up e.g. " 10 "
        every = plural(amount, unit[:-1])
        message = f"Saved. Reading gold every {every} (XE at most once a minute)."
        colour = GREEN
        try:
            save_settings(amount, unit)
        except OSError:
            message = (f"Reading gold every {every} (XE at most once a minute), but the "
                       f"setting could not be saved (this folder is not writable).")
        if self.interval_seconds() < FAST_SECONDS:
            message += ("\nKitco's gold price changes every 10 seconds, so reading faster "
                        "mostly repeats the same price.")
            colour = ORANGE
        self.interval_message.config(text=message, foreground=colour)
        self.update_both()        # read now, so the change is visible straight away

    def copy_result(self):
        data = self.last_data
        if data is None:
            return
        text = (core.client_block(data["result"], data["calculated_at"]) + "\n\n" +
                core.client_sentence(data["result"], data["calculated_at"]))
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.copy_button.config(text="Copied!")
        self.root.after(1500, lambda: self.copy_button.config(text="Copy result"))


def main():
    root = tk.Tk()
    app = App(root)
    root.protocol("WM_DELETE_WINDOW", app.close)
    root.mainloop()


if __name__ == "__main__":
    main()
