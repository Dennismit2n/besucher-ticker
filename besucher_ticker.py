"""
Besucher-Ticker v1.0 — GoatCounter-Overlay + System-Tray (Windows)
==================================================================
Zeigt die Besucherzahlen der eigenen GoatCounter-Seite live als
schwebendes, halbtransparentes Overlay. An/Aus über das Tray-Icon.

Aufbau und Bedienung sind bewusst dieselben wie beim Claude Usage
Ticker — nur die Datenquelle ist eine andere, und die Farben sind
schwarz / pink / blau.

Installation (einmalig):
    pip install requests pystray pillow

    Optional, aber empfohlen:
        pip install pywin32
    Damit wird das API-Token in der config.json per Windows-DPAPI
    verschlüsselt (gebunden an Benutzerkonto + Rechner) statt im
    Klartext zu liegen. Fehlt das Paket, läuft alles weiter.

API-Token besorgen (einmalig):
    1. Auf der eigenen GoatCounter-Seite anmelden
       (z. B. https://dennis-mit-2n.goatcounter.com)
    2. Oben rechts auf den Benutzernamen → "API"
    3. "Add new token" — Name frei wählbar, Häkchen bei
       "Read statistics" genügt. Mehr Rechte braucht der Ticker nicht.
    4. Token kopieren und beim ersten Start hier einfügen.

Start:
    python besucher_ticker.py   (oder als .exe / mit pythonw.exe)
    python besucher_ticker.py --selbsttest   (prüft die Rechenlogik)

Bedienung:
    - Overlay mit Maus verschiebbar (einfach ziehen)
    - Doppelklick aufs Overlay: Mini-/Maxi-Modus umschalten
    - Rechtsklick aufs Tray-Icon: Overlay an/aus, zur Mitte holen,
      Jetzt aktualisieren, Zeitraum, Zusatzblock, Einstellungen,
      Nächstes Theme, Zugang ändern, Beenden

Zu den Zahlen — was hier was bedeutet:
    GoatCounter unterscheidet zwei Dinge, und dieser Ticker vermischt
    sie nicht:
      * "Aufrufe"  = Seitenaufrufe insgesamt (Feld `total`)
      * "Besucher" = gezählte Besucher je Pfad (Feld `count` je Pfad)
    Die grossen Zahlen oben sind Aufrufe, die Balken sind Besucher.
    Deshalb ist die Summe der Balken NICHT die grosse Zahl.

    Ereignisse (die `data-count`-Klicks aus der Werkstatt) tragen in
    der API `event: true`. Sie werden aus den Werkzeug-Balken heraus-
    gefiltert — sonst zählte ein Klick doppelt: einmal als Seiten-
    aufruf und einmal als Ereignis.

Datenschutz: Das Token bleibt ausschliesslich lokal (config.json neben
dem Skript). Es werden nur Lesezugriffe auf die eigene Seite gemacht.
"""

from __future__ import annotations

import base64
import json
import os
import queue
import sys
import threading
import time
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
from datetime import datetime, timedelta, timezone

try:
    import msvcrt                       # nur Windows – für die Instanz-Sperre
except ImportError:
    msvcrt = None

try:
    import requests
    import pystray
    from PIL import Image, ImageDraw
except ImportError:
    print("Bitte zuerst installieren:  pip install requests pystray pillow")
    sys.exit(1)

try:
    import win32crypt                   # noqa: F401  (nur Verfügbarkeitstest)
    HAVE_DPAPI = True
except ImportError:
    HAVE_DPAPI = False

# ------------------------------ Konstanten ----------------------------------

# Als .exe (PyInstaller): config.json liegt neben der .exe, nicht im
# Temp-Entpackordner. Als .py: wie gehabt neben dem Skript.
if getattr(sys, "frozen", False):
    _BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    _BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(_BASE_DIR, "config.json")
LOCK_PATH = os.path.join(_BASE_DIR, ".besucher.lock")

RETRY_BACKOFF = (10, 30, 60)          # Sekunden nach normalen Fehlern
RATE_LIMIT_BACKOFF = (30, 120, 300)   # nach HTTP 429
RATE_LIMIT_MAX_WAIT = 3600
UA = "BesucherTicker/1.0 (personal read-only GoatCounter display)"

BAR_W, BAR_H = 190, 8

# GoatCounter erlaubt 4 Anfragen pro Sekunde. Ein Durchlauf braucht
# 3–4 Anfragen, die nacheinander laufen — deshalb ist selbst die
# schnellste Stufe hier unkritisch. Trotzdem keine Sekundenstufen:
# es gibt keinen Grund, fremde Server im Sekundentakt zu fragen.
REFRESH_CHOICES = [
    ("15 s", 15),
    ("30 s", 30),
    ("60 s  (Standard)", 60),
    ("2 min", 120),
    ("5 min", 300),
    ("15 min", 900),
    ("60 min", 3600),
]
DEFAULT_REFRESH = 60

COMPACT_DURATIONS = [("dauerhaft", None)] + [
    (f"{s} Sekunden", s) for s in (3, 5, 8, 15, 30, 60)
]

ALPHA_MIN, ALPHA_MAX = 20, 100

# Zeiträume für Balken und Zwischensumme: Schlüssel -> (Anzeigename, Tage)
# Tage = None bedeutet "alles".
ZEITRAEUME = [
    ("heute", "Heute", 0),
    ("7t", "7 Tage", 7),
    ("30t", "30 Tage", 30),
    ("gesamt", "Gesamt", None),
]
DEFAULT_ZEITRAUM = "heute"

# Zusatzblock unter den Balken
ZUSATZ_CHOICES = [
    ("aus", "aus", None),
    ("laender", "Länder", "locations"),
    ("verweise", "Verweise", "toprefs"),
]
DEFAULT_ZUSATZ = "verweise"

# Weit genug vor dem ersten Seitenaufruf, damit "Gesamt" wirklich alles
# erfasst. GoatCounter selbst kennt kein "seit Anbeginn"-Kürzel.
ALLTIME_START = "2020-01-01T00:00:00Z"
ALLTIME_CACHE_SECONDS = 600           # Gesamtsumme ändert sich träge

MAX_BALKEN_CHOICES = [3, 5, 8, 10, 12, 15]
DEFAULT_MAX_BALKEN = 8

# ------------------------------ Themes --------------------------------------
# Bewusst nur drei. Der Wunsch war eindeutig: schwarz, pinke Schrift,
# blaue Balken. Die beiden anderen sind Varianten desselben Gedankens.

THEMES: dict[str, dict] = {
    "Schwarz / Pink / Blau": {
        "bg": "#000000", "title": "#ff2d95", "text": "#ff8ac6",
        "muted": "#8a4a6b", "bar_bg": "#101a33", "bar": "#2f6bff",
        "big": "#ff2d95",
    },
    "Schwarz / Pink / Eisblau": {
        "bg": "#000000", "title": "#ff3fa4", "text": "#ffa3d5",
        "muted": "#8a4a6b", "bar_bg": "#0d2430", "bar": "#38bdf8",
        "big": "#ff3fa4",
    },
    "Tiefschwarz / Magenta / Indigo": {
        "bg": "#050008", "title": "#e619c8", "text": "#f3a6e6",
        "muted": "#7a3f74", "bar_bg": "#141033", "bar": "#6366f1",
        "big": "#e619c8",
    },
}
THEME_NAMES = list(THEMES.keys())
DEFAULT_THEME = "Schwarz / Pink / Blau"

# Pfad -> hübscher Name. Was hier nicht steht, bekommt den Pfad selbst
# als Beschriftung; das Overlay bleibt damit auch bei neuen Werkzeugen
# vollständig, nur eben unbeschönigt.
PFAD_NAMEN = {
    "/": "Startseite",
    "/werkstatt.html": "Werkstatt",
    "/wifi-qr": "wifi-qr",
    "/shrinkling": "shrinkling",
    "/bigday": "bigday",
    "/dreh-das-rad": "Dreh das Rad",
    "/collective-calc": "Collective-Calc",
    "/zaehlwerk": "Zählwerk",
    "/fontART-demo": "fontART",
    "/fontart-demo": "fontART",
}

# ------------------------------ Config --------------------------------------


def _encrypt_token(plain: str) -> str | None:
    """DPAPI-Blob als base64. None, wenn kein pywin32 da ist."""
    if not HAVE_DPAPI:
        return None
    try:
        blob = win32crypt.CryptProtectData(
            plain.encode("utf-8"), "BesucherTicker", None, None, None, 0)
        return base64.b64encode(blob).decode("ascii")
    except Exception:
        return None


def _decrypt_token(blob_b64: str) -> str | None:
    if not HAVE_DPAPI:
        return None
    try:
        blob = base64.b64decode(blob_b64.encode("ascii"))
        return win32crypt.CryptUnprotectData(
            blob, None, None, None, 0)[1].decode("utf-8")
    except Exception:
        return None


def load_config() -> dict:
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(cfg, dict):
        return {}
    enc = cfg.get("api_token_enc")
    if enc:
        plain = _decrypt_token(enc)
        if plain:
            cfg["api_token"] = plain
        elif not cfg.get("api_token"):
            # Blob da, aber nicht entschlüsselbar (anderer Rechner/Benutzer
            # oder pywin32 fehlt) -> Token gilt als nicht vorhanden.
            cfg.pop("api_token", None)
    return cfg


def save_config(cfg: dict) -> None:
    """Atomar schreiben: erst Temp-Datei, dann umbenennen. Ein Absturz
    mitten im Schreiben kann die config.json (samt Token) so nicht
    zerlegen."""
    out = dict(cfg)
    token = out.pop("api_token", None)
    out.pop("api_token_enc", None)
    if token:
        enc = _encrypt_token(token)
        if enc:
            out["api_token_enc"] = enc
        else:
            out["api_token"] = token
    tmp = CONFIG_PATH + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2, ensure_ascii=False)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, CONFIG_PATH)
    except OSError:
        try:
            os.unlink(tmp)
        except OSError:
            pass


# --------------------- Instanz-Sperre & Bildschirmgrenzen -------------------

_LOCK_HANDLE = None


def acquire_single_instance() -> bool:
    """True, wenn wir die einzige Instanz sind. Ohne das überschreiben
    sich zwei Ticker gegenseitig die config.json."""
    global _LOCK_HANDLE
    if msvcrt is None:
        return True
    try:
        _LOCK_HANDLE = open(LOCK_PATH, "a+b")
        msvcrt.locking(_LOCK_HANDLE.fileno(), msvcrt.LK_NBLCK, 1)
        return True
    except OSError:
        if _LOCK_HANDLE is not None:
            try:
                _LOCK_HANDLE.close()
            except OSError:
                pass
            _LOCK_HANDLE = None
        return False


def release_single_instance() -> None:
    global _LOCK_HANDLE
    if _LOCK_HANDLE is None:
        return
    try:
        msvcrt.locking(_LOCK_HANDLE.fileno(), msvcrt.LK_UNLCK, 1)
    except OSError:
        pass
    try:
        _LOCK_HANDLE.close()
    except OSError:
        pass
    _LOCK_HANDLE = None


def virtual_screen() -> tuple[int, int, int, int]:
    """(x, y, breite, höhe) über ALLE Monitore. tkinter kennt nur den
    primären — bei zwei Bildschirmen wäre die Prüfung sonst falsch."""
    try:
        import ctypes
        g = ctypes.windll.user32.GetSystemMetrics
        return g(76), g(77), g(78), g(79)
    except (AttributeError, OSError, ImportError):
        return 0, 0, 1920, 1080


def clamp_to_screen(x: int, y: int, w: int = 240, h: int = 140) -> tuple[int, int]:
    """Holt eine Position zurück auf sichtbares Gebiet. Nötig, weil das
    Overlay randlos ist: liegt es ausserhalb, kommt man ohne Neustart
    nicht mehr dran — es gibt keine Titelleiste."""
    vx, vy, vw, vh = virtual_screen()
    margin = 40
    if (x + w < vx + margin or x > vx + vw - margin
            or y + h < vy + margin or y > vy + vh - margin):
        return vx + 60, vy + 60
    return x, y


# ------------------------------ Zugangsdaten --------------------------------


def normalisiere_site(eingabe: str) -> str:
    """Nimmt 'dennis-mit-2n', 'dennis-mit-2n.goatcounter.com' oder die
    volle URL entgegen und liefert immer die Basis-URL ohne Schrägstrich
    am Ende."""
    s = (eingabe or "").strip()
    if not s:
        return ""
    s = s.rstrip("/")
    if s.startswith("http://"):
        s = "https://" + s[len("http://"):]
    if not s.startswith("https://"):
        if "." not in s:
            s = f"{s}.goatcounter.com"
        s = "https://" + s
    return s


def ask_site(parent: tk.Tk | None = None, vorgabe: str = "") -> str | None:
    own_root = False
    if parent is None:
        parent = tk.Tk()
        parent.withdraw()
        own_root = True
    wert = simpledialog.askstring(
        "Besucher-Ticker",
        "Adresse der GoatCounter-Seite:\n"
        "(z. B. dennis-mit-2n  oder  dennis-mit-2n.goatcounter.com)",
        initialvalue=vorgabe,
        parent=parent,
    )
    if own_root:
        parent.destroy()
    return normalisiere_site(wert) or None


def ask_token(parent: tk.Tk | None = None) -> str | None:
    """GUI-Dialog statt input() — funktioniert auch unter pythonw.exe."""
    own_root = False
    if parent is None:
        parent = tk.Tk()
        parent.withdraw()
        own_root = True
    token = simpledialog.askstring(
        "Besucher-Ticker",
        "API-Token einfügen\n"
        "(GoatCounter → Benutzername oben rechts → API →\n"
        "\"Add new token\", Recht \"Read statistics\" genügt):",
        parent=parent,
    )
    if own_root:
        parent.destroy()
    if token:
        token = token.strip()
    return token or None


# ------------------------------ API -----------------------------------------


class RateLimited(Exception):
    """HTTP 429. Trägt die vom Server gewünschte Wartezeit mit sich."""

    def __init__(self, retry_after: float | None = None):
        super().__init__("Rate-Limit erreicht (429)")
        self.retry_after = retry_after


class GoatCounterAPI:
    def __init__(self, site: str, token: str):
        self.session = requests.Session()
        self.site = ""
        self.set_zugang(site, token)

    def set_zugang(self, site: str, token: str) -> None:
        self.site = normalisiere_site(site)
        self.session.headers.update({
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": UA,
        })

    def _get(self, pfad: str, params: dict | None = None):
        url = f"{self.site}/api/v0{pfad}"
        r = self.session.get(url, params=params or {}, timeout=20)
        if r.status_code in (401, 403):
            raise PermissionError("Token ungültig oder ohne Leserecht")
        if r.status_code == 404:
            raise LookupError(f"Nicht gefunden: {pfad}")
        if r.status_code == 429:
            # GoatCounter meldet die Wartezeit in einem eigenen Kopffeld,
            # Retry-After ist der Standardweg. Beides prüfen.
            roh = (r.headers.get("X-Rate-Limit-Reset")
                   or r.headers.get("Retry-After"))
            secs = None
            if roh:
                try:
                    secs = min(float(roh), RATE_LIMIT_MAX_WAIT)
                except ValueError:
                    secs = None
            raise RateLimited(secs)
        r.raise_for_status()
        return r.json()

    def total(self, start: str, end: str) -> dict:
        return self._get("/stats/total", {"start": start, "end": end})

    def hits(self, start: str, end: str, limit: int = 100) -> dict:
        return self._get("/stats/hits",
                         {"start": start, "end": end, "limit": limit})

    def seite(self, page: str, start: str, end: str, limit: int = 10) -> dict:
        return self._get(f"/stats/{page}",
                         {"start": start, "end": end, "limit": limit})


# ------------------------------ Zeit & Parsing ------------------------------


def utc_stempel(dt: datetime) -> str:
    """GoatCounter will UTC, auf die volle Stunde gerundet."""
    dt = dt.astimezone(timezone.utc).replace(minute=0, second=0, microsecond=0)
    return dt.strftime("%Y-%m-%dT%H:%M:%SZ")


def zeitfenster(schluessel: str, jetzt: datetime | None = None) -> tuple[str, str]:
    """(start, ende) für einen Zeitraum-Schlüssel, in UTC.

    'heute' meint den lokalen Tag — mitteleuropäische Mitternacht ist
    in UTC eine volle Stunde, die Rundung verliert also nichts."""
    jetzt = jetzt or datetime.now().astimezone()
    ende = jetzt + timedelta(hours=1)          # laufende Stunde mitnehmen
    tage = dict((k, t) for k, _, t in ZEITRAEUME).get(schluessel, 0)
    if tage is None:
        return ALLTIME_START, utc_stempel(ende)
    tagesbeginn = jetzt.replace(hour=0, minute=0, second=0, microsecond=0)
    start = tagesbeginn - timedelta(days=tage)
    return utc_stempel(start), utc_stempel(ende)


def _zahl(v) -> int:
    try:
        return int(v)
    except (TypeError, ValueError):
        return 0


def pfad_label(pfad: str) -> str:
    """Ordnet einen Pfad einem Werkzeug zu.

    Alle Auftritte liegen unter derselben GoatCounter-Seite, unter-
    scheiden sich also nur im ersten Pfadabschnitt. Unterseiten eines
    Werkzeugs sollen nicht als eigener Balken erscheinen."""
    p = (pfad or "/").strip()
    if not p.startswith("/"):
        p = "/" + p
    if p in PFAD_NAMEN:
        return PFAD_NAMEN[p]
    erster = "/" + p.lstrip("/").split("/", 1)[0]
    if erster in PFAD_NAMEN:
        return PFAD_NAMEN[erster]
    if erster in ("/", ""):
        return PFAD_NAMEN["/"]
    return erster.lstrip("/") or "Startseite"


def auswerten_hits(payload: dict, max_balken: int) -> list[dict]:
    """Fasst die Pfadliste zu Werkzeug-Balken zusammen.

    Ereignisse (event: true) fliegen raus — sie sind die data-count-
    Klicks der Werkstatt und würden sonst neben dem Seitenaufruf ein
    zweites Mal zählen."""
    summen: dict[str, int] = {}
    for hit in (payload or {}).get("hits") or []:
        if not isinstance(hit, dict):
            continue
        if hit.get("event"):
            continue
        label = pfad_label(hit.get("path", "/"))
        summen[label] = summen.get(label, 0) + _zahl(hit.get("count"))
    reihen = [{"label": k, "count": v} for k, v in summen.items() if v > 0]
    reihen.sort(key=lambda r: (-r["count"], r["label"]))
    return reihen[:max_balken]


def auswerten_stats(payload: dict, max_zeilen: int = 4,
                    leer_label: str = "(unbekannt)") -> list[dict]:
    """Länder- oder Verweisliste in eine schlichte Reihenfolge bringen.

    Ein leerer Name heisst bei den Verweisen etwas anderes als bei den
    Ländern: dort sind es Leute, die die Adresse direkt eingegeben oder
    ein Lesezeichen benutzt haben — nicht "unbekannt", sondern "direkt".
    Deshalb kommt die Beschriftung von aussen."""
    reihen = []
    for s in (payload or {}).get("stats") or []:
        if not isinstance(s, dict):
            continue
        name = (s.get("name") or "").strip() or leer_label
        reihen.append({"label": name, "count": _zahl(s.get("count"))})
    reihen.sort(key=lambda r: -r["count"])
    return reihen[:max_zeilen]


def zahl_kurz(n: int) -> str:
    """Tausenderpunkte deutscher Schreibweise, ohne locale-Abhängigkeit."""
    return f"{n:,}".replace(",", ".")


# ------------------------------ App -----------------------------------------


class TickerApp:
    def __init__(self):
        self.cfg = load_config()

        site = self.cfg.get("site")
        if not site:
            site = ask_site()
            if not site:
                sys.exit(0)
            self.cfg["site"] = site
        token = self.cfg.get("api_token")
        if not token:
            token = ask_token()
            if not token:
                sys.exit(0)
            self.cfg["api_token"] = token
        save_config(self.cfg)

        self.refresh_seconds = float(self.cfg.get("refresh_seconds", DEFAULT_REFRESH))
        self.alpha_pct = int(self.cfg.get("alpha_pct", 88))
        self.theme_name = self.cfg.get("theme", DEFAULT_THEME)
        if self.theme_name not in THEMES:
            self.theme_name = DEFAULT_THEME
        self.compact_duration = self.cfg.get("compact_duration", None)
        self.zeitraum = self.cfg.get("zeitraum", DEFAULT_ZEITRAUM)
        if self.zeitraum not in [k for k, _, _ in ZEITRAEUME]:
            self.zeitraum = DEFAULT_ZEITRAUM
        self.zusatz = self.cfg.get("zusatz", DEFAULT_ZUSATZ)
        if self.zusatz not in [k for k, _, _ in ZUSATZ_CHOICES]:
            self.zusatz = DEFAULT_ZUSATZ
        self.max_balken = int(self.cfg.get("max_balken", DEFAULT_MAX_BALKEN))

        self.api = GoatCounterAPI(self.cfg["site"], self.cfg["api_token"])
        self.ui_queue: queue.Queue = queue.Queue()
        self.stop_event = threading.Event()
        self.wake_event = threading.Event()
        self._refresh_lock = threading.Lock()
        self._rate_limit_hits = 0

        # Gesamtsumme ändert sich träge und kostet eine eigene Anfrage —
        # deshalb gecacht statt bei jedem Durchlauf neu geholt.
        self._alltime_wert: int | None = None
        self._alltime_zeit = 0.0

        self.compact = bool(self.cfg.get("compact", False))
        self.visible = True
        self.daten: dict = {}
        self.icon: pystray.Icon | None = None
        self.settings_win: tk.Toplevel | None = None
        self._compact_revert_job: str | None = None

        self._build_ui()
        self._build_tray()

        self.worker = threading.Thread(target=self._refresh_loop, daemon=True)
        self.worker.start()
        self.root.after(200, self._process_queue)

    # ---------- Theme-Zugriff ----------

    @property
    def th(self) -> dict:
        return THEMES[self.theme_name]

    def zeitraum_name(self) -> str:
        return next((n for k, n, _ in ZEITRAEUME if k == self.zeitraum), "Heute")

    def zusatz_page(self) -> str | None:
        return next((p for k, _, p in ZUSATZ_CHOICES if k == self.zusatz), None)

    def zusatz_name(self) -> str:
        return next((n for k, n, _ in ZUSATZ_CHOICES if k == self.zusatz), "aus")

    # ---------- UI ----------

    def _build_ui(self):
        self.root = tk.Tk()
        self.root.title("Besucher")
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.attributes("-alpha", self.alpha_pct / 100.0)
        self.root.configure(bg=self.th["bg"])

        x, y = clamp_to_screen(int(self.cfg.get("pos_x", 60)),
                               int(self.cfg.get("pos_y", 60)))
        self.cfg["pos_x"], self.cfg["pos_y"] = x, y
        self.root.geometry(f"+{x}+{y}")

        self.frame = tk.Frame(self.root, bg=self.th["bg"], padx=10, pady=8)
        self.frame.pack()

        self.title_lbl = tk.Label(self.frame, text="Besucher",
                                  fg=self.th["title"], bg=self.th["bg"],
                                  font=("Segoe UI", 9, "bold"))
        self.title_lbl.pack(anchor="w")

        self.body = tk.Frame(self.frame, bg=self.th["bg"])
        self.body.pack(fill="x")

        self.status_lbl = tk.Label(self.frame, text="lade…",
                                   fg=self.th["muted"], bg=self.th["bg"],
                                   font=("Segoe UI", 8))
        self.status_lbl.pack(anchor="w")

        self._bind_drag_recursive(self.root)
        self.root.protocol("WM_DELETE_WINDOW", self.quit)

    def _bind_drag_recursive(self, widget):
        widget.bind("<Button-1>", self._drag_start)
        widget.bind("<B1-Motion>", self._drag_move)
        widget.bind("<ButtonRelease-1>", self._drag_end)
        widget.bind("<Double-Button-1>", self._toggle_compact)
        for child in widget.winfo_children():
            self._bind_drag_recursive(child)

    def _drag_start(self, event):
        self._drag_off = (event.x_root - self.root.winfo_x(),
                          event.y_root - self.root.winfo_y())

    def _drag_move(self, event):
        ox, oy = getattr(self, "_drag_off", (0, 0))
        x, y = event.x_root - ox, event.y_root - oy
        self.root.geometry(f"+{x}+{y}")
        # Position nur merken, NICHT speichern: sonst schriebe jede
        # einzelne Mausbewegung die config.json (mit dem Token darin).
        self.cfg["pos_x"], self.cfg["pos_y"] = x, y

    def _drag_end(self, _event=None):
        x, y = clamp_to_screen(int(self.cfg.get("pos_x", 60)),
                               int(self.cfg.get("pos_y", 60)))
        if (x, y) != (self.cfg.get("pos_x"), self.cfg.get("pos_y")):
            self.root.geometry(f"+{x}+{y}")
            self.cfg["pos_x"], self.cfg["pos_y"] = x, y
        save_config(self.cfg)

    def _center_overlay(self):
        vx, vy, vw, vh = virtual_screen()
        x, y = vx + vw // 2 - 120, vy + vh // 2 - 70
        self.root.geometry(f"+{x}+{y}")
        self.cfg["pos_x"], self.cfg["pos_y"] = x, y
        save_config(self.cfg)
        if not self.visible:
            self._toggle_visible()

    # ---------- Mini/Maxi ----------

    def _toggle_compact(self, _event=None):
        self._set_compact(not self.compact, from_user=True)

    def _set_compact(self, value: bool, from_user: bool = False):
        if self._compact_revert_job:
            try:
                self.root.after_cancel(self._compact_revert_job)
            except (tk.TclError, ValueError):
                pass
            self._compact_revert_job = None

        self.compact = value
        self.cfg["compact"] = self.compact
        save_config(self.cfg)
        self._render()

        if from_user and self.compact_duration:
            ms = int(self.compact_duration * 1000)
            self._compact_revert_job = self.root.after(
                ms, lambda: self._set_compact(not value, from_user=False))

    # ---------- Rendering ----------

    def apply_theme(self):
        t = self.th
        self.root.configure(bg=t["bg"])
        self.frame.configure(bg=t["bg"])
        self.body.configure(bg=t["bg"])
        self.title_lbl.configure(fg=t["title"], bg=t["bg"])
        self.status_lbl.configure(fg=t["muted"], bg=t["bg"])
        self._render()

    def _render(self):
        t = self.th
        for w in self.body.winfo_children():
            w.destroy()

        self.title_lbl.config(text=f"Besucher · {self.zeitraum_name()}")

        if not self.daten:
            self._bind_drag_recursive(self.root)
            return

        heute = self.daten.get("heute", 0)
        gesamt = self.daten.get("gesamt")
        zeitraum_total = self.daten.get("zeitraum_total", 0)

        if self.compact:
            tk.Label(self.body, text=zahl_kurz(zeitraum_total),
                     fg=t["big"], bg=t["bg"],
                     font=("Segoe UI", 16, "bold")).pack(anchor="w")
            tk.Label(self.body, text=f"Aufrufe · {self.zeitraum_name()}",
                     fg=t["muted"], bg=t["bg"],
                     font=("Segoe UI", 8)).pack(anchor="w")
            self._bind_drag_recursive(self.root)
            return

        # --- Kopfzahlen ---
        # Die grosse Zahl gehoert zum gewaehlten Zeitraum. Stuende dort
        # immer "heute", zeigte das Overlay nachts um zwei eine grosse
        # Null, obwohl "7 Tage" eingestellt ist.
        kopf = tk.Frame(self.body, bg=t["bg"])
        kopf.pack(fill="x", pady=(2, 0))
        tk.Label(kopf, text=zahl_kurz(zeitraum_total), fg=t["big"], bg=t["bg"],
                 font=("Segoe UI", 18, "bold")).pack(side="left")
        rechts = tk.Frame(kopf, bg=t["bg"])
        rechts.pack(side="right", anchor="s")
        tk.Label(rechts, text=f"Aufrufe · {self.zeitraum_name()}",
                 fg=t["muted"], bg=t["bg"],
                 font=("Segoe UI", 8)).pack(anchor="e")
        if self.zeitraum != "heute":
            tk.Label(rechts, text=f"{zahl_kurz(heute)} heute",
                     fg=t["text"], bg=t["bg"],
                     font=("Segoe UI", 9, "bold")).pack(anchor="e")
        if gesamt is not None:
            tk.Label(rechts, text=f"{zahl_kurz(gesamt)} gesamt",
                     fg=t["text"], bg=t["bg"],
                     font=("Segoe UI", 9, "bold")).pack(anchor="e")

        # --- Balken je Werkzeug ---
        balken = self.daten.get("balken") or []
        if balken:
            tk.Label(self.body, text="Besucher je Werkzeug",
                     fg=t["muted"], bg=t["bg"],
                     font=("Segoe UI", 8)).pack(anchor="w", pady=(6, 0))
            hoechst = max(b["count"] for b in balken) or 1
            for b in balken:
                row = tk.Frame(self.body, bg=t["bg"])
                row.pack(fill="x", pady=(3, 0))
                head = tk.Frame(row, bg=t["bg"])
                head.pack(fill="x")
                tk.Label(head, text=b["label"], fg=t["text"], bg=t["bg"],
                         font=("Segoe UI", 9)).pack(side="left")
                tk.Label(head, text=zahl_kurz(b["count"]), fg=t["text"],
                         bg=t["bg"],
                         font=("Segoe UI", 9, "bold")).pack(side="right")
                bar_bg = tk.Frame(row, bg=t["bar_bg"], width=BAR_W, height=BAR_H)
                bar_bg.pack(fill="x", pady=(1, 0))
                bar_bg.pack_propagate(False)
                fill_w = max(2, int(BAR_W * b["count"] / hoechst))
                tk.Frame(bar_bg, bg=t["bar"], width=fill_w,
                         height=BAR_H).place(x=0, y=0)

        # --- Zusatzblock ---
        zusatz = self.daten.get("zusatz") or []
        if zusatz and self.zusatz_page():
            tk.Label(self.body, text=self.zusatz_name(), fg=t["muted"],
                     bg=t["bg"], font=("Segoe UI", 8)).pack(anchor="w",
                                                            pady=(6, 0))
            for z in zusatz:
                row = tk.Frame(self.body, bg=t["bg"])
                row.pack(fill="x")
                tk.Label(row, text=z["label"][:22], fg=t["text"], bg=t["bg"],
                         font=("Segoe UI", 8)).pack(side="left")
                tk.Label(row, text=zahl_kurz(z["count"]), fg=t["muted"],
                         bg=t["bg"], font=("Segoe UI", 8)).pack(side="right")

        self._bind_drag_recursive(self.root)

    # ---------- Tray ----------

    def _tray_image(self) -> Image.Image:
        t = self.th
        bg_rgb = tuple(int(t["bg"][i:i+2], 16) for i in (1, 3, 5)) + (255,)
        img = Image.new("RGBA", (64, 64), bg_rgb)
        d = ImageDraw.Draw(img)
        balken = self.daten.get("balken") or []
        hoechst = max((b["count"] for b in balken), default=0) or 1
        # Drei Säulen als Sinnbild für die Statistik, in Blau.
        for i, b in enumerate(balken[:3]):
            h = max(4, int(38 * b["count"] / hoechst))
            x0 = 10 + i * 16
            d.rectangle([x0, 54 - h, x0 + 11, 54], fill=t["bar"])
        if not balken:
            d.rectangle([10, 44, 54, 54], fill=t["bar_bg"])
        d.ellipse([44, 6, 58, 20], fill=t["title"])
        return img

    def _build_tray(self):
        zeitraum_menu = pystray.Menu(*[
            pystray.MenuItem(
                name,
                (lambda k: (lambda *_: self.ui_queue.put(("zeitraum", k))))(key),
                checked=(lambda k: (lambda _i: self.zeitraum == k))(key),
                radio=True)
            for key, name, _ in ZEITRAEUME
        ])
        zusatz_menu = pystray.Menu(*[
            pystray.MenuItem(
                name,
                (lambda k: (lambda *_: self.ui_queue.put(("zusatz", k))))(key),
                checked=(lambda k: (lambda _i: self.zusatz == k))(key),
                radio=True)
            for key, name, _ in ZUSATZ_CHOICES
        ])
        menu = pystray.Menu(
            pystray.MenuItem("Overlay an/aus", self._tray_toggle, default=True),
            pystray.MenuItem("Overlay zur Mitte holen", self._tray_center),
            pystray.MenuItem("Jetzt aktualisieren", self._tray_refresh),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Zeitraum", zeitraum_menu),
            pystray.MenuItem("Zusatzblock", zusatz_menu),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Einstellungen…", self._tray_settings),
            pystray.MenuItem("Nächstes Theme", self._tray_next_theme),
            pystray.MenuItem("Zugang ändern", self._tray_change_zugang),
            pystray.MenuItem("Beenden", self._tray_quit),
        )
        self.icon = pystray.Icon("besucher_ticker", self._tray_image(),
                                 "Besucher-Ticker", menu)
        threading.Thread(target=self.icon.run, daemon=True).start()

    def _tray_toggle(self, *_):
        self.ui_queue.put(("toggle", None))

    def _tray_center(self, *_):
        self.ui_queue.put(("center", None))

    def _tray_refresh(self, *_):
        threading.Thread(target=self._refresh_once, daemon=True).start()

    def _tray_settings(self, *_):
        self.ui_queue.put(("settings", None))

    def _tray_next_theme(self, *_):
        self.ui_queue.put(("next_theme", None))

    def _tray_change_zugang(self, *_):
        self.ui_queue.put(("change_zugang", None))

    def _tray_quit(self, *_):
        self.ui_queue.put(("quit", None))

    # ---------- Einstellungsfenster ----------

    def _open_settings(self):
        if self.settings_win is not None and self.settings_win.winfo_exists():
            self.settings_win.lift()
            self.settings_win.focus_force()
            return

        t = self.th
        win = tk.Toplevel(self.root)
        self.settings_win = win
        win.title("Besucher-Ticker — Einstellungen")
        win.configure(bg=t["bg"], padx=14, pady=12)
        win.attributes("-topmost", True)
        win.resizable(False, False)
        win.geometry(f"+{self.root.winfo_x() + 40}+{self.root.winfo_y() + 40}")

        def lbl(text, row, pady=(8, 2)):
            tk.Label(win, text=text, fg=t["title"], bg=t["bg"],
                     font=("Segoe UI", 9, "bold")).grid(
                row=row, column=0, columnspan=2, sticky="w", pady=pady)

        style = ttk.Style(win)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        # --- Update-Frequenz ---
        lbl("Update-Frequenz", 0, pady=(0, 2))
        self._refresh_var = tk.StringVar()
        current_label = next(
            (name for name, sec in REFRESH_CHOICES
             if abs(sec - self.refresh_seconds) < 0.01),
            f"{self.refresh_seconds:g} s")
        self._refresh_var.set(current_label)
        cb_refresh = ttk.Combobox(win, textvariable=self._refresh_var,
                                  values=[n for n, _ in REFRESH_CHOICES],
                                  state="readonly", width=30)
        cb_refresh.grid(row=1, column=0, columnspan=2, sticky="we")

        def on_refresh(_e=None):
            sel = self._refresh_var.get()
            for name, sec in REFRESH_CHOICES:
                if name == sel:
                    self.refresh_seconds = float(sec)
                    self.cfg["refresh_seconds"] = self.refresh_seconds
                    save_config(self.cfg)
                    self.wake_event.set()
                    break
        cb_refresh.bind("<<ComboboxSelected>>", on_refresh)

        tk.Label(win, text="GoatCounter erlaubt 4 Anfragen je Sekunde —"
                           "\nselbst 15 s sind unkritisch.",
                 fg=t["muted"], bg=t["bg"], font=("Segoe UI", 8),
                 justify="left").grid(row=2, column=0, columnspan=2, sticky="w")

        # --- Zeitraum ---
        lbl("Zeitraum für Balken und Zwischensumme", 3)
        self._zr_var = tk.StringVar(value=self.zeitraum_name())
        cb_zr = ttk.Combobox(win, textvariable=self._zr_var,
                             values=[n for _, n, _ in ZEITRAEUME],
                             state="readonly", width=30)
        cb_zr.grid(row=4, column=0, columnspan=2, sticky="we")

        def on_zr(_e=None):
            sel = self._zr_var.get()
            for key, name, _ in ZEITRAEUME:
                if name == sel:
                    self._set_zeitraum(key)
                    break
        cb_zr.bind("<<ComboboxSelected>>", on_zr)

        # --- Zusatzblock ---
        lbl("Zusatzblock unter den Balken", 5)
        self._zus_var = tk.StringVar(value=self.zusatz_name())
        cb_zus = ttk.Combobox(win, textvariable=self._zus_var,
                              values=[n for _, n, _ in ZUSATZ_CHOICES],
                              state="readonly", width=30)
        cb_zus.grid(row=6, column=0, columnspan=2, sticky="we")

        def on_zus(_e=None):
            sel = self._zus_var.get()
            for key, name, _ in ZUSATZ_CHOICES:
                if name == sel:
                    self._set_zusatz(key)
                    break
        cb_zus.bind("<<ComboboxSelected>>", on_zus)

        # --- Anzahl Balken ---
        lbl("Höchstens so viele Balken", 7)
        self._bal_var = tk.StringVar(value=str(self.max_balken))
        cb_bal = ttk.Combobox(win, textvariable=self._bal_var,
                              values=[str(n) for n in MAX_BALKEN_CHOICES],
                              state="readonly", width=30)
        cb_bal.grid(row=8, column=0, columnspan=2, sticky="we")

        def on_bal(_e=None):
            try:
                self.max_balken = int(self._bal_var.get())
            except ValueError:
                return
            self.cfg["max_balken"] = self.max_balken
            save_config(self.cfg)
            self.wake_event.set()
        cb_bal.bind("<<ComboboxSelected>>", on_bal)

        # --- Transparenz ---
        lbl(f"Transparenz ({ALPHA_MIN}–{ALPHA_MAX} %)", 9)
        self._alpha_var = tk.IntVar(value=self.alpha_pct)

        def on_alpha(_v=None):
            self.alpha_pct = int(self._alpha_var.get())
            self.root.attributes("-alpha", self.alpha_pct / 100.0)
            self.cfg["alpha_pct"] = self.alpha_pct
            save_config(self.cfg)

        tk.Scale(win, from_=ALPHA_MIN, to=ALPHA_MAX, orient="horizontal",
                 variable=self._alpha_var, command=on_alpha, bg=t["bg"],
                 fg=t["text"], troughcolor=t["bar_bg"], highlightthickness=0,
                 length=230).grid(row=10, column=0, columnspan=2, sticky="we")

        # --- Mini-/Maxi-Modus ---
        lbl("Mini-/Maxi-Modus", 11)
        btn_frame = tk.Frame(win, bg=t["bg"])
        btn_frame.grid(row=12, column=0, columnspan=2, sticky="w")
        tk.Button(btn_frame, text="Mini", width=8,
                  command=lambda: self._set_compact(True, from_user=True)
                  ).pack(side="left", padx=(0, 6))
        tk.Button(btn_frame, text="Maxi", width=8,
                  command=lambda: self._set_compact(False, from_user=True)
                  ).pack(side="left")

        tk.Label(win, text="Nach Umschalten zurück nach:", fg=t["text"],
                 bg=t["bg"], font=("Segoe UI", 8)).grid(
            row=13, column=0, columnspan=2, sticky="w", pady=(6, 2))
        self._dur_var = tk.StringVar(
            value=next((n for n, s in COMPACT_DURATIONS
                        if s == self.compact_duration), "dauerhaft"))
        cb_dur = ttk.Combobox(win, textvariable=self._dur_var,
                              values=[n for n, _ in COMPACT_DURATIONS],
                              state="readonly", width=30)
        cb_dur.grid(row=14, column=0, columnspan=2, sticky="we")

        def on_dur(_e=None):
            sel = self._dur_var.get()
            for name, sec in COMPACT_DURATIONS:
                if name == sel:
                    self.compact_duration = sec
                    self.cfg["compact_duration"] = sec
                    save_config(self.cfg)
                    break
        cb_dur.bind("<<ComboboxSelected>>", on_dur)

        # --- Theme ---
        lbl(f"Farben ({len(THEME_NAMES)} Varianten)", 15)
        self._theme_var = tk.StringVar(value=self.theme_name)
        cb_theme = ttk.Combobox(win, textvariable=self._theme_var,
                                values=THEME_NAMES, state="readonly", width=30)
        cb_theme.grid(row=16, column=0, columnspan=2, sticky="we")

        def set_theme(name: str):
            if name in THEMES:
                self.theme_name = name
                self.cfg["theme"] = name
                save_config(self.cfg)
                self.apply_theme()
                if self.icon:
                    self.icon.icon = self._tray_image()
                # Einstellungsfenster live mitfärben
                win.destroy()
                self.settings_win = None
                self._open_settings()

        cb_theme.bind("<<ComboboxSelected>>",
                      lambda _e: set_theme(self._theme_var.get()))

        tk.Button(win, text="Schließen", command=win.destroy).grid(
            row=17, column=0, columnspan=2, pady=(12, 0))

        def on_close():
            self.settings_win = None
            win.destroy()
        win.protocol("WM_DELETE_WINDOW", on_close)

    def _set_zeitraum(self, key: str):
        self.zeitraum = key
        self.cfg["zeitraum"] = key
        save_config(self.cfg)
        self._render()
        threading.Thread(target=self._refresh_once, daemon=True).start()

    def _set_zusatz(self, key: str):
        self.zusatz = key
        self.cfg["zusatz"] = key
        save_config(self.cfg)
        self._render()
        threading.Thread(target=self._refresh_once, daemon=True).start()

    def _next_theme(self):
        idx = (THEME_NAMES.index(self.theme_name) + 1) % len(THEME_NAMES)
        self.theme_name = THEME_NAMES[idx]
        self.cfg["theme"] = self.theme_name
        save_config(self.cfg)
        self.apply_theme()
        if self.icon:
            self.icon.icon = self._tray_image()

    # ---------- Worker ----------

    def _refresh_loop(self):
        backoff_idx = 0
        while not self.stop_event.is_set():
            ok, wait_hint = self._refresh_once()
            if ok:
                backoff_idx = 0
                self._rate_limit_hits = 0
                wait = self.refresh_seconds
            elif wait_hint is not None:
                idx = min(self._rate_limit_hits, len(RATE_LIMIT_BACKOFF) - 1)
                wait = max(wait_hint, RATE_LIMIT_BACKOFF[idx],
                           self.refresh_seconds)
                self._rate_limit_hits += 1
            else:
                wait = max(RETRY_BACKOFF[min(backoff_idx, len(RETRY_BACKOFF) - 1)],
                           self.refresh_seconds)
                backoff_idx += 1
            deadline = time.monotonic() + wait
            while not self.stop_event.is_set():
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    break
                if self.wake_event.wait(timeout=min(0.25, remaining)):
                    self.wake_event.clear()
                    break

    def _refresh_once(self) -> tuple[bool, float | None]:
        """(erfolgreich, gewünschte Mindestwartezeit). Der zweite Wert ist
        nur bei 429 gesetzt."""
        if not self._refresh_lock.acquire(blocking=False):
            return True, None
        try:
            jetzt = datetime.now().astimezone()
            heute_start, heute_ende = zeitfenster("heute", jetzt)
            zr_start, zr_ende = zeitfenster(self.zeitraum, jetzt)

            heute_total = _zahl(self.api.total(heute_start, heute_ende)
                                .get("total"))

            if self.zeitraum == "heute":
                zeitraum_total = heute_total
            else:
                zeitraum_total = _zahl(self.api.total(zr_start, zr_ende)
                                       .get("total"))

            balken = auswerten_hits(self.api.hits(zr_start, zr_ende, limit=100),
                                    self.max_balken)

            zusatz = []
            page = self.zusatz_page()
            if page:
                try:
                    zusatz = auswerten_stats(
                        self.api.seite(page, zr_start, zr_ende, limit=10),
                        leer_label=("(direkt)" if page == "toprefs"
                                    else "(unbekannt)"))
                except LookupError:
                    # Ältere GoatCounter-Fassungen kennen die Seite nicht.
                    # Kein Grund, den ganzen Durchlauf scheitern zu lassen.
                    zusatz = []

            gesamt = self._gesamt_holen(heute_ende)

            self.ui_queue.put(("data", {
                "heute": heute_total,
                "zeitraum_total": zeitraum_total,
                "balken": balken,
                "zusatz": zusatz,
                "gesamt": gesamt,
            }))
            return True, None
        except PermissionError:
            self.ui_queue.put(("auth_error", None))
            return False, None
        except RateLimited as e:
            self.ui_queue.put(("rate_limit", e.retry_after))
            return False, e.retry_after or RATE_LIMIT_BACKOFF[0]
        except (requests.RequestException, ValueError, LookupError) as e:
            self.ui_queue.put(("net_error", str(e)))
            return False, None
        finally:
            self._refresh_lock.release()

    def _gesamt_holen(self, ende: str) -> int | None:
        """Gesamtsumme, höchstens alle 10 Minuten frisch geholt."""
        if (self._alltime_wert is not None
                and time.monotonic() - self._alltime_zeit < ALLTIME_CACHE_SECONDS):
            return self._alltime_wert
        try:
            wert = _zahl(self.api.total(ALLTIME_START, ende).get("total"))
        except (requests.RequestException, ValueError, LookupError):
            return self._alltime_wert          # alter Wert ist besser als keiner
        self._alltime_wert = wert
        self._alltime_zeit = time.monotonic()
        return wert

    # ---------- Queue / Mainloop ----------

    def _process_queue(self):
        try:
            while True:
                kind, payload = self.ui_queue.get_nowait()
                if kind == "data":
                    self.daten = payload
                    self.status_lbl.config(
                        text=datetime.now().strftime("aktualisiert %H:%M:%S"),
                        fg=self.th["muted"])
                    self._render()
                    if self.icon:
                        self.icon.icon = self._tray_image()
                elif kind == "net_error":
                    self.status_lbl.config(
                        text="Netzwerkfehler – neuer Versuch…",
                        fg="#ff9f43")
                elif kind == "rate_limit":
                    txt = "Rate-Limit (429) – Pause"
                    if payload:
                        txt += f" {int(payload)}s"
                    self.status_lbl.config(text=txt, fg="#ff4d4d")
                elif kind == "auth_error":
                    self.status_lbl.config(text="Token ungültig!",
                                           fg="#ff4d4d")
                    self._change_zugang_dialog()
                elif kind == "change_zugang":
                    self._change_zugang_dialog()
                elif kind == "settings":
                    self._open_settings()
                elif kind == "next_theme":
                    self._next_theme()
                elif kind == "zeitraum":
                    self._set_zeitraum(payload)
                elif kind == "zusatz":
                    self._set_zusatz(payload)
                elif kind == "toggle":
                    self._toggle_visible()
                elif kind == "center":
                    self._center_overlay()
                elif kind == "quit":
                    self.quit()
                    return
        except queue.Empty:
            pass
        self.root.after(200, self._process_queue)

    def _change_zugang_dialog(self):
        site = ask_site(self.root, self.cfg.get("site", ""))
        if not site:
            return
        token = ask_token(self.root)
        if not token:
            return
        self.cfg["site"] = site
        self.cfg["api_token"] = token
        save_config(self.cfg)
        self.api.set_zugang(site, token)
        self._alltime_wert = None
        threading.Thread(target=self._refresh_once, daemon=True).start()

    def _toggle_visible(self):
        self.visible = not self.visible
        if self.visible:
            self.root.deiconify()
            self.root.attributes("-topmost", True)
        else:
            self.root.withdraw()

    def quit(self):
        self.stop_event.set()
        self.wake_event.set()
        if self.icon:
            try:
                self.icon.stop()
            except Exception:
                pass
        try:
            self.root.destroy()
        except tk.TclError:
            pass
        release_single_instance()

    def run(self):
        self.root.mainloop()


# ------------------------------ Selbsttest ----------------------------------


def selbsttest() -> int:
    """Prüft die Rechen- und Zuordnungslogik ohne Netz und ohne Fenster."""
    fehler: list[str] = []

    def pruefe(name, ist, soll):
        if ist != soll:
            fehler.append(f"{name}: {ist!r} statt {soll!r}")

    # Pfad-Zuordnung
    pruefe("pfad /", pfad_label("/"), "Startseite")
    pruefe("pfad werkstatt", pfad_label("/werkstatt.html"), "Werkstatt")
    pruefe("pfad wifi", pfad_label("/wifi-qr/"), "wifi-qr")
    pruefe("pfad unterseite", pfad_label("/zaehlwerk/hilfe.html"), "Zählwerk")
    pruefe("pfad unbekannt", pfad_label("/neues-tool/"), "neues-tool")

    # Ereignisse dürfen NICHT als Werkzeug-Balken erscheinen, sonst
    # zählte jeder Werkstatt-Klick doppelt.
    payload = {"hits": [
        {"path": "/wifi-qr/", "count": 10, "event": False},
        {"path": "/wifi-qr/hilfe", "count": 5, "event": False},
        {"path": "werkstatt-wifi-open", "count": 99, "event": True},
        {"path": "/bigday/", "count": 7, "event": False},
        {"path": "/leer/", "count": 0, "event": False},
    ]}
    balken = auswerten_hits(payload, 8)
    pruefe("balken anzahl", len(balken), 2)
    pruefe("balken erster", balken[0], {"label": "wifi-qr", "count": 15})
    pruefe("balken zweiter", balken[1], {"label": "bigday", "count": 7})

    # Begrenzung greift und sortiert nach Grösse
    viele = {"hits": [{"path": f"/t{i}/", "count": i, "event": False}
                      for i in range(1, 20)]}
    b2 = auswerten_hits(viele, 3)
    pruefe("balken grenze", [b["count"] for b in b2], [19, 18, 17])

    # Zusatzliste
    stats = {"stats": [{"name": "Germany", "count": 4},
                       {"name": "", "count": 9},
                       {"name": "France", "count": 1}]}
    z = auswerten_stats(stats, 2)
    pruefe("zusatz", [(x["label"], x["count"]) for x in z],
           [("(unbekannt)", 9), ("Germany", 4)])
    z2 = auswerten_stats(stats, 1, leer_label="(direkt)")
    pruefe("zusatz direkt", z2[0]["label"], "(direkt)")

    # Zeitfenster: 'heute' beginnt an lokaler Mitternacht, 7 Tage davor.
    jetzt = datetime(2026, 8, 1, 14, 30, tzinfo=timezone.utc)
    s0, e0 = zeitfenster("heute", jetzt)
    pruefe("heute start", s0, "2026-08-01T00:00:00Z")
    pruefe("heute ende", e0, "2026-08-01T15:00:00Z")
    s7, _ = zeitfenster("7t", jetzt)
    pruefe("7t start", s7, "2026-07-25T00:00:00Z")
    sg, _ = zeitfenster("gesamt", jetzt)
    pruefe("gesamt start", sg, ALLTIME_START)

    # Adress-Normalisierung
    for eingabe in ("dennis-mit-2n", "dennis-mit-2n.goatcounter.com",
                    "https://dennis-mit-2n.goatcounter.com/",
                    "http://dennis-mit-2n.goatcounter.com"):
        pruefe(f"site {eingabe}", normalisiere_site(eingabe),
               "https://dennis-mit-2n.goatcounter.com")

    pruefe("zahl", zahl_kurz(1234567), "1.234.567")

    if fehler:
        print("SELBSTTEST FEHLGESCHLAGEN:")
        for f in fehler:
            print("  -", f)
        return 1
    print("Selbsttest bestanden.")
    return 0


# ------------------------------ Main ----------------------------------------

if __name__ == "__main__":
    if "--selbsttest" in sys.argv:
        sys.exit(selbsttest())

    if not acquire_single_instance():
        _r = tk.Tk()
        _r.withdraw()
        messagebox.showinfo(
            "Besucher-Ticker",
            "Der Ticker läuft bereits.\n\n"
            "Das Fenster steckt im Infobereich der Taskleiste — "
            "Rechtsklick aufs Symbol.")
        _r.destroy()
        sys.exit(0)
    try:
        TickerApp().run()
    except KeyboardInterrupt:
        pass
    finally:
        release_single_instance()
