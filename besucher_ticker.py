"""
Besucher-Ticker v1.3.0 · GoatCounter-Overlay + System-Tray (Windows)
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
       (z. B. https://deinname.goatcounter.com)
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
      Jetzt aktualisieren, Zeitraum, Zusatzblock, Zählweise,
      Einstellungen, Nächstes Theme, Zugang ändern, Sprache, Beenden

Zu den Zahlen, was hier was bedeutet:
    GoatCounter speichert seit Langem keine einzelnen Aufrufe mehr,
    sondern Besucher je Pfad (Feld `count` in /stats/hits). Klickt
    dieselbe Person mehrmals auf dasselbe, zählt das einmal. Die API
    liefert in /stats/total die Summe über alle Pfade (`total`) und
    davon den Anteil der Ereignisse (`total_events`).

    Zwei Zählweisen, im Tray unter "Zählweise":
      * Seitenbesuche: grosse Zahl = total - total_events, die Balken
        sind die Seiten je Werkzeug. Ereignisse bleiben draussen.
      * Kachel-Klicks: nur Ereignisse, deren Pfad mit "kachel-"
        beginnt. So meldet die Werkstatt seit dem 12.11.2026 jeden
        Klick auf eine Kachel ("kachel-ram") oder eine Mini-Kachel
        darin ("kachel-ram-anleitung"), und sonst nichts. Grosse Zahl
        = Summe aller Kachel-Pfade, Balken = je Werkzeug zusammen-
        gefasst, beschriftet mit dem Titel, den die Seite mitschickt.
      * automatisch (Standard): Kachel-Klicks, sobald es auf der Seite
        je einen gab, sonst Seitenbesuche.
    Die Balken zeigen die grössten Werkzeuge; passen alle hinein,
    ergeben sie zusammen genau die grosse Zahl.

Datenschutz: Das Token bleibt ausschliesslich lokal (config.json neben
dem Skript). Es werden nur Lesezugriffe auf die eigene Seite gemacht.
"""

from __future__ import annotations

import base64
import ctypes
import json
import math
import os
import queue
import sys
import threading
import time
import tkinter as tk
from tkinter import simpledialog, ttk
import tkinter.font as tkfont
from datetime import datetime, timedelta, timezone

try:
    import msvcrt                       # nur Windows – für die Instanz-Sperre
except ImportError:
    msvcrt = None

try:
    import requests
    import pystray
    from PIL import Image, ImageChops, ImageDraw, ImageFilter
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
UA = "BesucherTicker/1.3.0 (personal read-only GoatCounter display)"

BAR_W, BAR_H = 190, 8

# GoatCounter erlaubt 4 Anfragen pro Sekunde. Ein Durchlauf braucht
# 2–5 Anfragen (Kachel-Klicks: eine je 100 Pfade mehr, höchstens
# HITS_SEITEN_MAX je Liste), die nacheinander
# laufen, mit mindestens API_ABSTAND Sekunden dazwischen. Nacheinander
# allein reicht nicht: Antwortet der Server in 0,1 s, wären das sonst
# zehn Anfragen je Sekunde. Trotzdem keine Sekundenstufen:
# es gibt keinen Grund, fremde Server im Sekundentakt zu fragen.
API_ABSTAND = 0.26
# Die Beschriftungen entstehen erst bei der Anzeige (refresh_name,
# dauer_name), damit sie der gewählten Sprache folgen.
REFRESH_CHOICES = [15, 30, 60, 120, 300, 900, 3600]
DEFAULT_REFRESH = 60

COMPACT_DURATIONS = [None, 3, 5, 8, 15, 30, 60]     # None = dauerhaft

ALPHA_MIN, ALPHA_MAX = 20, 100

# Zeiträume für Balken und Zwischensumme: Schlüssel -> (Textschlüssel, Tage)
# Tage = None bedeutet "alles".
ZEITRAEUME = [
    ("heute", "zeitHeute", 0),
    ("7t", "zeit7", 7),
    ("30t", "zeit30", 30),
    ("gesamt", "zeitGesamt", None),
]
DEFAULT_ZEITRAUM = "heute"

# Zusatzblock unter den Balken
ZUSATZ_CHOICES = [
    ("aus", "zusAus", None),
    ("laender", "zusLaender", "locations"),
    ("verweise", "zusVerweise", "toprefs"),
]
DEFAULT_ZUSATZ = "verweise"

# Weit genug vor dem ersten Seitenaufruf, damit "Gesamt" wirklich alles
# erfasst. GoatCounter selbst kennt kein "seit Anbeginn"-Kürzel.
ALLTIME_START = "2020-01-01T00:00:00Z"
ALLTIME_CACHE_SECONDS = 600           # Gesamtsumme ändert sich träge
FRISCH_SEKUNDEN = 10                  # "frisch" = aus demselben Durchlauf

MAX_BALKEN_CHOICES = [3, 5, 8, 10, 12, 15]
DEFAULT_MAX_BALKEN = 8

# Was gezählt wird: Schlüssel -> Textschlüssel. Siehe oben "Zu den Zahlen".
ZAEHLARTEN = [
    ("auto", "artAuto"),
    ("seiten", "artSeiten"),
    ("klicks", "artKlicks"),
]
DEFAULT_ZAEHLART = "auto"

# So beginnen die Ereignisse, die die Werkstatt für Klicks auf ihre
# Kacheln schickt: "kachel-<werkzeug>" und "kachel-<werkzeug>-<teil>".
KACHEL_PRAEFIX = "kachel-"

# /stats/hits liefert höchstens 100 Pfade je Anfrage, die meistbesuchten
# zuerst. Für die Summe aller Kachel-Klicks wird weitergeblättert, aber nie
# mehr als so viele Seiten (2000 Pfade; die Werkstatt hat gut 300).
HITS_SEITEN_MAX = 20

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

# Die Theme-Namen sind zugleich Schlüssel in der config.json und bleiben
# deshalb deutsch. Angezeigt wird der Name in der gewählten Sprache.
THEME_TEXT = {
    "Schwarz / Pink / Blau": "themeSPB",
    "Schwarz / Pink / Eisblau": "themeSPE",
    "Tiefschwarz / Magenta / Indigo": "themeTMI",
}

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

# ------------------------------ Sprachen ------------------------------------
# Alle sichtbaren Texte in 14 Sprachen. Gewählt wird im Tray unter
# "Sprache"; ohne Wahl gilt die Windows-Anzeigesprache, sonst Englisch.

TEXTE: dict[str, dict[str, str]] = {
    "de": {
        "_name": "Deutsch",
        "titel": "Besucher",
        "lade": "lade…",
        "heuteN": "{n} heute",
        "gesamtN": "{n} gesamt",
        "jeWerkzeug": "Besucher je Werkzeug",
        "artAuto": "automatisch",
        "artSeiten": "Seitenbesuche",
        "artKlicks": "Kachel-Klicks",
        "jeKachel": "Klicks je Kachel",
        "mZaehlart": "Zählweise",
        "setZaehlart": "Was gezählt wird",
        "startseite": "Startseite",
        "zeitHeute": "Heute",
        "zeit7": "7 Tage",
        "zeit30": "30 Tage",
        "zeitGesamt": "Gesamt",
        "zusAus": "aus",
        "zusLaender": "Länder",
        "zusVerweise": "Verweise",
        "standard": "Standard",
        "dauerhaft": "dauerhaft",
        "sekunden": "{n} Sekunden",
        "themeSPB": "Schwarz / Pink / Blau",
        "themeSPE": "Schwarz / Pink / Eisblau",
        "themeTMI": "Tiefschwarz / Magenta / Indigo",
        "frageSite": "Adresse der GoatCounter-Seite:\n(z. B. deinname  oder  deinname.goatcounter.com)",
        "frageToken": "API-Token einfügen\n(GoatCounter → Benutzername oben rechts → API →\n\"Add new token\", Recht \"Read statistics\" genügt):",
        "mOverlay": "Overlay an/aus",
        "mMitte": "Overlay zur Mitte holen",
        "mJetzt": "Jetzt aktualisieren",
        "mZeitraum": "Zeitraum",
        "mZusatz": "Zusatzblock",
        "mEinstellungen": "Einstellungen…",
        "mTheme": "Nächstes Theme",
        "mZugang": "Zugang ändern",
        "mSprache": "Sprache",
        "mBeenden": "Beenden",
        "setTitel": "Besucher-Ticker · Einstellungen",
        "setFreq": "Update-Frequenz",
        "setFreqHinweis": "GoatCounter erlaubt 4 Anfragen je Sekunde,\nselbst 15 s sind unkritisch.",
        "setZeitraum": "Zeitraum für Balken und Zwischensumme",
        "setZusatz": "Zusatzblock unter den Balken",
        "setBalken": "Höchstens so viele Balken",
        "setAlpha": "Transparenz ({min}–{max} %)",
        "setMiniMaxi": "Mini-/Maxi-Modus",
        "mini": "Mini",
        "maxi": "Maxi",
        "setZurueck": "Nach Umschalten zurück nach:",
        "setFarben": "Farben ({n} Varianten)",
        "schliessen": "Schließen",
        "ok": "OK",
        "abbrechen": "Abbrechen",
        "aktualisiert": "aktualisiert {zeit}",
        "netzFehler": "Netzwerkfehler, neuer Versuch…",
        "ratePause": "Rate-Limit (429), Pause",
        "tokenUngueltig": "Token ungültig!",
        "laeuftBereits": "Der Ticker läuft bereits.\n\nDas Fenster steckt im Infobereich der Taskleiste, Rechtsklick aufs Symbol.",
        "unbekannt": "unbekannt",
        "direkt": "direkt",
    },
    "en": {
        "_name": "English",
        "titel": "Visitors",
        "lade": "loading…",
        "heuteN": "{n} today",
        "gesamtN": "{n} total",
        "jeWerkzeug": "Visitors per tool",
        "artAuto": "automatic",
        "artSeiten": "Page visits",
        "artKlicks": "Tile clicks",
        "jeKachel": "Clicks per tile",
        "mZaehlart": "Counting",
        "setZaehlart": "What is counted",
        "startseite": "Home page",
        "zeitHeute": "Today",
        "zeit7": "7 days",
        "zeit30": "30 days",
        "zeitGesamt": "All time",
        "zusAus": "off",
        "zusLaender": "Countries",
        "zusVerweise": "Referrers",
        "standard": "default",
        "dauerhaft": "permanently",
        "sekunden": "{n} seconds",
        "themeSPB": "Black / Pink / Blue",
        "themeSPE": "Black / Pink / Ice blue",
        "themeTMI": "Deep black / Magenta / Indigo",
        "frageSite": "Address of the GoatCounter site:\n(e.g. yourname  or  yourname.goatcounter.com)",
        "frageToken": "Paste the API token\n(GoatCounter → user name at the top right → API →\n“Add new token”, the “Read statistics” permission is enough):",
        "mOverlay": "Show/hide overlay",
        "mMitte": "Bring overlay to centre",
        "mJetzt": "Refresh now",
        "mZeitraum": "Time range",
        "mZusatz": "Extra block",
        "mEinstellungen": "Settings…",
        "mTheme": "Next theme",
        "mZugang": "Change access",
        "mSprache": "Language",
        "mBeenden": "Quit",
        "setTitel": "Besucher-Ticker · Settings",
        "setFreq": "Update frequency",
        "setFreqHinweis": "GoatCounter allows 4 requests per second,\nso even 15 s is no problem.",
        "setZeitraum": "Time range for bars and subtotal",
        "setZusatz": "Extra block below the bars",
        "setBalken": "At most this many bars",
        "setAlpha": "Opacity ({min} to {max} %)",
        "setMiniMaxi": "Mini/maxi mode",
        "mini": "Mini",
        "maxi": "Maxi",
        "setZurueck": "Switch back after:",
        "setFarben": "Colours ({n} variants)",
        "schliessen": "Close",
        "ok": "OK",
        "abbrechen": "Cancel",
        "aktualisiert": "updated {zeit}",
        "netzFehler": "Network error, retrying…",
        "ratePause": "Rate limit (429), pausing",
        "tokenUngueltig": "Token invalid!",
        "laeuftBereits": "The ticker is already running.\n\nIts window sits in the notification area of the taskbar, right-click the icon.",
        "unbekannt": "unknown",
        "direkt": "direct",
    },
    "es": {
        "_name": "Español",
        "titel": "Visitantes",
        "lade": "cargando…",
        "heuteN": "{n} hoy",
        "gesamtN": "{n} en total",
        "jeWerkzeug": "Visitantes por herramienta",
        "artAuto": "automático",
        "artSeiten": "Visitas a páginas",
        "artKlicks": "Clics en tarjetas",
        "jeKachel": "Clics por tarjeta",
        "mZaehlart": "Recuento",
        "setZaehlart": "Qué se cuenta",
        "startseite": "Página de inicio",
        "zeitHeute": "Hoy",
        "zeit7": "7 días",
        "zeit30": "30 días",
        "zeitGesamt": "Todo",
        "zusAus": "no",
        "zusLaender": "Países",
        "zusVerweise": "Referencias",
        "standard": "predeterminado",
        "dauerhaft": "siempre",
        "sekunden": "{n} segundos",
        "themeSPB": "Negro / Rosa / Azul",
        "themeSPE": "Negro / Rosa / Azul hielo",
        "themeTMI": "Negro intenso / Magenta / Índigo",
        "frageSite": "Dirección del sitio de GoatCounter:\n(p. ej. tunombre  o  tunombre.goatcounter.com)",
        "frageToken": "Pega el token de la API\n(GoatCounter → nombre de usuario arriba a la derecha → API →\n«Add new token»; basta con el permiso «Read statistics»):",
        "mOverlay": "Mostrar/ocultar overlay",
        "mMitte": "Traer el overlay al centro",
        "mJetzt": "Actualizar ahora",
        "mZeitraum": "Periodo",
        "mZusatz": "Bloque extra",
        "mEinstellungen": "Ajustes…",
        "mTheme": "Siguiente tema",
        "mZugang": "Cambiar acceso",
        "mSprache": "Idioma",
        "mBeenden": "Salir",
        "setTitel": "Besucher-Ticker · Ajustes",
        "setFreq": "Frecuencia de actualización",
        "setFreqHinweis": "GoatCounter permite 4 peticiones por segundo,\nasí que incluso 15 s no es problema.",
        "setZeitraum": "Periodo para barras y subtotal",
        "setZusatz": "Bloque extra bajo las barras",
        "setBalken": "Como máximo tantas barras",
        "setAlpha": "Opacidad (de {min} a {max} %)",
        "setMiniMaxi": "Modo mini/maxi",
        "mini": "Mini",
        "maxi": "Maxi",
        "setZurueck": "Volver tras:",
        "setFarben": "Colores ({n} variantes)",
        "schliessen": "Cerrar",
        "ok": "Aceptar",
        "abbrechen": "Cancelar",
        "aktualisiert": "actualizado {zeit}",
        "netzFehler": "Error de red, reintentando…",
        "ratePause": "Límite de peticiones (429), pausa",
        "tokenUngueltig": "¡Token no válido!",
        "laeuftBereits": "El ticker ya está en marcha.\n\nSu ventana está en el área de notificación de la barra de tareas: haz clic derecho en el icono.",
        "unbekannt": "desconocido",
        "direkt": "directo",
    },
    "fr": {
        "_name": "Français",
        "titel": "Visiteurs",
        "lade": "chargement…",
        "heuteN": "{n} aujourd’hui",
        "gesamtN": "{n} au total",
        "jeWerkzeug": "Visiteurs par outil",
        "artAuto": "automatique",
        "artSeiten": "Visites de pages",
        "artKlicks": "Clics sur les vignettes",
        "jeKachel": "Clics par vignette",
        "mZaehlart": "Comptage",
        "setZaehlart": "Ce qui est compté",
        "startseite": "Page d’accueil",
        "zeitHeute": "Aujourd’hui",
        "zeit7": "7 jours",
        "zeit30": "30 jours",
        "zeitGesamt": "Depuis le début",
        "zusAus": "non",
        "zusLaender": "Pays",
        "zusVerweise": "Référents",
        "standard": "par défaut",
        "dauerhaft": "en permanence",
        "sekunden": "{n} secondes",
        "themeSPB": "Noir / Rose / Bleu",
        "themeSPE": "Noir / Rose / Bleu glacier",
        "themeTMI": "Noir profond / Magenta / Indigo",
        "frageSite": "Adresse du site GoatCounter :\n(p. ex. votrenom  ou  votrenom.goatcounter.com)",
        "frageToken": "Collez le jeton d’API\n(GoatCounter → nom d’utilisateur en haut à droite → API →\n« Add new token », le droit « Read statistics » suffit) :",
        "mOverlay": "Afficher/masquer l’overlay",
        "mMitte": "Ramener l’overlay au centre",
        "mJetzt": "Actualiser maintenant",
        "mZeitraum": "Période",
        "mZusatz": "Bloc supplémentaire",
        "mEinstellungen": "Réglages…",
        "mTheme": "Thème suivant",
        "mZugang": "Changer l’accès",
        "mSprache": "Langue",
        "mBeenden": "Quitter",
        "setTitel": "Besucher-Ticker · Réglages",
        "setFreq": "Fréquence de mise à jour",
        "setFreqHinweis": "GoatCounter autorise 4 requêtes par seconde,\nmême 15 s ne pose aucun problème.",
        "setZeitraum": "Période pour les barres et le sous-total",
        "setZusatz": "Bloc supplémentaire sous les barres",
        "setBalken": "Au plus autant de barres",
        "setAlpha": "Opacité (de {min} à {max} %)",
        "setMiniMaxi": "Mode mini/maxi",
        "mini": "Mini",
        "maxi": "Maxi",
        "setZurueck": "Revenir après :",
        "setFarben": "Couleurs ({n} variantes)",
        "schliessen": "Fermer",
        "ok": "OK",
        "abbrechen": "Annuler",
        "aktualisiert": "mis à jour {zeit}",
        "netzFehler": "Erreur réseau, nouvel essai…",
        "ratePause": "Limite de requêtes (429), pause",
        "tokenUngueltig": "Jeton non valide !",
        "laeuftBereits": "Le ticker est déjà lancé.\n\nSa fenêtre se trouve dans la zone de notification de la barre des tâches : faites un clic droit sur l’icône.",
        "unbekannt": "inconnu",
        "direkt": "direct",
    },
    "it": {
        "_name": "Italiano",
        "titel": "Visitatori",
        "lade": "caricamento…",
        "heuteN": "{n} oggi",
        "gesamtN": "{n} in totale",
        "jeWerkzeug": "Visitatori per strumento",
        "artAuto": "automatico",
        "artSeiten": "Visite alle pagine",
        "artKlicks": "Clic sulle schede",
        "jeKachel": "Clic per scheda",
        "mZaehlart": "Conteggio",
        "setZaehlart": "Cosa viene contato",
        "startseite": "Pagina iniziale",
        "zeitHeute": "Oggi",
        "zeit7": "7 giorni",
        "zeit30": "30 giorni",
        "zeitGesamt": "Da sempre",
        "zusAus": "no",
        "zusLaender": "Paesi",
        "zusVerweise": "Referrer",
        "standard": "predefinito",
        "dauerhaft": "sempre",
        "sekunden": "{n} secondi",
        "themeSPB": "Nero / Rosa / Blu",
        "themeSPE": "Nero / Rosa / Azzurro ghiaccio",
        "themeTMI": "Nero profondo / Magenta / Indaco",
        "frageSite": "Indirizzo del sito GoatCounter:\n(per es. tuonome  o  tuonome.goatcounter.com)",
        "frageToken": "Incolla il token API\n(GoatCounter → nome utente in alto a destra → API →\n«Add new token», basta il permesso «Read statistics»):",
        "mOverlay": "Mostra/nascondi overlay",
        "mMitte": "Riporta l’overlay al centro",
        "mJetzt": "Aggiorna ora",
        "mZeitraum": "Periodo",
        "mZusatz": "Blocco extra",
        "mEinstellungen": "Impostazioni…",
        "mTheme": "Tema successivo",
        "mZugang": "Cambia accesso",
        "mSprache": "Lingua",
        "mBeenden": "Esci",
        "setTitel": "Besucher-Ticker · Impostazioni",
        "setFreq": "Frequenza di aggiornamento",
        "setFreqHinweis": "GoatCounter consente 4 richieste al secondo,\nanche 15 s non sono un problema.",
        "setZeitraum": "Periodo per barre e subtotale",
        "setZusatz": "Blocco extra sotto le barre",
        "setBalken": "Al massimo tante barre",
        "setAlpha": "Opacità (da {min} a {max} %)",
        "setMiniMaxi": "Modalità mini/maxi",
        "mini": "Mini",
        "maxi": "Maxi",
        "setZurueck": "Torna indietro dopo:",
        "setFarben": "Colori ({n} varianti)",
        "schliessen": "Chiudi",
        "ok": "OK",
        "abbrechen": "Annulla",
        "aktualisiert": "aggiornato {zeit}",
        "netzFehler": "Errore di rete, nuovo tentativo…",
        "ratePause": "Limite di richieste (429), pausa",
        "tokenUngueltig": "Token non valido!",
        "laeuftBereits": "Il ticker è già in esecuzione.\n\nLa sua finestra è nell’area di notifica della barra delle applicazioni: fai clic destro sull’icona.",
        "unbekannt": "sconosciuto",
        "direkt": "diretto",
    },
    "nl": {
        "_name": "Nederlands",
        "titel": "Bezoekers",
        "lade": "laden…",
        "heuteN": "{n} vandaag",
        "gesamtN": "{n} in totaal",
        "jeWerkzeug": "Bezoekers per tool",
        "artAuto": "automatisch",
        "artSeiten": "Paginabezoeken",
        "artKlicks": "Tegelklikken",
        "jeKachel": "Klikken per tegel",
        "mZaehlart": "Telwijze",
        "setZaehlart": "Wat er geteld wordt",
        "startseite": "Startpagina",
        "zeitHeute": "Vandaag",
        "zeit7": "7 dagen",
        "zeit30": "30 dagen",
        "zeitGesamt": "Altijd",
        "zusAus": "uit",
        "zusLaender": "Landen",
        "zusVerweise": "Verwijzers",
        "standard": "standaard",
        "dauerhaft": "altijd",
        "sekunden": "{n} seconden",
        "themeSPB": "Zwart / Roze / Blauw",
        "themeSPE": "Zwart / Roze / IJsblauw",
        "themeTMI": "Diepzwart / Magenta / Indigo",
        "frageSite": "Adres van de GoatCounter-site:\n(bijv. jouwnaam  of  jouwnaam.goatcounter.com)",
        "frageToken": "API-token plakken\n(GoatCounter → gebruikersnaam rechtsboven → API →\n‘Add new token’, het recht ‘Read statistics’ is genoeg):",
        "mOverlay": "Overlay tonen/verbergen",
        "mMitte": "Overlay naar het midden halen",
        "mJetzt": "Nu bijwerken",
        "mZeitraum": "Periode",
        "mZusatz": "Extra blok",
        "mEinstellungen": "Instellingen…",
        "mTheme": "Volgend thema",
        "mZugang": "Toegang wijzigen",
        "mSprache": "Taal",
        "mBeenden": "Afsluiten",
        "setTitel": "Besucher-Ticker · Instellingen",
        "setFreq": "Updatefrequentie",
        "setFreqHinweis": "GoatCounter staat 4 verzoeken per seconde toe,\ndus zelfs 15 s is geen probleem.",
        "setZeitraum": "Periode voor balken en subtotaal",
        "setZusatz": "Extra blok onder de balken",
        "setBalken": "Hoogstens zoveel balken",
        "setAlpha": "Dekking ({min} tot {max} %)",
        "setMiniMaxi": "Mini-/maximodus",
        "mini": "Mini",
        "maxi": "Maxi",
        "setZurueck": "Terugschakelen na:",
        "setFarben": "Kleuren ({n} varianten)",
        "schliessen": "Sluiten",
        "ok": "OK",
        "abbrechen": "Annuleren",
        "aktualisiert": "bijgewerkt {zeit}",
        "netzFehler": "Netwerkfout, nieuwe poging…",
        "ratePause": "Limiet bereikt (429), pauze",
        "tokenUngueltig": "Token ongeldig!",
        "laeuftBereits": "De ticker draait al.\n\nHet venster zit in het systeemvak van de taakbalk: klik met rechts op het pictogram.",
        "unbekannt": "onbekend",
        "direkt": "direct",
    },
    "pl": {
        "_name": "Polski",
        "titel": "Odwiedziny",
        "lade": "wczytywanie…",
        "heuteN": "dziś: {n}",
        "gesamtN": "razem: {n}",
        "jeWerkzeug": "Odwiedziny według narzędzia",
        "artAuto": "automatycznie",
        "artSeiten": "Odwiedziny stron",
        "artKlicks": "Kliknięcia kafelków",
        "jeKachel": "Kliknięcia według kafelka",
        "mZaehlart": "Liczenie",
        "setZaehlart": "Co jest liczone",
        "startseite": "Strona główna",
        "zeitHeute": "Dziś",
        "zeit7": "7 dni",
        "zeit30": "30 dni",
        "zeitGesamt": "Od początku",
        "zusAus": "wył.",
        "zusLaender": "Kraje",
        "zusVerweise": "Odsyłacze",
        "standard": "domyślnie",
        "dauerhaft": "na stałe",
        "sekunden": "{n} s",
        "themeSPB": "Czarny / Różowy / Niebieski",
        "themeSPE": "Czarny / Różowy / Lodowy błękit",
        "themeTMI": "Głęboka czerń / Magenta / Indygo",
        "frageSite": "Adres strony GoatCounter:\n(np. twojanazwa  lub  twojanazwa.goatcounter.com)",
        "frageToken": "Wklej token API\n(GoatCounter → nazwa użytkownika w prawym górnym rogu → API →\n„Add new token”, wystarczy uprawnienie „Read statistics”):",
        "mOverlay": "Pokaż/ukryj nakładkę",
        "mMitte": "Przenieś nakładkę na środek",
        "mJetzt": "Odśwież teraz",
        "mZeitraum": "Okres",
        "mZusatz": "Dodatkowy blok",
        "mEinstellungen": "Ustawienia…",
        "mTheme": "Następny motyw",
        "mZugang": "Zmień dostęp",
        "mSprache": "Język",
        "mBeenden": "Zakończ",
        "setTitel": "Besucher-Ticker · Ustawienia",
        "setFreq": "Częstotliwość odświeżania",
        "setFreqHinweis": "GoatCounter pozwala na 4 zapytania na sekundę,\nwięc nawet 15 s to żaden problem.",
        "setZeitraum": "Okres dla pasków i sumy częściowej",
        "setZusatz": "Dodatkowy blok pod paskami",
        "setBalken": "Najwyżej tyle pasków",
        "setAlpha": "Krycie (od {min} do {max} %)",
        "setMiniMaxi": "Tryb mini/maxi",
        "mini": "Mini",
        "maxi": "Maxi",
        "setZurueck": "Wróć po:",
        "setFarben": "Kolory (warianty: {n})",
        "schliessen": "Zamknij",
        "ok": "OK",
        "abbrechen": "Anuluj",
        "aktualisiert": "aktualizacja {zeit}",
        "netzFehler": "Błąd sieci, ponawiam…",
        "ratePause": "Limit zapytań (429), przerwa",
        "tokenUngueltig": "Nieprawidłowy token!",
        "laeuftBereits": "Ticker już działa.\n\nJego okno jest w obszarze powiadomień paska zadań: kliknij ikonę prawym przyciskiem.",
        "unbekannt": "nieznane",
        "direkt": "bezpośrednio",
    },
    "pt": {
        "_name": "Português",
        "titel": "Visitantes",
        "lade": "carregando…",
        "heuteN": "{n} hoje",
        "gesamtN": "{n} no total",
        "jeWerkzeug": "Visitantes por ferramenta",
        "artAuto": "automático",
        "artSeiten": "Visitas às páginas",
        "artKlicks": "Cliques nos cartões",
        "jeKachel": "Cliques por cartão",
        "mZaehlart": "Contagem",
        "setZaehlart": "O que é contado",
        "startseite": "Página inicial",
        "zeitHeute": "Hoje",
        "zeit7": "7 dias",
        "zeit30": "30 dias",
        "zeitGesamt": "Desde sempre",
        "zusAus": "não",
        "zusLaender": "Países",
        "zusVerweise": "Origens",
        "standard": "padrão",
        "dauerhaft": "sempre",
        "sekunden": "{n} segundos",
        "themeSPB": "Preto / Rosa / Azul",
        "themeSPE": "Preto / Rosa / Azul-gelo",
        "themeTMI": "Preto profundo / Magenta / Índigo",
        "frageSite": "Endereço do site no GoatCounter:\n(ex.: seunome  ou  seunome.goatcounter.com)",
        "frageToken": "Cole o token da API\n(GoatCounter → nome de usuário no canto superior direito → API →\n«Add new token», a permissão «Read statistics» basta):",
        "mOverlay": "Mostrar/ocultar overlay",
        "mMitte": "Trazer o overlay para o centro",
        "mJetzt": "Atualizar agora",
        "mZeitraum": "Período",
        "mZusatz": "Bloco extra",
        "mEinstellungen": "Configurações…",
        "mTheme": "Próximo tema",
        "mZugang": "Mudar acesso",
        "mSprache": "Idioma",
        "mBeenden": "Sair",
        "setTitel": "Besucher-Ticker · Configurações",
        "setFreq": "Frequência de atualização",
        "setFreqHinweis": "O GoatCounter permite 4 requisições por segundo,\nentão até 15 s é tranquilo.",
        "setZeitraum": "Período para as barras e o subtotal",
        "setZusatz": "Bloco extra abaixo das barras",
        "setBalken": "No máximo esta quantidade de barras",
        "setAlpha": "Opacidade (de {min} a {max} %)",
        "setMiniMaxi": "Modo mini/maxi",
        "mini": "Mini",
        "maxi": "Maxi",
        "setZurueck": "Voltar depois de:",
        "setFarben": "Cores ({n} variantes)",
        "schliessen": "Fechar",
        "ok": "OK",
        "abbrechen": "Cancelar",
        "aktualisiert": "atualizado {zeit}",
        "netzFehler": "Erro de rede, tentando de novo…",
        "ratePause": "Limite de requisições (429), pausa",
        "tokenUngueltig": "Token inválido!",
        "laeuftBereits": "O ticker já está em execução.\n\nA janela fica na área de notificação da barra de tarefas: clique no ícone com o botão direito.",
        "unbekannt": "desconhecido",
        "direkt": "direto",
    },
    "tr": {
        "_name": "Türkçe",
        "titel": "Ziyaretçiler",
        "lade": "yükleniyor…",
        "heuteN": "bugün {n}",
        "gesamtN": "toplam {n}",
        "jeWerkzeug": "Araç başına ziyaretçi",
        "artAuto": "otomatik",
        "artSeiten": "Sayfa ziyaretleri",
        "artKlicks": "Kart tıklamaları",
        "jeKachel": "Kart başına tıklama",
        "mZaehlart": "Sayım",
        "setZaehlart": "Neyin sayıldığı",
        "startseite": "Ana sayfa",
        "zeitHeute": "Bugün",
        "zeit7": "7 gün",
        "zeit30": "30 gün",
        "zeitGesamt": "Tüm zamanlar",
        "zusAus": "kapalı",
        "zusLaender": "Ülkeler",
        "zusVerweise": "Yönlendirenler",
        "standard": "varsayılan",
        "dauerhaft": "kalıcı",
        "sekunden": "{n} saniye",
        "themeSPB": "Siyah / Pembe / Mavi",
        "themeSPE": "Siyah / Pembe / Buz mavisi",
        "themeTMI": "Koyu siyah / Macenta / Çivit",
        "frageSite": "GoatCounter sitesinin adresi:\n(örn. adiniz  veya  adiniz.goatcounter.com)",
        "frageToken": "API anahtarını yapıştırın\n(GoatCounter → sağ üstte kullanıcı adı → API →\n“Add new token”, “Read statistics” izni yeterli):",
        "mOverlay": "Katmanı göster/gizle",
        "mMitte": "Katmanı ortaya getir",
        "mJetzt": "Şimdi güncelle",
        "mZeitraum": "Zaman aralığı",
        "mZusatz": "Ek blok",
        "mEinstellungen": "Ayarlar…",
        "mTheme": "Sonraki tema",
        "mZugang": "Erişimi değiştir",
        "mSprache": "Dil",
        "mBeenden": "Çık",
        "setTitel": "Besucher-Ticker · Ayarlar",
        "setFreq": "Güncelleme sıklığı",
        "setFreqHinweis": "GoatCounter saniyede 4 isteğe izin verir,\nyani 15 s bile sorun değil.",
        "setZeitraum": "Çubuklar ve ara toplam için zaman aralığı",
        "setZusatz": "Çubukların altındaki ek blok",
        "setBalken": "En fazla bu kadar çubuk",
        "setAlpha": "Opaklık (%{min} ile %{max} arası)",
        "setMiniMaxi": "Mini/maksi modu",
        "mini": "Mini",
        "maxi": "Maksi",
        "setZurueck": "Şu süre sonra geri dön:",
        "setFarben": "Renkler ({n} çeşit)",
        "schliessen": "Kapat",
        "ok": "Tamam",
        "abbrechen": "İptal",
        "aktualisiert": "güncellendi {zeit}",
        "netzFehler": "Ağ hatası, yeniden deneniyor…",
        "ratePause": "İstek sınırı (429), mola",
        "tokenUngueltig": "Anahtar geçersiz!",
        "laeuftBereits": "Ticker zaten çalışıyor.\n\nPenceresi görev çubuğunun bildirim alanında: simgeye sağ tıklayın.",
        "unbekannt": "bilinmiyor",
        "direkt": "doğrudan",
    },
    "ru": {
        "_name": "Русский",
        "titel": "Посетители",
        "lade": "загрузка…",
        "heuteN": "сегодня: {n}",
        "gesamtN": "всего: {n}",
        "jeWerkzeug": "Посетители по инструментам",
        "artAuto": "автоматически",
        "artSeiten": "Посещения страниц",
        "artKlicks": "Клики по плиткам",
        "jeKachel": "Клики по каждой плитке",
        "mZaehlart": "Подсчёт",
        "setZaehlart": "Что считается",
        "startseite": "Главная страница",
        "zeitHeute": "Сегодня",
        "zeit7": "7 дней",
        "zeit30": "30 дней",
        "zeitGesamt": "За всё время",
        "zusAus": "выкл.",
        "zusLaender": "Страны",
        "zusVerweise": "Источники",
        "standard": "по умолчанию",
        "dauerhaft": "постоянно",
        "sekunden": "{n} с",
        "themeSPB": "Чёрный / Розовый / Синий",
        "themeSPE": "Чёрный / Розовый / Ледяной синий",
        "themeTMI": "Глубокий чёрный / Пурпурный / Индиго",
        "frageSite": "Адрес сайта в GoatCounter:\n(напр. yourname  или  yourname.goatcounter.com)",
        "frageToken": "Вставьте токен API\n(GoatCounter → имя пользователя справа вверху → API →\n«Add new token», достаточно права «Read statistics»):",
        "mOverlay": "Показать/скрыть оверлей",
        "mMitte": "Вернуть оверлей в центр",
        "mJetzt": "Обновить сейчас",
        "mZeitraum": "Период",
        "mZusatz": "Дополнительный блок",
        "mEinstellungen": "Настройки…",
        "mTheme": "Следующая тема",
        "mZugang": "Изменить доступ",
        "mSprache": "Язык",
        "mBeenden": "Выход",
        "setTitel": "Besucher-Ticker · Настройки",
        "setFreq": "Частота обновления",
        "setFreqHinweis": "GoatCounter разрешает 4 запроса в секунду,\nтак что даже 15 с не проблема.",
        "setZeitraum": "Период для полос и промежуточной суммы",
        "setZusatz": "Дополнительный блок под полосами",
        "setBalken": "Не больше стольких полос",
        "setAlpha": "Непрозрачность (от {min} до {max} %)",
        "setMiniMaxi": "Режим мини/макси",
        "mini": "Мини",
        "maxi": "Макси",
        "setZurueck": "Вернуть через:",
        "setFarben": "Цвета (вариантов: {n})",
        "schliessen": "Закрыть",
        "ok": "ОК",
        "abbrechen": "Отмена",
        "aktualisiert": "обновлено {zeit}",
        "netzFehler": "Ошибка сети, повтор…",
        "ratePause": "Лимит запросов (429), пауза",
        "tokenUngueltig": "Токен недействителен!",
        "laeuftBereits": "Тикер уже запущен.\n\nЕго окно находится в области уведомлений панели задач: щёлкните по значку правой кнопкой.",
        "unbekannt": "неизвестно",
        "direkt": "напрямую",
    },
    "hi": {
        "_name": "हिन्दी",
        "titel": "विज़िटर",
        "lade": "लोड हो रहा है…",
        "heuteN": "आज {n}",
        "gesamtN": "कुल {n}",
        "jeWerkzeug": "हर टूल के विज़िटर",
        "artAuto": "स्वचालित",
        "artSeiten": "पेज विज़िट",
        "artKlicks": "कार्ड क्लिक",
        "jeKachel": "हर कार्ड के क्लिक",
        "mZaehlart": "गिनती",
        "setZaehlart": "क्या गिना जाता है",
        "startseite": "होम पेज",
        "zeitHeute": "आज",
        "zeit7": "7 दिन",
        "zeit30": "30 दिन",
        "zeitGesamt": "शुरू से अब तक",
        "zusAus": "बंद",
        "zusLaender": "देश",
        "zusVerweise": "रेफ़रर",
        "standard": "डिफ़ॉल्ट",
        "dauerhaft": "हमेशा",
        "sekunden": "{n} सेकंड",
        "themeSPB": "काला / गुलाबी / नीला",
        "themeSPE": "काला / गुलाबी / बर्फ़ीला नीला",
        "themeTMI": "गहरा काला / मैजेंटा / इंडिगो",
        "frageSite": "GoatCounter साइट का पता:\n(जैसे yourname  या  yourname.goatcounter.com)",
        "frageToken": "API टोकन चिपकाएँ\n(GoatCounter → ऊपर दाईं ओर उपयोगकर्ता नाम → API →\n“Add new token”, “Read statistics” अनुमति काफ़ी है):",
        "mOverlay": "ओवरले दिखाएँ/छिपाएँ",
        "mMitte": "ओवरले को बीच में लाएँ",
        "mJetzt": "अभी अपडेट करें",
        "mZeitraum": "समय अवधि",
        "mZusatz": "अतिरिक्त ब्लॉक",
        "mEinstellungen": "सेटिंग्स…",
        "mTheme": "अगली थीम",
        "mZugang": "पहुँच बदलें",
        "mSprache": "भाषा",
        "mBeenden": "बंद करें",
        "setTitel": "Besucher-Ticker · सेटिंग्स",
        "setFreq": "अपडेट की आवृत्ति",
        "setFreqHinweis": "GoatCounter हर सेकंड 4 अनुरोधों की अनुमति देता है,\nइसलिए 15 s भी कोई समस्या नहीं।",
        "setZeitraum": "बार और उप-योग के लिए समय अवधि",
        "setZusatz": "बार के नीचे अतिरिक्त ब्लॉक",
        "setBalken": "अधिकतम इतने बार",
        "setAlpha": "अपारदर्शिता ({min} से {max} %)",
        "setMiniMaxi": "मिनी/मैक्सी मोड",
        "mini": "मिनी",
        "maxi": "मैक्सी",
        "setZurueck": "इतने बाद वापस जाएँ:",
        "setFarben": "रंग ({n} प्रकार)",
        "schliessen": "बंद करें",
        "ok": "ठीक है",
        "abbrechen": "रद्द करें",
        "aktualisiert": "अपडेट {zeit}",
        "netzFehler": "नेटवर्क त्रुटि, फिर से कोशिश…",
        "ratePause": "अनुरोध सीमा (429), विराम",
        "tokenUngueltig": "टोकन अमान्य!",
        "laeuftBereits": "टिकर पहले से चल रहा है।\n\nइसकी विंडो टास्कबार के सूचना क्षेत्र में है: आइकन पर राइट-क्लिक करें।",
        "unbekannt": "अज्ञात",
        "direkt": "सीधे",
    },
    "zh": {
        "_name": "中文",
        "titel": "访客",
        "lade": "正在加载…",
        "heuteN": "今天 {n}",
        "gesamtN": "总计 {n}",
        "jeWerkzeug": "各工具的访客",
        "artAuto": "自动",
        "artSeiten": "页面访问",
        "artKlicks": "卡片点击",
        "jeKachel": "各卡片的点击",
        "mZaehlart": "计数方式",
        "setZaehlart": "统计什么",
        "startseite": "首页",
        "zeitHeute": "今天",
        "zeit7": "7 天",
        "zeit30": "30 天",
        "zeitGesamt": "全部时间",
        "zusAus": "关",
        "zusLaender": "国家",
        "zusVerweise": "来源",
        "standard": "默认",
        "dauerhaft": "一直",
        "sekunden": "{n} 秒",
        "themeSPB": "黑 / 粉 / 蓝",
        "themeSPE": "黑 / 粉 / 冰蓝",
        "themeTMI": "深黑 / 品红 / 靛蓝",
        "frageSite": "GoatCounter 站点地址：\n（例如 yourname  或  yourname.goatcounter.com）",
        "frageToken": "粘贴 API 令牌\n（GoatCounter → 右上角的用户名 → API →\n“Add new token”，有“Read statistics”权限就够了）：",
        "mOverlay": "显示/隐藏悬浮窗",
        "mMitte": "把悬浮窗移到屏幕中央",
        "mJetzt": "立即更新",
        "mZeitraum": "时间范围",
        "mZusatz": "附加区块",
        "mEinstellungen": "设置…",
        "mTheme": "下一个主题",
        "mZugang": "更改访问凭据",
        "mSprache": "语言",
        "mBeenden": "退出",
        "setTitel": "Besucher-Ticker · 设置",
        "setFreq": "更新频率",
        "setFreqHinweis": "GoatCounter 允许每秒 4 个请求，\n所以就算 15 秒也没问题。",
        "setZeitraum": "条形图和小计的时间范围",
        "setZusatz": "条形图下方的附加区块",
        "setBalken": "最多显示多少条",
        "setAlpha": "不透明度（{min} 到 {max} %）",
        "setMiniMaxi": "迷你/完整模式",
        "mini": "迷你",
        "maxi": "完整",
        "setZurueck": "切换后多久恢复：",
        "setFarben": "颜色（{n} 种）",
        "schliessen": "关闭",
        "ok": "确定",
        "abbrechen": "取消",
        "aktualisiert": "更新于 {zeit}",
        "netzFehler": "网络错误，正在重试…",
        "ratePause": "请求超限 (429)，暂停",
        "tokenUngueltig": "令牌无效！",
        "laeuftBereits": "Ticker 已在运行。\n\n它的窗口在任务栏的通知区域：右键单击图标。",
        "unbekannt": "未知",
        "direkt": "直接访问",
    },
    "ja": {
        "_name": "日本語",
        "titel": "訪問者",
        "lade": "読み込み中…",
        "heuteN": "今日 {n}",
        "gesamtN": "合計 {n}",
        "jeWerkzeug": "ツールごとの訪問者",
        "artAuto": "自動",
        "artSeiten": "ページ訪問",
        "artKlicks": "カードのクリック",
        "jeKachel": "カードごとのクリック",
        "mZaehlart": "集計方法",
        "setZaehlart": "何を数えるか",
        "startseite": "トップページ",
        "zeitHeute": "今日",
        "zeit7": "7 日間",
        "zeit30": "30 日間",
        "zeitGesamt": "全期間",
        "zusAus": "オフ",
        "zusLaender": "国",
        "zusVerweise": "参照元",
        "standard": "標準",
        "dauerhaft": "ずっと",
        "sekunden": "{n} 秒",
        "themeSPB": "黒 / ピンク / 青",
        "themeSPE": "黒 / ピンク / アイスブルー",
        "themeTMI": "漆黒 / マゼンタ / インディゴ",
        "frageSite": "GoatCounter サイトのアドレス：\n（例：yourname  または  yourname.goatcounter.com）",
        "frageToken": "API トークンを貼り付けてください\n（GoatCounter → 右上のユーザー名 → API →\n「Add new token」。「Read statistics」の権限で十分です）：",
        "mOverlay": "オーバーレイの表示/非表示",
        "mMitte": "オーバーレイを中央に戻す",
        "mJetzt": "今すぐ更新",
        "mZeitraum": "期間",
        "mZusatz": "追加ブロック",
        "mEinstellungen": "設定…",
        "mTheme": "次のテーマ",
        "mZugang": "アクセス情報を変更",
        "mSprache": "言語",
        "mBeenden": "終了",
        "setTitel": "Besucher-Ticker · 設定",
        "setFreq": "更新頻度",
        "setFreqHinweis": "GoatCounter は毎秒 4 件までリクエストできるので、\n15 秒でも問題ありません。",
        "setZeitraum": "バーと小計の期間",
        "setZusatz": "バーの下の追加ブロック",
        "setBalken": "バーの最大数",
        "setAlpha": "不透明度（{min}〜{max} %）",
        "setMiniMaxi": "ミニ/マキシ表示",
        "mini": "ミニ",
        "maxi": "マキシ",
        "setZurueck": "切り替え後に戻すまで：",
        "setFarben": "色（{n} 種類）",
        "schliessen": "閉じる",
        "ok": "OK",
        "abbrechen": "キャンセル",
        "aktualisiert": "更新 {zeit}",
        "netzFehler": "ネットワークエラー、再試行中…",
        "ratePause": "リクエスト上限 (429)、一時停止",
        "tokenUngueltig": "トークンが無効です！",
        "laeuftBereits": "Ticker はすでに起動しています。\n\nウィンドウはタスクバーの通知領域にあります。アイコンを右クリックしてください。",
        "unbekannt": "不明",
        "direkt": "直接",
    },
    "ko": {
        "_name": "한국어",
        "titel": "방문자",
        "lade": "불러오는 중…",
        "heuteN": "오늘 {n}",
        "gesamtN": "전체 {n}",
        "jeWerkzeug": "도구별 방문자",
        "artAuto": "자동",
        "artSeiten": "페이지 방문",
        "artKlicks": "타일 클릭",
        "jeKachel": "타일별 클릭",
        "mZaehlart": "집계 방식",
        "setZaehlart": "무엇을 집계할지",
        "startseite": "홈페이지",
        "zeitHeute": "오늘",
        "zeit7": "7일",
        "zeit30": "30일",
        "zeitGesamt": "전체 기간",
        "zusAus": "끔",
        "zusLaender": "국가",
        "zusVerweise": "유입 경로",
        "standard": "기본",
        "dauerhaft": "계속",
        "sekunden": "{n}초",
        "themeSPB": "검정 / 분홍 / 파랑",
        "themeSPE": "검정 / 분홍 / 얼음 파랑",
        "themeTMI": "깊은 검정 / 자홍 / 남색",
        "frageSite": "GoatCounter 사이트 주소:\n(예: yourname  또는  yourname.goatcounter.com)",
        "frageToken": "API 토큰을 붙여 넣으세요\n(GoatCounter → 오른쪽 위 사용자 이름 → API →\n‘Add new token’, ‘Read statistics’ 권한이면 충분합니다):",
        "mOverlay": "오버레이 표시/숨기기",
        "mMitte": "오버레이를 가운데로",
        "mJetzt": "지금 업데이트",
        "mZeitraum": "기간",
        "mZusatz": "추가 블록",
        "mEinstellungen": "설정…",
        "mTheme": "다음 테마",
        "mZugang": "접속 정보 변경",
        "mSprache": "언어",
        "mBeenden": "종료",
        "setTitel": "Besucher-Ticker · 설정",
        "setFreq": "업데이트 빈도",
        "setFreqHinweis": "GoatCounter는 초당 4건의 요청을 허용하므로\n15초도 문제없습니다.",
        "setZeitraum": "막대와 소계의 기간",
        "setZusatz": "막대 아래 추가 블록",
        "setBalken": "최대 막대 수",
        "setAlpha": "불투명도 ({min}~{max} %)",
        "setMiniMaxi": "미니/맥시 모드",
        "mini": "미니",
        "maxi": "맥시",
        "setZurueck": "전환 후 되돌릴 시간:",
        "setFarben": "색상 ({n}가지)",
        "schliessen": "닫기",
        "ok": "확인",
        "abbrechen": "취소",
        "aktualisiert": "업데이트 {zeit}",
        "netzFehler": "네트워크 오류, 다시 시도 중…",
        "ratePause": "요청 한도 (429), 잠시 멈춤",
        "tokenUngueltig": "토큰이 올바르지 않습니다!",
        "laeuftBereits": "Ticker가 이미 실행 중입니다.\n\n창은 작업 표시줄의 알림 영역에 있습니다. 아이콘을 마우스 오른쪽 버튼으로 클릭하세요.",
        "unbekannt": "알 수 없음",
        "direkt": "직접",
    },
}

# Windows-Sprachkennung (untere 10 Bit) -> Sprachcode
_LANGID = {0x07: "de", 0x09: "en", 0x0A: "es", 0x0C: "fr", 0x10: "it", 0x13: "nl", 0x15: "pl",
           0x16: "pt", 0x1F: "tr", 0x19: "ru", 0x39: "hi", 0x04: "zh", 0x11: "ja", 0x12: "ko"}

SPRACHE = "de"                        # beim Start aus config.json oder Windows gesetzt


def systemsprache() -> str:
    """Windows-Anzeigesprache, wenn der Ticker sie spricht, sonst Englisch."""
    try:
        import ctypes
        return _LANGID.get(ctypes.windll.kernel32.GetUserDefaultUILanguage() & 0x3FF, "en")
    except (ImportError, AttributeError, OSError):
        return "en"


def waehle_sprache(cfg: dict) -> str:
    s = cfg.get("sprache")
    return s if s in TEXTE else systemsprache()


def tx(schluessel: str, **werte) -> str:
    """Text in der gewählten Sprache; {n}, {zeit} usw. werden eingesetzt."""
    s = TEXTE[SPRACHE].get(schluessel) or TEXTE["en"].get(schluessel, schluessel)
    for k, v in werte.items():
        s = s.replace("{" + k + "}", str(v))
    return s


def anzeige(label: str) -> str:
    """Interne Beschriftungen der Balken und Listen in der gewählten Sprache."""
    if label == "Startseite":
        return tx("startseite")
    if label in ("(unbekannt)", "(direkt)"):
        return f"({tx(label[1:-1])})"
    return label


def refresh_name(sek: float) -> str:
    name = f"{sek / 60:g} min" if sek >= 120 else f"{sek:g} s"
    return name + (f"  ({tx('standard')})" if sek == DEFAULT_REFRESH else "")


def dauer_name(sek: int | None) -> str:
    return tx("dauerhaft") if sek is None else tx("sekunden", n=sek)


def theme_anzeige(name: str) -> str:
    return tx(THEME_TEXT[name]) if name in THEME_TEXT else name

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
    """Nimmt 'deinname', 'deinname.goatcounter.com' oder die
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


def titelleiste_dunkel(fenster: tk.Misc) -> None:
    """Dunkle Titelleiste, sobald das Fenster zum ersten Mal erscheint. Ohne das bleibt sie
    weiß, auch wenn Windows auf dunkel steht (DWMWA_USE_IMMERSIVE_DARK_MODE = 20,
    Windows 10 ab 20H1 und Windows 11)."""
    def setzen(e=None):
        if e is not None and e.widget is not fenster:
            return
        try:
            import ctypes
            an = ctypes.c_int(1)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(int(fenster.wm_frame(), 16), 20,
                                                       ctypes.byref(an), ctypes.sizeof(an))
        except Exception:
            pass
    fenster.bind("<Map>", setzen, add="+")


def knopf_farben(t: dict) -> dict:
    """Dunkle tk.Button-Knöpfe in den Farben des Themes."""
    return dict(bg=t["bar_bg"], fg=t["text"], activebackground=t["bar"], activeforeground="#ffffff",
                relief=tk.FLAT, bd=0, highlightthickness=1, highlightbackground=t["muted"],
                highlightcolor=t["title"], cursor="hand2")


class Eingabe(simpledialog.Dialog):
    """Wie simpledialog.askstring, aber dunkel in den Farben des Tickers und mit Knöpfen
    in der eingestellten Sprache (Pythons eigener Dialog ist hellgrau und beschriftet
    fest mit "OK" und "Cancel")."""

    def __init__(self, parent: tk.Misc, titel: str, frage: str, vorgabe: str = "",
                 farben: dict | None = None):
        self.frage = frage
        self.vorgabe = vorgabe
        self.t = farben or THEMES[DEFAULT_THEME]
        super().__init__(parent, titel)

    def body(self, master):
        t = self.t
        self.configure(bg=t["bg"])
        master.configure(bg=t["bg"])
        titelleiste_dunkel(self)
        tk.Label(master, text=self.frage, justify=tk.LEFT, fg=t["text"], bg=t["bg"]).grid(row=0, padx=5, sticky=tk.W)
        self.feld = tk.Entry(master, name="entry", bg=t["bar_bg"], fg=t["text"], insertbackground=t["title"],
                             selectbackground=t["bar"], selectforeground="#ffffff", relief=tk.FLAT,
                             highlightthickness=1, highlightbackground=t["muted"], highlightcolor=t["title"])
        self.feld.grid(row=1, padx=5, pady=(6, 0), ipady=3, sticky=tk.W + tk.E)
        if self.vorgabe:
            self.feld.insert(0, self.vorgabe)
            self.feld.select_range(0, tk.END)
        return self.feld

    def buttonbox(self):
        t = self.t
        box = tk.Frame(self, bg=t["bg"])
        knopf = dict(width=10, **knopf_farben(t))
        tk.Button(box, text=tx("ok"), command=self.ok, default=tk.ACTIVE, **knopf).pack(side=tk.LEFT, padx=5, pady=8)
        tk.Button(box, text=tx("abbrechen"), command=self.cancel, **knopf).pack(side=tk.LEFT, padx=5, pady=8)
        self.bind("<Return>", self.ok)
        self.bind("<Escape>", self.cancel)
        box.pack()

    def apply(self):
        self.result = self.feld.get()


class Hinweis(simpledialog.Dialog):
    """Dunkler Hinweis mit einem Knopf, statt messagebox.showinfo (der Windows-Dialog bleibt hell)."""

    def __init__(self, parent: tk.Misc, titel: str, text: str, farben: dict | None = None):
        self.text = text
        self.t = farben or THEMES[DEFAULT_THEME]
        super().__init__(parent, titel)

    def body(self, master):
        t = self.t
        self.configure(bg=t["bg"])
        master.configure(bg=t["bg"])
        titelleiste_dunkel(self)
        tk.Label(master, text=self.text, justify=tk.LEFT, wraplength=360, fg=t["text"],
                 bg=t["bg"]).grid(row=0, padx=8, pady=(6, 0), sticky=tk.W)

    def buttonbox(self):
        box = tk.Frame(self, bg=self.t["bg"])
        tk.Button(box, text=tx("ok"), command=self.ok, default=tk.ACTIVE, width=10,
                  **knopf_farben(self.t)).pack(padx=5, pady=8)
        self.bind("<Return>", self.ok)
        self.bind("<Escape>", self.cancel)
        box.pack()


def ask_site(parent: tk.Tk | None = None, vorgabe: str = "",
             farben: dict | None = None) -> str | None:
    own_root = False
    if parent is None:
        parent = tk.Tk()
        parent.withdraw()
        own_root = True
    wert = Eingabe(parent, "Besucher-Ticker", tx("frageSite"), vorgabe, farben).result
    if own_root:
        parent.destroy()
    return normalisiere_site(wert) or None


def ask_token(parent: tk.Tk | None = None, farben: dict | None = None) -> str | None:
    """GUI-Dialog statt input() — funktioniert auch unter pythonw.exe."""
    own_root = False
    if parent is None:
        parent = tk.Tk()
        parent.withdraw()
        own_root = True
    token = Eingabe(parent, "Besucher-Ticker", tx("frageToken"), farben=farben).result
    if own_root:
        parent.destroy()
    if token:
        token = token.strip()
    return token or None


# ------------------------------ Beschlag ------------------------------------
# Der Look der Werkstatt: eine Platte mit runder Außenkante, darin ein eckiger
# Rahmen, an jeder Ecke eine kleine Schraube in Weißgold, jede anders gedreht.
# Das Licht kommt von rechts oben, darum fällt der Schatten nach links unten.
#
# tkinter kann weder weiche Schatten noch geglättete Rundungen. Deshalb liegt
# hinter der Anzeige ein zweites Fenster, das nur ein Bild zeigt (Platte,
# Rahmen, Schrauben, Schatten) und dort ein Loch hat, wo die Anzeige sitzt.

B_RADIUS = 14                          # Rundung der Platte
B_INNEN = 5                            # Abstand des Rahmens von der Kante
B_ABSTAND = 3                          # Abstand der Schraube vom Rahmen
B_SCHRAUBE = 10                        # Größe einer Schraube
B_RAND = B_INNEN + B_ABSTAND + B_SCHRAUBE   # so breit ist die Platte um die Anzeige
B_SCHATTEN = (-7, 8)                   # Versatz: nach links unten
B_WEICH = 9                            # Unschärfe des Schattens
B_LUFT = 26                            # Platz für den Schatten um die Platte
B_WINKEL = (32, -58, 78, -14)          # Schlitz je Schraube, in Grad


def _hex(farbe: str) -> tuple[int, int, int]:
    return int(farbe[1:3], 16), int(farbe[3:5], 16), int(farbe[5:7], 16)


def _mische(a: str, b: str, anteil: float) -> tuple[int, int, int]:
    """Farbe a mit dem Anteil von b."""
    x, y = _hex(a), _hex(b)
    return tuple(round(x[i] + (y[i] - x[i]) * anteil) for i in range(3))


def _verlauf(stufen, t: float):
    """Wert eines Farbverlaufs an der Stelle t. stufen: (stelle, (r, g, b, deckung))"""
    if t <= stufen[0][0]:
        return stufen[0][1]
    for (t0, f0), (t1, f1) in zip(stufen, stufen[1:]):
        if t <= t1:
            k = (t - t0) / ((t1 - t0) or 1)
            return tuple(f0[i] + (f1[i] - f0[i]) * k for i in range(4))
    return stufen[-1][1]


_KOPF = ((0.0, (255, 255, 255, 1)), (0.2, (255, 252, 238, 1)), (0.5, (241, 228, 188, 1)),
         (0.8, (196, 169, 100, 1)), (1.0, (122, 98, 38, 1)))
_RING = ((0.0, (138, 106, 34, .55)), (0.5, (255, 243, 196, 0)), (1.0, (255, 251, 234, .7)))
_SCHRAUBEN: dict = {}


def schraube(winkel: float, groesse: int = B_SCHRAUBE):
    """Eine Schraube in Weißgold, Punkt für Punkt gerechnet: Schatten nach
    links unten, gewölbter Kopf, Schlitz im Winkel, Glanzpunkt rechts oben.
    Dieselbe Zeichnung wie auf der Startseite, dort als SVG auf 12 × 12."""
    if (winkel, groesse) in _SCHRAUBEN:
        return _SCHRAUBEN[(winkel, groesse)]
    n = groesse * 6                    # sechsfach gerechnet, dann verkleinert: glatte Kanten
    co, si = math.cos(math.radians(-winkel)), math.sin(math.radians(-winkel))
    gc, gs = math.cos(math.radians(35)), math.sin(math.radians(35))
    punkte = []

    def strich(x, y, x0, x1, y0):
        """Abstand zu einer waagrechten Strecke mit runden Enden"""
        return math.hypot(x - max(x0, min(x1, x)), y - y0)

    for j in range(n):
        v = (j + .5) * 12 / n
        for i in range(n):
            u = (i + .5) * 12 / n
            r = g = b = a = 0.0        # Farbe mal Deckung, wird Schicht für Schicht belegt

            def auf(farbe, deckung):
                nonlocal r, g, b, a
                r = farbe[0] * deckung + r * (1 - deckung)
                g = farbe[1] * deckung + g * (1 - deckung)
                b = farbe[2] * deckung + b * (1 - deckung)
                a = deckung + a * (1 - deckung)

            # Schatten der Schraube auf der Platte
            if math.hypot(u - 5.1, v - 6.9) <= 5:
                t = math.hypot(u - 4.3, v - 7.7) / 5
                auf((0, 0, 0), .6 if t <= .55 else max(0.0, .6 * (1 - (t - .55) / .45)))
            m = math.hypot(u - 6, v - 6)
            # Kopf, heller Fleck rechts oben
            if m <= 4.3:
                f = _verlauf(_KOPF, math.hypot(u - 7.376, v - 4.28) / 6.708)
                auf(f, 1)
            # dunkle Kante
            if abs(m - 4.3) <= .225:
                auf((85, 62, 12), .6)
            # feiner Ring: oben rechts dunkel, unten links hell
            if abs(m - 2.9) <= .25:
                f = _verlauf(_RING, ((8.9 - u) / 5.8 + (v - 3.1) / 5.8) / 2)
                auf(f, f[3])
            # Schlitz und sein heller Rand
            x = 6 + (u - 6) * co - (v - 6) * si
            y = 6 + (u - 6) * si + (v - 6) * co
            if strich(x, y, 2.5, 9.5, 6) <= .6:
                auf((58, 42, 7), 1)
            if strich(x, y, 2.7, 9.3, 6.62) <= .15:
                auf((255, 247, 218), .75)
            # Glanzpunkt
            ex = (u - 7.7) * gc - (v - 3.9) * gs
            ey = (u - 7.7) * gs + (v - 3.9) * gc
            if (ex / 1.05) ** 2 + (ey / .6) ** 2 <= 1:
                auf((255, 255, 255), .9)
            punkte.append((round(r), round(g), round(b), round(a * 255)))

    bild = Image.new("RGBa", (n, n))
    bild.putdata(punkte)
    bild = bild.resize((groesse, groesse), Image.LANCZOS).convert("RGBA")
    _SCHRAUBEN[(winkel, groesse)] = bild
    return bild


def beschlag_bild(w: int, h: int, deckkraft: float, grund: str, linie: str):
    """Das Bild hinter einer Anzeige von w × h Punkten: Schatten, Platte,
    Rahmen, Schrauben. Wo die Anzeige sitzt, bleibt es durchsichtig.
    grund ist die Farbe der Platte, linie färbt Kante und Rahmen (gedämpft)."""
    pw, ph = w + 2 * B_RAND, h + 2 * B_RAND
    bw, bh = pw + 2 * B_LUFT, ph + 2 * B_LUFT
    k = 4                              # Platte vierfach zeichnen, dann verkleinern

    form = Image.new("L", (bw * k, bh * k), 0)
    kante = [B_LUFT * k, B_LUFT * k, (B_LUFT + pw) * k - 1, (B_LUFT + ph) * k - 1]
    ImageDraw.Draw(form).rounded_rectangle(kante, radius=B_RADIUS * k, fill=255)

    gross = Image.new("RGB", (bw * k, bh * k), _hex(grund))
    d = ImageDraw.Draw(gross)
    d.rounded_rectangle(kante, radius=B_RADIUS * k, width=k,
                        outline=_mische(grund, linie, .22))
    i = (B_LUFT + B_INNEN) * k
    d.rectangle([i, i, (B_LUFT + pw - B_INNEN) * k - 1, (B_LUFT + ph - B_INNEN) * k - 1],
                width=k, outline=_mische(grund, linie, .40))
    platte = gross.resize((bw, bh), Image.LANCZOS).convert("RGBA")
    form = form.resize((bw, bh), Image.LANCZOS)
    platte.putalpha(form)

    # Schatten: dieselbe Form, versetzt und weich. Unter der Platte selbst
    # ausgespart, sonst wäre sie dunkler als die Anzeige in ihrer Mitte.
    schatten = ImageChops.offset(form, B_SCHATTEN[0], B_SCHATTEN[1])
    schatten = schatten.filter(ImageFilter.GaussianBlur(B_WEICH)).point(lambda v: int(v * .62))
    schatten = ImageChops.multiply(schatten, ImageChops.invert(form))
    bild = Image.new("RGBA", (bw, bh), (0, 0, 0, 0))
    bild.putalpha(schatten)
    bild.alpha_composite(platte)

    links = B_LUFT + B_INNEN + B_ABSTAND
    rechts = B_LUFT + pw - B_INNEN - B_ABSTAND - B_SCHRAUBE
    unten = B_LUFT + ph - B_INNEN - B_ABSTAND - B_SCHRAUBE
    for (x, y), winkel in zip(((links, links), (rechts, links), (links, unten), (rechts, unten)),
                              B_WINKEL):
        bild.alpha_composite(schraube(winkel), (x, y))

    loch = B_LUFT + B_RAND
    bild.paste((0, 0, 0, 0), (loch, loch, loch + w, loch + h))
    if deckkraft < 1:
        bild.putalpha(bild.getchannel("A").point(lambda v: int(v * deckkraft)))
    return bild


# Windows zeigt ein Bild mit weichen Kanten nur über ein „geschichtetes"
# Fenster. Das Bild wird dem Fenster direkt übergeben, tkinter malt dort nichts.

if os.name == "nt":
    class _PUNKT(ctypes.Structure):
        _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

    class _MASS(ctypes.Structure):
        _fields_ = [("cx", ctypes.c_long), ("cy", ctypes.c_long)]

    class _MISCHUNG(ctypes.Structure):
        _fields_ = [("BlendOp", ctypes.c_ubyte), ("BlendFlags", ctypes.c_ubyte),
                    ("SourceConstantAlpha", ctypes.c_ubyte), ("AlphaFormat", ctypes.c_ubyte)]

    class _BILDKOPF(ctypes.Structure):
        _fields_ = [("biSize", ctypes.c_uint32), ("biWidth", ctypes.c_long),
                    ("biHeight", ctypes.c_long), ("biPlanes", ctypes.c_uint16),
                    ("biBitCount", ctypes.c_uint16), ("biCompression", ctypes.c_uint32),
                    ("biSizeImage", ctypes.c_uint32), ("biXPelsPerMeter", ctypes.c_long),
                    ("biYPelsPerMeter", ctypes.c_long), ("biClrUsed", ctypes.c_uint32),
                    ("biClrImportant", ctypes.c_uint32)]

_WIN = None


def _win():
    """user32 und gdi32 mit sauberen Typen (64-Bit-Zeiger), einmal eingerichtet."""
    global _WIN
    if _WIN is None:
        z, p = ctypes.c_void_p, ctypes.POINTER
        u, g = ctypes.WinDLL("user32"), ctypes.WinDLL("gdi32")
        u.GetAncestor.restype, u.GetAncestor.argtypes = z, [z, ctypes.c_uint]
        u.GetDC.restype, u.GetDC.argtypes = z, [z]
        u.ReleaseDC.argtypes = [z, z]
        u.GetWindowLongPtrW.restype, u.GetWindowLongPtrW.argtypes = ctypes.c_ssize_t, [z, ctypes.c_int]
        u.SetWindowLongPtrW.restype = ctypes.c_ssize_t
        u.SetWindowLongPtrW.argtypes = [z, ctypes.c_int, ctypes.c_ssize_t]
        u.UpdateLayeredWindow.argtypes = [z, z, p(_PUNKT), p(_MASS), z, p(_PUNKT),
                                          ctypes.c_uint32, p(_MISCHUNG), ctypes.c_uint32]
        g.CreateCompatibleDC.restype, g.CreateCompatibleDC.argtypes = z, [z]
        g.CreateDIBSection.restype = z
        g.CreateDIBSection.argtypes = [z, p(_BILDKOPF), ctypes.c_uint, p(z), z, ctypes.c_uint32]
        g.SelectObject.restype, g.SelectObject.argtypes = z, [z, z]
        g.DeleteObject.argtypes = [z]
        g.DeleteDC.argtypes = [z]
        _WIN = (u, g)
    return _WIN


def schicht_vorbereiten(tk_id: int) -> int:
    """Macht aus einem tkinter-Fenster ein geschichtetes. Rückgabe: sein Handle."""
    u, _ = _win()
    hwnd = u.GetAncestor(tk_id, 2) or tk_id
    stil = u.GetWindowLongPtrW(hwnd, -20)
    u.SetWindowLongPtrW(hwnd, -20, stil | 0x00080000 | 0x00000080)   # geschichtet, kein Taskleisten-Knopf
    return hwnd


def schicht_zeigen(hwnd: int, bild, x: int, y: int) -> bool:
    """Legt das Bild (mit Durchsicht je Punkt) an die Stelle x, y auf den Bildschirm."""
    u, g = _win()
    w, h = bild.size
    # Windows will Farbe mal Deckung (gerundet, sonst sieht man die Naht zur
    # Anzeige), Reihenfolge Blau-Grün-Rot
    roh = bild.convert("RGBa").tobytes("raw", "BGRa")
    kopf = _BILDKOPF(ctypes.sizeof(_BILDKOPF), w, -h, 1, 32, 0, 0, 0, 0, 0, 0)
    schirm = u.GetDC(None)
    speicher = g.CreateCompatibleDC(schirm)
    stelle = ctypes.c_void_p()
    karte = g.CreateDIBSection(schirm, ctypes.byref(kopf), 0, ctypes.byref(stelle), None, 0)
    if not karte or not stelle:
        g.DeleteDC(speicher)
        u.ReleaseDC(None, schirm)
        return False
    ctypes.memmove(stelle, roh, len(roh))
    alt = g.SelectObject(speicher, karte)
    ok = u.UpdateLayeredWindow(hwnd, schirm, ctypes.byref(_PUNKT(x, y)), ctypes.byref(_MASS(w, h)),
                               speicher, ctypes.byref(_PUNKT(0, 0)), 0,
                               ctypes.byref(_MISCHUNG(0, 0, 255, 1)), 2)
    g.SelectObject(speicher, alt)
    g.DeleteObject(karte)
    g.DeleteDC(speicher)
    u.ReleaseDC(None, schirm)
    return bool(ok)


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
        self._letzte_anfrage = 0.0
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
        warten = API_ABSTAND - (time.monotonic() - self._letzte_anfrage)
        if warten > 0:
            time.sleep(warten)
        self._letzte_anfrage = time.monotonic()
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

    def hits_alle(self, start: str, end: str,
                  max_seiten: int = HITS_SEITEN_MAX) -> list[dict]:
        """Alle Pfade des Zeitraums, je Anfrage höchstens 100.

        Weitergeblättert wird, wie GoatCounter es vorsieht: Die schon
        gelieferten Pfade gehen als `exclude_paths` (ihre IDs, mit Komma
        getrennt) in die nächste Anfrage, bis `more` falsch ist."""
        alle: list[dict] = []
        gesehen: list[str] = []
        for _ in range(max_seiten):
            params = {"start": start, "end": end, "limit": 100}
            if gesehen:
                params["exclude_paths"] = ",".join(gesehen)
            antwort = self._get("/stats/hits", params) or {}
            neu = [h for h in antwort.get("hits") or [] if isinstance(h, dict)]
            alle.extend(neu)
            ids = [str(h["path_id"]) for h in neu if h.get("path_id") is not None]
            if not antwort.get("more") or not ids:
                break
            gesehen.extend(ids)
        return alle

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


def _hitliste(hits) -> list[dict]:
    """Nimmt die Antwort von /stats/hits oder schon die Liste daraus."""
    if isinstance(hits, dict):
        hits = hits.get("hits")
    return [h for h in hits or [] if isinstance(h, dict)]


def auswerten_hits(hits, max_balken: int) -> list[dict]:
    """Zählweise Seitenbesuche: Pfadliste zu Werkzeug-Balken zusammenfassen.

    Ereignisse (event: true) fliegen raus, sie sind Klicks und keine
    Seiten. Zählt eine Seite beides, kämen sonst Klicks auf die Seiten-
    besuche drauf."""
    summen: dict[str, int] = {}
    for hit in _hitliste(hits):
        if hit.get("event"):
            continue
        label = pfad_label(hit.get("path", "/"))
        summen[label] = summen.get(label, 0) + _zahl(hit.get("count"))
    reihen = [{"label": k, "count": v} for k, v in summen.items() if v > 0]
    reihen.sort(key=lambda r: (-r["count"], r["label"]))
    return reihen[:max_balken]


def besucher_seiten(total_payload: dict) -> int:
    """Zählweise Seitenbesuche: `total` zählt die Ereignisse mit, die
    gehören aber nicht dazu. Ältere GoatCounter-Fassungen ohne
    `total_events` liefern dort nichts, dann bleibt `total`."""
    t = total_payload or {}
    return max(0, _zahl(t.get("total")) - _zahl(t.get("total_events")))


def kachel_schluessel(pfad: str) -> str | None:
    """'kachel-ram' und 'kachel-ram-anleitung' -> 'ram', sonst None."""
    p = (pfad or "").strip().lstrip("/")
    if not p.startswith(KACHEL_PRAEFIX):
        return None
    return p[len(KACHEL_PRAEFIX):].split("-", 1)[0] or None


def ist_kachel_klick(hit: dict) -> bool:
    return bool(hit.get("event")) and kachel_schluessel(hit.get("path", "")) is not None


def hat_kachel_klicks(hits) -> bool:
    return any(ist_kachel_klick(h) and _zahl(h.get("count")) > 0
               for h in _hitliste(hits))


def auswerten_klicks(hits, max_balken: int) -> tuple[list[dict], int]:
    """Zählweise Kachel-Klicks: (Balken je Werkzeug, Summe aller Kacheln).

    Große Kachel und Mini-Kacheln eines Werkzeugs landen im selben
    Balken. Die Beschriftung ist der Titel, den die Seite mitschickt
    (der Werkzeugname auf der Kachel), ersatzweise der Schlüssel. Die
    Summe zählt alle Kacheln, auch die, die nicht mehr als Balken
    hineinpassen."""
    summen: dict[str, int] = {}
    namen: dict[str, str] = {}
    for hit in _hitliste(hits):
        if not ist_kachel_klick(hit):
            continue
        pfad = (hit.get("path") or "").strip().lstrip("/")
        key = kachel_schluessel(pfad)
        summen[key] = summen.get(key, 0) + _zahl(hit.get("count"))
        titel = (hit.get("title") or "").strip()
        # Die große Kachel hat Vorrang; ein Titel, der nur der Pfad ist,
        # taugt nicht als Name.
        if titel and titel != pfad and (key not in namen
                                        or pfad == KACHEL_PRAEFIX + key):
            namen[key] = titel
    reihen = [{"label": namen.get(k, k), "count": v}
              for k, v in summen.items() if v > 0]
    reihen.sort(key=lambda r: (-r["count"], r["label"]))
    return reihen[:max_balken], sum(summen.values())


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


_TRENNER = {"en": ",", "hi": ",", "zh": ",", "ja": ",", "ko": ",", "fr": "\u202f", "ru": "\u00a0", "pl": "\u00a0"}


def zahl_kurz(n: int) -> str:
    """Tausendertrennung in der Schreibweise der Sprache, ohne locale-Abhängigkeit."""
    return f"{n:,}".replace(",", _TRENNER.get(SPRACHE, "."))


# ------------------------------ App -----------------------------------------


class TickerApp:
    def __init__(self):
        global SPRACHE
        self.cfg = load_config()
        SPRACHE = waehle_sprache(self.cfg)

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
        self.zaehlart = self.cfg.get("zaehlart", DEFAULT_ZAEHLART)
        if self.zaehlart not in [k for k, _ in ZAEHLARTEN]:
            self.zaehlart = DEFAULT_ZAEHLART

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
        # Dasselbe für die Pfadliste seit Anbeginn: Sie liefert bei den
        # Kachel-Klicks die Gesamtsumme und entscheidet "automatisch".
        self._alltime_hits: list[dict] | None = None
        self._alltime_hits_zeit = 0.0
        # Zählt jeden Zugangswechsel. Ein Abruf, der noch für die alte Seite
        # lief, legt sein Ergebnis dann nicht mehr ab.
        self._zugang_gen = 0
        # "automatisch" hat schon einmal Kachel-Klicks gesehen (bleibt so).
        self._auto_klicks = False

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

    def zeitraum_name(self, key: str | None = None) -> str:
        key = key or self.zeitraum
        return tx(next((n for k, n, _ in ZEITRAEUME if k == key), "zeitHeute"))

    def zusatz_page(self) -> str | None:
        return next((p for k, _, p in ZUSATZ_CHOICES if k == self.zusatz), None)

    def zusatz_name(self) -> str:
        return tx(next((n for k, n, _ in ZUSATZ_CHOICES if k == self.zusatz), "zusAus"))

    def zaehlart_name(self) -> str:
        return tx(next((n for k, n in ZAEHLARTEN if k == self.zaehlart), "artAuto"))

    # ---------- UI ----------

    def _build_ui(self):
        self.root = tk.Tk()
        self.root.title(tx("titel"))
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.attributes("-alpha", self.alpha_pct / 100.0)
        self.root.configure(bg=self.th["bg"])

        x, y = clamp_to_screen(int(self.cfg.get("pos_x", 60)),
                               int(self.cfg.get("pos_y", 60)))
        self.cfg["pos_x"], self.cfg["pos_y"] = x, y
        self.root.geometry(f"+{x}+{y}")

        self._baue_platte()
        # Mit Beschlag hält die Platte den Abstand zum Rand, ohne ihn der Rahmen selbst.
        self.frame = tk.Frame(self.root, bg=self.th["bg"],
                              padx=3 if self.platte else 10, pady=1 if self.platte else 8)
        self.frame.pack()

        self.title_lbl = tk.Label(self.frame, text=tx("titel"),
                                  fg=self.th["title"], bg=self.th["bg"],
                                  font=("Segoe UI", 9, "bold"))
        self.title_lbl.pack(anchor="w")

        self.body = tk.Frame(self.frame, bg=self.th["bg"])
        self.body.pack(fill="x")

        self.status_lbl = tk.Label(self.frame, text=tx("lade"),
                                   fg=self.th["muted"], bg=self.th["bg"],
                                   font=("Segoe UI", 8))
        self.status_lbl.pack(anchor="w")

        self._bind_drag_recursive(self.root)
        self.root.bind("<Configure>", self._root_bewegt, add="+")
        self.root.protocol("WM_DELETE_WINDOW", self.quit)

    # ---------- Beschlag ----------

    def _baue_platte(self):
        """Das Fenster hinter der Anzeige: Platte, Rahmen, Schrauben, Schatten."""
        self.platte = None
        self._platte_hwnd = None
        self._platte_bild = None
        self._platte_mass = None
        self._platte_stand = None
        if os.name != "nt":
            return
        try:
            p = tk.Toplevel(self.root)
            p.withdraw()
            p.overrideredirect(True)
            p.attributes("-topmost", True)
            p.configure(bg=self.th["bg"])
            p.update_idletasks()
            self._platte_hwnd = schicht_vorbereiten(p.winfo_id())
            self.platte = p
        except Exception:               # ohne Beschlag läuft die Anzeige trotzdem
            self.platte = None

    def _root_bewegt(self, e):
        if e.widget is self.root:
            self._platte_folgen()

    def _platte_folgen(self, x=None, y=None):
        """Zieht die Platte hinter die Anzeige: neue Stelle, bei Bedarf neues Bild
        (andere Größe, anderes Theme, andere Deckkraft)."""
        if not self.platte or not self.visible:
            return
        w, h = self.root.winfo_width(), self.root.winfo_height()
        if w <= 1 or h <= 1:
            return
        if x is None:
            x, y = self.root.winfo_x(), self.root.winfo_y()
        mass = (w, h, self.alpha_pct, self.theme_name)
        if self._platte_mass != mass:
            self._platte_bild = beschlag_bild(w, h, self.alpha_pct / 100.0,
                                              self.th["bg"], self.th["title"])
            self._platte_mass = mass
        px, py = x - B_RAND - B_LUFT, y - B_RAND - B_LUFT
        if self._platte_stand == (mass, px, py):
            return
        try:
            bw, bh = self._platte_bild.size
            self.platte.geometry(f"{bw}x{bh}+{px}+{py}")
            if schicht_zeigen(self._platte_hwnd, self._platte_bild, px, py):
                self._platte_stand = (mass, px, py)
            if self.platte.state() == "withdrawn":
                self.platte.deiconify()
                self.root.lift()
        except (tk.TclError, OSError):
            pass

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
        self._platte_folgen(x, y)
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
        self._platte_folgen()

    def _render(self):
        t = self.th
        for w in self.body.winfo_children():
            w.destroy()

        self.title_lbl.config(text=f"{tx('titel')} · "
                                   f"{self.zeitraum_name(self.daten.get('zeitraum'))}")

        if not self.daten:
            self._bind_drag_recursive(self.root)
            return

        heute = self.daten.get("heute", 0)
        gesamt = self.daten.get("gesamt")
        zeitraum_total = self.daten.get("zeitraum_total", 0)
        # Beschriftet wird, was tatsächlich gezählt wurde, auch bei
        # "automatisch".
        klicks = self.daten.get("art") == "klicks"
        gross = tx("artKlicks" if klicks else "artSeiten")
        # Die Zahlen gehören zu dem Zeitraum, mit dem sie geholt wurden. Bis
        # der nächste Durchlauf da ist, bleibt deshalb dessen Name stehen.
        zr_daten = self.daten.get("zeitraum", self.zeitraum)

        if self.compact:
            tk.Label(self.body, text=zahl_kurz(zeitraum_total),
                     fg=t["big"], bg=t["bg"],
                     font=("Segoe UI", 16, "bold")).pack(anchor="w")
            tk.Label(self.body, text=f"{gross} · {self.zeitraum_name(zr_daten)}",
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
        tk.Label(rechts, text=f"{gross} · {self.zeitraum_name(zr_daten)}",
                 fg=t["muted"], bg=t["bg"],
                 font=("Segoe UI", 8)).pack(anchor="e")
        if zr_daten != "heute":
            tk.Label(rechts, text=tx("heuteN", n=zahl_kurz(heute)),
                     fg=t["text"], bg=t["bg"],
                     font=("Segoe UI", 9, "bold")).pack(anchor="e")
        if gesamt is not None:
            tk.Label(rechts, text=tx("gesamtN", n=zahl_kurz(gesamt)),
                     fg=t["text"], bg=t["bg"],
                     font=("Segoe UI", 9, "bold")).pack(anchor="e")

        # --- Balken je Werkzeug ---
        balken = self.daten.get("balken") or []
        if balken:
            tk.Label(self.body, text=tx("jeKachel" if klicks else "jeWerkzeug"),
                     fg=t["muted"], bg=t["bg"],
                     font=("Segoe UI", 8)).pack(anchor="w", pady=(6, 0))
            hoechst = max(b["count"] for b in balken) or 1
            for b in balken:
                row = tk.Frame(self.body, bg=t["bg"])
                row.pack(fill="x", pady=(3, 0))
                head = tk.Frame(row, bg=t["bg"])
                head.pack(fill="x")
                tk.Label(head, text=anzeige(b["label"]), fg=t["text"], bg=t["bg"],
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
                tk.Label(row, text=anzeige(z["label"])[:22], fg=t["text"], bg=t["bg"],
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
                tx(name),
                (lambda k: (lambda *_: self.ui_queue.put(("zeitraum", k))))(key),
                checked=(lambda k: (lambda _i: self.zeitraum == k))(key),
                radio=True)
            for key, name, _ in ZEITRAEUME
        ])
        zusatz_menu = pystray.Menu(*[
            pystray.MenuItem(
                tx(name),
                (lambda k: (lambda *_: self.ui_queue.put(("zusatz", k))))(key),
                checked=(lambda k: (lambda _i: self.zusatz == k))(key),
                radio=True)
            for key, name, _ in ZUSATZ_CHOICES
        ])
        zaehlart_menu = pystray.Menu(*[
            pystray.MenuItem(
                tx(name),
                (lambda k: (lambda *_: self.ui_queue.put(("zaehlart", k))))(key),
                checked=(lambda k: (lambda _i: self.zaehlart == k))(key),
                radio=True)
            for key, name in ZAEHLARTEN
        ])
        sprachen_menu = pystray.Menu(*[
            pystray.MenuItem(
                TEXTE[code]["_name"],
                (lambda c: (lambda *_: self.ui_queue.put(("sprache", c))))(code),
                checked=(lambda c: (lambda _i: SPRACHE == c))(code),
                radio=True)
            for code in TEXTE
        ])
        menu = pystray.Menu(
            pystray.MenuItem(tx("mOverlay"), self._tray_toggle, default=True),
            pystray.MenuItem(tx("mMitte"), self._tray_center),
            pystray.MenuItem(tx("mJetzt"), self._tray_refresh),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(tx("mZeitraum"), zeitraum_menu),
            pystray.MenuItem(tx("mZusatz"), zusatz_menu),
            pystray.MenuItem(tx("mZaehlart"), zaehlart_menu),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(tx("mEinstellungen"), self._tray_settings),
            pystray.MenuItem(tx("mTheme"), self._tray_next_theme),
            pystray.MenuItem(tx("mZugang"), self._tray_change_zugang),
            pystray.MenuItem(tx("mSprache"), sprachen_menu),
            pystray.MenuItem(tx("mBeenden"), self._tray_quit),
        )
        if self.icon:                       # Sprachwechsel: altes Symbol abbauen
            try:
                self.icon.stop()
            except Exception:
                pass
        self.icon = pystray.Icon("besucher_ticker", self._tray_image(),
                                 "Besucher-Ticker", menu)
        threading.Thread(target=self.icon.run, daemon=True).start()

    def _tray_toggle(self, *_):
        self.ui_queue.put(("toggle", None))

    def _tray_center(self, *_):
        self.ui_queue.put(("center", None))

    def _tray_refresh(self, *_):
        self.wake_event.set()

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
        win.title(tx("setTitel"))
        win.configure(bg=t["bg"], padx=14, pady=12)
        win.attributes("-topmost", True)
        win.resizable(False, False)
        win.geometry(f"+{self.root.winfo_x() + 40}+{self.root.winfo_y() + 40}")
        titelleiste_dunkel(win)

        def lbl(text, row, pady=(8, 2)):
            tk.Label(win, text=text, fg=t["title"], bg=t["bg"],
                     font=("Segoe UI", 9, "bold")).grid(
                row=row, column=0, columnspan=2, sticky="w", pady=pady)

        style = ttk.Style(win)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        # Auswahllisten dunkel: Feld, Pfeil und die aufgeklappte Liste (ein tk.Listbox,
        # der nur über die Optionsdatenbank Farben bekommt)
        style.configure("TCombobox", fieldbackground=t["bar_bg"], background=t["bar_bg"],
                        foreground=t["text"], arrowcolor=t["title"], bordercolor=t["muted"],
                        lightcolor=t["bar_bg"], darkcolor=t["bar_bg"], insertcolor=t["title"])
        style.map("TCombobox",
                  fieldbackground=[("readonly", t["bar_bg"])],
                  foreground=[("readonly", t["text"])],
                  background=[("pressed", t["bar"]), ("active", t["bar"])],
                  arrowcolor=[("pressed", "#ffffff"), ("active", "#ffffff")],
                  selectbackground=[("readonly", t["bar_bg"])],
                  selectforeground=[("readonly", t["text"])])
        win.option_add("*TCombobox*Listbox.background", t["bar_bg"])
        win.option_add("*TCombobox*Listbox.foreground", t["text"])
        win.option_add("*TCombobox*Listbox.selectBackground", t["bar"])
        win.option_add("*TCombobox*Listbox.selectForeground", "#ffffff")

        # Auswahllisten 30 Zeichen breit wie bisher, breiter nur, wenn ein
        # Eintrag in der gewählten Sprache sonst abgeschnitten würde.
        schrift = tkfont.nametofont("TkTextFont")
        alle = ([refresh_name(s) for s in REFRESH_CHOICES]
                + [tx(n) for _, n, _ in ZEITRAEUME + ZUSATZ_CHOICES]
                + [tx(n) for _, n in ZAEHLARTEN]
                + [dauer_name(s) for s in COMPACT_DURATIONS]
                + [theme_anzeige(n) for n in THEME_NAMES])
        breite = max(30, -(-max(schrift.measure(w) for w in alle)
                           // schrift.measure("0")) + 1)

        # --- Update-Frequenz ---
        lbl(tx("setFreq"), 0, pady=(0, 2))
        self._refresh_var = tk.StringVar()
        current_label = next(
            (refresh_name(sec) for sec in REFRESH_CHOICES
             if abs(sec - self.refresh_seconds) < 0.01),
            f"{self.refresh_seconds:g} s")
        self._refresh_var.set(current_label)
        cb_refresh = ttk.Combobox(win, textvariable=self._refresh_var,
                                  values=[refresh_name(s) for s in REFRESH_CHOICES],
                                  state="readonly", width=breite)
        cb_refresh.grid(row=1, column=0, columnspan=2, sticky="we")

        def on_refresh(_e=None):
            sel = self._refresh_var.get()
            for sec in REFRESH_CHOICES:
                if refresh_name(sec) == sel:
                    self.refresh_seconds = float(sec)
                    self.cfg["refresh_seconds"] = self.refresh_seconds
                    save_config(self.cfg)
                    self.wake_event.set()
                    break
        cb_refresh.bind("<<ComboboxSelected>>", on_refresh)

        tk.Label(win, text=tx("setFreqHinweis"),
                 fg=t["muted"], bg=t["bg"], font=("Segoe UI", 8),
                 justify="left").grid(row=2, column=0, columnspan=2, sticky="w")

        # --- Zeitraum ---
        lbl(tx("setZeitraum"), 3)
        self._zr_var = tk.StringVar(value=self.zeitraum_name())
        cb_zr = ttk.Combobox(win, textvariable=self._zr_var,
                             values=[tx(n) for _, n, _ in ZEITRAEUME],
                             state="readonly", width=breite)
        cb_zr.grid(row=4, column=0, columnspan=2, sticky="we")

        def on_zr(_e=None):
            sel = self._zr_var.get()
            for key, name, _ in ZEITRAEUME:
                if tx(name) == sel:
                    self._set_zeitraum(key)
                    break
        cb_zr.bind("<<ComboboxSelected>>", on_zr)

        # --- Zusatzblock ---
        lbl(tx("setZusatz"), 5)
        self._zus_var = tk.StringVar(value=self.zusatz_name())
        cb_zus = ttk.Combobox(win, textvariable=self._zus_var,
                              values=[tx(n) for _, n, _ in ZUSATZ_CHOICES],
                              state="readonly", width=breite)
        cb_zus.grid(row=6, column=0, columnspan=2, sticky="we")

        def on_zus(_e=None):
            sel = self._zus_var.get()
            for key, name, _ in ZUSATZ_CHOICES:
                if tx(name) == sel:
                    self._set_zusatz(key)
                    break
        cb_zus.bind("<<ComboboxSelected>>", on_zus)

        # --- Zählweise ---
        lbl(tx("setZaehlart"), 7)
        self._art_var = tk.StringVar(value=self.zaehlart_name())
        cb_art = ttk.Combobox(win, textvariable=self._art_var,
                              values=[tx(n) for _, n in ZAEHLARTEN],
                              state="readonly", width=breite)
        cb_art.grid(row=8, column=0, columnspan=2, sticky="we")

        def on_art(_e=None):
            sel = self._art_var.get()
            for key, name in ZAEHLARTEN:
                if tx(name) == sel:
                    self._set_zaehlart(key)
                    break
        cb_art.bind("<<ComboboxSelected>>", on_art)

        # --- Anzahl Balken ---
        lbl(tx("setBalken"), 9)
        self._bal_var = tk.StringVar(value=str(self.max_balken))
        cb_bal = ttk.Combobox(win, textvariable=self._bal_var,
                              values=[str(n) for n in MAX_BALKEN_CHOICES],
                              state="readonly", width=breite)
        cb_bal.grid(row=10, column=0, columnspan=2, sticky="we")

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
        lbl(tx("setAlpha", min=ALPHA_MIN, max=ALPHA_MAX), 11)
        self._alpha_var = tk.IntVar(value=self.alpha_pct)

        def on_alpha(_v=None):
            self.alpha_pct = int(self._alpha_var.get())
            self.root.attributes("-alpha", self.alpha_pct / 100.0)
            self._platte_folgen()
            self.cfg["alpha_pct"] = self.alpha_pct
            save_config(self.cfg)

        tk.Scale(win, from_=ALPHA_MIN, to=ALPHA_MAX, orient="horizontal",
                 variable=self._alpha_var, command=on_alpha, bg=t["bg"],
                 fg=t["text"], troughcolor=t["bar_bg"], highlightthickness=0,
                 activebackground=t["bar"],
                 length=230).grid(row=12, column=0, columnspan=2, sticky="we")

        # --- Mini-/Maxi-Modus ---
        lbl(tx("setMiniMaxi"), 13)
        btn_frame = tk.Frame(win, bg=t["bg"])
        btn_frame.grid(row=14, column=0, columnspan=2, sticky="w")
        tk.Button(btn_frame, text=tx("mini"), width=8,
                  command=lambda: self._set_compact(True, from_user=True),
                  **knopf_farben(t)).pack(side="left", padx=(0, 6))
        tk.Button(btn_frame, text=tx("maxi"), width=8,
                  command=lambda: self._set_compact(False, from_user=True),
                  **knopf_farben(t)).pack(side="left")

        tk.Label(win, text=tx("setZurueck"), fg=t["text"],
                 bg=t["bg"], font=("Segoe UI", 8)).grid(
            row=15, column=0, columnspan=2, sticky="w", pady=(6, 2))
        self._dur_var = tk.StringVar(
            value=dauer_name(self.compact_duration
                             if self.compact_duration in COMPACT_DURATIONS else None))
        cb_dur = ttk.Combobox(win, textvariable=self._dur_var,
                              values=[dauer_name(s) for s in COMPACT_DURATIONS],
                              state="readonly", width=breite)
        cb_dur.grid(row=16, column=0, columnspan=2, sticky="we")

        def on_dur(_e=None):
            sel = self._dur_var.get()
            for sec in COMPACT_DURATIONS:
                if dauer_name(sec) == sel:
                    self.compact_duration = sec
                    self.cfg["compact_duration"] = sec
                    save_config(self.cfg)
                    break
        cb_dur.bind("<<ComboboxSelected>>", on_dur)

        # --- Theme ---
        lbl(tx("setFarben", n=len(THEME_NAMES)), 17)
        self._theme_var = tk.StringVar(value=theme_anzeige(self.theme_name))
        cb_theme = ttk.Combobox(win, textvariable=self._theme_var,
                                values=[theme_anzeige(n) for n in THEME_NAMES],
                                state="readonly", width=breite)
        cb_theme.grid(row=18, column=0, columnspan=2, sticky="we")

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
                      lambda _e: set_theme(next(
                          (n for n in THEME_NAMES
                           if theme_anzeige(n) == self._theme_var.get()), "")))

        tk.Button(win, text=tx("schliessen"), command=win.destroy, width=12,
                  **knopf_farben(t)).grid(row=19, column=0, columnspan=2, pady=(12, 0))

        tk.Label(win, text="© 2026 Dennis_mit_2n", fg=t["muted"], bg=t["bg"],
                 font=("Segoe UI", 8)).grid(
            row=20, column=0, columnspan=2, pady=(6, 0))

        def on_close():
            self.settings_win = None
            win.destroy()
        win.protocol("WM_DELETE_WINDOW", on_close)

    def _set_zeitraum(self, key: str):
        self.zeitraum = key
        self.cfg["zeitraum"] = key
        save_config(self.cfg)
        self._render()
        self.wake_event.set()               # siehe _refresh_loop

    def _set_zusatz(self, key: str):
        self.zusatz = key
        self.cfg["zusatz"] = key
        save_config(self.cfg)
        self._render()
        self.wake_event.set()

    def _set_zaehlart(self, key: str):
        if key not in [k for k, _ in ZAEHLARTEN]:
            return
        self.zaehlart = key
        self.cfg["zaehlart"] = key
        save_config(self.cfg)
        self._alltime_wert = None           # Gesamtsumme gehört zur Zählweise
        self.wake_event.set()

    def _next_theme(self):
        idx = (THEME_NAMES.index(self.theme_name) + 1) % len(THEME_NAMES)
        self.theme_name = THEME_NAMES[idx]
        self.cfg["theme"] = self.theme_name
        save_config(self.cfg)
        self.apply_theme()
        if self.icon:
            self.icon.icon = self._tray_image()

    def _set_sprache(self, code: str):
        global SPRACHE
        if code not in TEXTE or code == SPRACHE:
            return
        SPRACHE = code
        self.cfg["sprache"] = code
        save_config(self.cfg)
        self._render()
        self._build_tray()                  # Menübeschriftungen neu aufbauen
        if self.settings_win is not None and self.settings_win.winfo_exists():
            self.settings_win.destroy()
            self.settings_win = None
            self._open_settings()
        self.wake_event.set()               # Statuszeile folgt mit dem nächsten Abruf

    # ---------- Worker ----------

    def _refresh_loop(self):
        # Einziger Ort, der abruft. Wer neue Zahlen will (Zeitraum, Zusatz,
        # Zählweise, Zugang, "Jetzt aktualisieren"), setzt wake_event: Läuft
        # gerade ein Durchlauf, folgt der nächste direkt danach, statt dass
        # ein zweiter Thread am Schloss abprallt und die Änderung verpufft.
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
            zeitraum = self.zeitraum           # einmal lesen: die Oberfläche
            max_balken = self.max_balken       # darf sie währenddessen ändern
            heute_start, heute_ende = zeitfenster("heute", jetzt)
            zr_start, zr_ende = zeitfenster(zeitraum, jetzt)
            art = self._zaehlart_wirksam(heute_ende)

            if art == "klicks":
                balken, heute_total = auswerten_klicks(
                    self.api.hits_alle(heute_start, heute_ende), max_balken)
                if zeitraum == "heute":
                    zeitraum_total = heute_total
                elif zeitraum == "gesamt":
                    balken, zeitraum_total = auswerten_klicks(
                        self._alltime_hits_holen(heute_ende, frisch=True),
                        max_balken)
                else:
                    balken, zeitraum_total = auswerten_klicks(
                        self.api.hits_alle(zr_start, zr_ende), max_balken)
                gesamt = auswerten_klicks(
                    self._alltime_hits_holen(heute_ende), 1)[1]
            else:
                heute_total = besucher_seiten(self.api.total(heute_start, heute_ende))
                if zeitraum == "heute":
                    zeitraum_total = heute_total
                else:
                    zeitraum_total = besucher_seiten(self.api.total(zr_start, zr_ende))
                balken = auswerten_hits(self.api.hits(zr_start, zr_ende, limit=100),
                                        max_balken)
                gesamt = self._gesamt_holen(heute_ende)

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

            self.ui_queue.put(("data", {
                "art": art,
                "zeitraum": zeitraum,
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

    def _zugang_neu(self):
        """Andere Seite oder anderes Token: alles Gemerkte gehört zur alten."""
        self._zugang_gen += 1
        self._alltime_wert = None
        self._alltime_hits = None
        self._auto_klicks = False

    def _gesamt_holen(self, ende: str) -> int | None:
        """Gesamtsumme der Seitenbesuche, höchstens alle 10 Minuten frisch."""
        if (self._alltime_wert is not None
                and time.monotonic() - self._alltime_zeit < ALLTIME_CACHE_SECONDS):
            return self._alltime_wert
        gen = self._zugang_gen
        try:
            wert = besucher_seiten(self.api.total(ALLTIME_START, ende))
        except (requests.RequestException, ValueError, LookupError):
            return self._alltime_wert          # alter Wert ist besser als keiner
        if gen != self._zugang_gen:
            return wert                        # Zugang inzwischen gewechselt
        self._alltime_wert = wert
        self._alltime_zeit = time.monotonic()
        return wert

    def _alltime_hits_holen(self, ende: str, frisch: bool = False) -> list[dict]:
        """Alle Pfade seit Anbeginn, höchstens alle 10 Minuten frisch geholt.
        frisch=True (Zeitraum "Gesamt") nimmt nur eine Liste aus demselben
        Durchlauf, sonst wird neu geholt. Scheitert der Abruf, gilt die alte
        Liste; gibt es noch keine, scheitert der Durchlauf und wird wiederholt."""
        alter = time.monotonic() - self._alltime_hits_zeit
        if (self._alltime_hits is not None
                and alter < (FRISCH_SEKUNDEN if frisch else ALLTIME_CACHE_SECONDS)):
            return self._alltime_hits
        gen = self._zugang_gen
        try:
            hits = self.api.hits_alle(ALLTIME_START, ende)
        except (requests.RequestException, ValueError, LookupError):
            if self._alltime_hits is None:
                raise
            return self._alltime_hits
        if gen != self._zugang_gen:
            return hits                        # Zugang inzwischen gewechselt
        self._alltime_hits = hits
        self._alltime_hits_zeit = time.monotonic()
        return hits

    def _zaehlart_wirksam(self, ende: str) -> str:
        """'seiten' oder 'klicks'. Bei "automatisch": Kachel-Klicks, sobald
        die Seite je einen gemeldet hat. Das kippt nicht hin und her, die
        Liste seit Anbeginn wird nur länger."""
        art = self.zaehlart
        if art in ("seiten", "klicks"):
            return art
        try:
            if hat_kachel_klicks(self._alltime_hits_holen(ende)):
                self._auto_klicks = True
        except (requests.RequestException, ValueError, LookupError):
            # Ohne Liste lässt sich nichts Neues entscheiden. Seitenbesuche
            # brauchen sie nicht, also daran soll der Durchlauf nicht scheitern.
            pass
        return "klicks" if self._auto_klicks else "seiten"

    # ---------- Queue / Mainloop ----------

    def _process_queue(self):
        try:
            while True:
                kind, payload = self.ui_queue.get_nowait()
                if kind == "data":
                    self.daten = payload
                    self.status_lbl.config(
                        text=tx("aktualisiert", zeit=datetime.now().strftime("%H:%M:%S")),
                        fg=self.th["muted"])
                    self._render()
                    if self.icon:
                        self.icon.icon = self._tray_image()
                elif kind == "net_error":
                    self.status_lbl.config(
                        text=tx("netzFehler"),
                        fg="#ff9f43")
                elif kind == "rate_limit":
                    txt = tx("ratePause")
                    if payload:
                        txt += f" {int(payload)}s"
                    self.status_lbl.config(text=txt, fg="#ff4d4d")
                elif kind == "auth_error":
                    self.status_lbl.config(text=tx("tokenUngueltig"),
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
                elif kind == "zaehlart":
                    self._set_zaehlart(payload)
                elif kind == "toggle":
                    self._toggle_visible()
                elif kind == "center":
                    self._center_overlay()
                elif kind == "sprache":
                    self._set_sprache(payload)
                elif kind == "quit":
                    self.quit()
                    return
        except queue.Empty:
            pass
        self.root.after(200, self._process_queue)

    def _change_zugang_dialog(self):
        site = ask_site(self.root, self.cfg.get("site", ""), self.th)
        if not site:
            return
        token = ask_token(self.root, self.th)
        if not token:
            return
        self.cfg["site"] = site
        self.cfg["api_token"] = token
        save_config(self.cfg)
        self.api.set_zugang(site, token)
        self._zugang_neu()
        self.wake_event.set()

    def _toggle_visible(self):
        self.visible = not self.visible
        if self.visible:
            self.root.deiconify()
            self.root.attributes("-topmost", True)
            self._platte_stand = None
            self.root.after(30, self._platte_folgen)
        else:
            self.root.withdraw()
            if self.platte:
                self.platte.withdraw()

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

    # Zählweise Seitenbesuche: Ereignisse dürfen NICHT als Werkzeug-Balken
    # erscheinen, sie sind Klicks und keine Seiten.
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
    pruefe("balken liste", auswerten_hits(payload["hits"], 8), balken)

    # Seitenbesuche: total zählt die Ereignisse mit, die gehören nicht dazu.
    pruefe("seiten ohne events", besucher_seiten({"total": 50, "total_events": 12}), 38)
    pruefe("seiten alt", besucher_seiten({"total": 50}), 50)
    pruefe("seiten leer", besucher_seiten({}), 0)

    # Kachel-Klicks: Schlüssel aus dem Ereignispfad
    pruefe("kachel gross", kachel_schluessel("kachel-ram"), "ram")
    pruefe("kachel mini", kachel_schluessel("kachel-ram-anleitung"), "ram")
    pruefe("kachel slash", kachel_schluessel("/kachel-wifi-readme"), "wifi")
    pruefe("kachel fremd", kachel_schluessel("werkstatt-ram-news"), None)
    pruefe("kachel leer", kachel_schluessel("kachel-"), None)
    pruefe("kachel seite", kachel_schluessel("/kachel/"), None)

    klicks = {"hits": [
        {"path": "/", "count": 40, "event": False, "title": "Werkstatt"},
        {"path": "kachel-ram", "count": 9, "event": True, "title": "Real_RAM_cooler"},
        {"path": "kachel-ram-anleitung", "count": 3, "event": True, "title": "Real_RAM_cooler"},
        {"path": "kachel-ram-news", "count": 1, "event": True, "title": "kachel-ram-news"},
        {"path": "kachel-wifi-readme", "count": 4, "event": True, "title": "WiFi QR Card"},
        {"path": "kachel-bigday", "count": 2, "event": True, "title": ""},
        {"path": "kachel-leer", "count": 0, "event": True, "title": "Leer"},
        {"path": "werkstatt-ram-news", "count": 77, "event": True, "title": "x"},
        {"path": "kachel-seite", "count": 5, "event": False, "title": "keine Kachel"},
    ]}
    kb, ksum = auswerten_klicks(klicks, 8)
    pruefe("klicks balken", [(b["label"], b["count"]) for b in kb],
           [("Real_RAM_cooler", 13), ("WiFi QR Card", 4), ("bigday", 2)])
    pruefe("klicks summe", ksum, 19)
    kb1, ksum1 = auswerten_klicks(klicks["hits"], 1)
    pruefe("klicks grenze", (len(kb1), ksum1), (1, 19))
    pruefe("klicks keine", auswerten_klicks({"hits": []}, 8), ([], 0))
    # Titel der großen Kachel schlägt den einer Mini-Kachel, egal in
    # welcher Reihenfolge sie kommen.
    vorrang = [{"path": "kachel-ram-news", "count": 5, "event": True, "title": "Alt"},
               {"path": "kachel-ram", "count": 1, "event": True, "title": "Real_RAM_cooler"}]
    pruefe("klicks titel", auswerten_klicks(vorrang, 8)[0][0]["label"], "Real_RAM_cooler")
    pruefe("hat klicks", hat_kachel_klicks(klicks), True)
    pruefe("hat keine", hat_kachel_klicks({"hits": [
        {"path": "kachel-ram", "count": 0, "event": True},
        {"path": "werkstatt-ram-news", "count": 3, "event": True},
        {"path": "kachel-x", "count": 3, "event": False}]}), False)

    # Blättern: exclude_paths sammelt die IDs, Schluss bei more = false.
    class _FakeAPI(GoatCounterAPI):
        def __init__(self):
            self.anfragen = []

        def _get(self, pfad, params=None):
            self.anfragen.append(dict(params or {}))
            seite = len(self.anfragen)
            if seite == 1:
                return {"hits": [{"path_id": 1}, {"path_id": 2}], "more": True}
            if seite == 2:
                return {"hits": [{"path_id": 3}], "more": True}
            return {"hits": [{"path_id": 4}], "more": False}

    fake = _FakeAPI()
    alle = fake.hits_alle("a", "b")
    pruefe("blaettern anzahl", [h["path_id"] for h in alle], [1, 2, 3, 4])
    pruefe("blaettern exclude", [a.get("exclude_paths") for a in fake.anfragen],
           [None, "1,2", "1,2,3"])
    fake2 = _FakeAPI()
    pruefe("blaettern grenze", len(fake2.hits_alle("a", "b", max_seiten=2)), 3)

    # Ganze Durchläufe ohne Fenster: eine App ohne __init__ und eine API,
    # die nach Zeitraum antwortet ('2020…' = seit Anbeginn).
    class _ZeitAPI:
        def __init__(self, anbeginn, heute, fehler=False):
            self.anbeginn, self.heute, self.fehler = anbeginn, heute, fehler
            self.abrufe = 0

        def _liste(self, start):
            self.abrufe += 1
            if self.fehler:
                raise requests.ConnectionError("weg")
            return self.anbeginn if start == ALLTIME_START else self.heute

        def hits_alle(self, start, end):
            return list(self._liste(start))

        def hits(self, start, end, limit=100):
            return {"hits": self._liste(start)}

        def total(self, start, end):
            liste = self._liste(start)
            return {"total": sum(h["count"] for h in liste),
                    "total_events": sum(h["count"] for h in liste if h["event"])}

    def _app(api, art="auto", zeitraum="heute"):
        a = TickerApp.__new__(TickerApp)
        a.api, a.zaehlart, a.zeitraum, a.zusatz, a.max_balken = api, art, zeitraum, "aus", 8
        a._alltime_wert, a._alltime_zeit = None, 0.0
        a._alltime_hits, a._alltime_hits_zeit = None, 0.0
        a._zugang_gen, a._auto_klicks = 0, False
        a._refresh_lock, a.ui_queue = threading.Lock(), queue.Queue()
        return a

    alt = [{"path": "/", "count": 50, "event": False, "title": ""},
           {"path": "kachel-ram", "count": 4, "event": True, "title": "Real_RAM_cooler"}]
    neu = [{"path": "kachel-ram-news", "count": 2, "event": True, "title": "Real_RAM_cooler"},
           {"path": "kachel-wifi", "count": 1, "event": True, "title": "WiFi QR Card"}]
    a1 = _app(_ZeitAPI(alt + neu, neu))
    pruefe("auto kacheln", a1._zaehlart_wirksam("x"), "klicks")
    n = a1.api.abrufe
    a1._zaehlart_wirksam("x")
    pruefe("auto zwischenspeicher", a1.api.abrufe, n)
    a2 = _app(_ZeitAPI(alt[:1], alt[:1]))
    pruefe("auto ohne kacheln", a2._zaehlart_wirksam("x"), "seiten")
    a3 = _app(_ZeitAPI([], [], fehler=True))
    pruefe("auto netzfehler", a3._zaehlart_wirksam("x"), "seiten")
    a1._alltime_hits_zeit -= ALLTIME_CACHE_SECONDS + 1    # Liste veraltet
    a1.api.fehler = True
    pruefe("auto merkt kacheln", a1._zaehlart_wirksam("x"), "klicks")
    a1._zugang_neu()
    pruefe("zugang vergisst", (a1._alltime_hits, a1._auto_klicks, a1._zugang_gen),
           (None, False, 1))

    a4 = _app(_ZeitAPI(alt + neu, neu), zeitraum="gesamt")
    pruefe("durchlauf klicks ok", a4._refresh_once(), (True, None))
    d = a4.ui_queue.get_nowait()[1]
    pruefe("durchlauf klicks", (d["art"], d["zeitraum"], d["heute"], d["zeitraum_total"],
                                d["gesamt"], [b["label"] for b in d["balken"]]),
           ("klicks", "gesamt", 3, 7, 7, ["Real_RAM_cooler", "WiFi QR Card"]))
    a5 = _app(_ZeitAPI(alt + neu, neu), art="seiten", zeitraum="7t")
    a5._refresh_once()
    d = a5.ui_queue.get_nowait()[1]
    pruefe("durchlauf seiten", (d["art"], d["heute"], d["gesamt"], d["balken"]),
           ("seiten", 0, 50, []))

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
    for eingabe in ("deinname", "deinname.goatcounter.com",
                    "https://deinname.goatcounter.com/",
                    "http://deinname.goatcounter.com"):
        pruefe(f"site {eingabe}", normalisiere_site(eingabe),
               "https://deinname.goatcounter.com")

    pruefe("zahl", zahl_kurz(1234567), "1.234.567")

    # Jede Sprache kennt jeden Text, und die Anzeige übersetzt die
    # internen Beschriftungen.
    for code, tabelle in TEXTE.items():
        pruefe(f"texte {code}", sorted(set(TEXTE["de"]) - set(tabelle)), [])
    pruefe("anzeige direkt", anzeige("(direkt)"), "(direkt)")
    pruefe("anzeige start", anzeige("Startseite"), "Startseite")
    pruefe("refresh standard", refresh_name(60), "60 s  (Standard)")
    pruefe("refresh minuten", refresh_name(120), "2 min")

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
        _cfg = load_config()
        SPRACHE = waehle_sprache(_cfg)
        Hinweis(_r, "Besucher-Ticker", tx("laeuftBereits"),
                THEMES.get(_cfg.get("theme"), THEMES[DEFAULT_THEME]))
        _r.destroy()
        sys.exit(0)
    try:
        TickerApp().run()
    except KeyboardInterrupt:
        pass
    finally:
        release_single_instance()
