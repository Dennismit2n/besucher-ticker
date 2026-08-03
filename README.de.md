# Besucher-Ticker 📊

**Deine Besucherzahlen im Blick, während du arbeitest.** Eine kleine schwebende Anzeige für Windows, die die Zahlen der eigenen [GoatCounter](https://www.goatcounter.com)-Seite live zeigt: Aufrufe im gewählten Zeitraum, Besucher je Seite als Balken, dazu Länder oder Verweise. Schwarz, pinke Schrift, blaue Balken.

![Windows](https://img.shields.io/badge/Windows-10%20%7C%2011-818cf8) ![Voraussetzung](https://img.shields.io/badge/Voraussetzung-eigenes%20GoatCounter--Konto-4338ca) ![Sprache](https://img.shields.io/badge/Sprache-Deutsch-6366f1) ![Lizenz](https://img.shields.io/badge/Lizenz-MIT-a5b4fc)

[English version](README.md)

---

## Wofür

Das GoatCounter-Dashboard ist zum Hinsetzen und Anschauen. Diese Anzeige hier ist zum Nebenherlaufen: Sie steht in einer Bildschirmecke und beantwortet, ohne dass du eine Seite öffnest — *wie viele waren da, und wo waren sie?*

## Voraussetzung

Eine eigene Webseite, die mit GoatCounter zählt. Der Ticker zählt selbst nichts — er zeigt an, was dein GoatCounter-Konto schon weiß. Ohne Konto zeigt er folgerichtig gar nichts.

## Was angezeigt wird

| | |
|---|---|
| **Aufrufe** | die große Zahl: Seitenaufrufe im gewählten Zeitraum (heute / 7 Tage / 30 Tage / gesamt), daneben der Heute- und der Gesamtwert |
| **Balken** | Besucher je Seite — Unterseiten werden ihrer Seite zugeschlagen statt einzeln aufzutauchen |
| **Zusatzblock** | wahlweise Länder oder Verweise, oder aus |

**Die Summe der Balken ergibt nicht die große Zahl — mit Absicht.** GoatCounter kennt zwei verschiedene Größen: *Aufrufe* (wie oft wurde etwas geöffnet) und *Besucher* (wie viele Leute waren da). Der Ticker vermischt sie nicht, die Beschriftung hält sie auseinander.

Klick-Ereignisse (GoatCounters `data-count`-Events) werden aus den Balken herausgefiltert. Sonst zählte ein einziger Klick doppelt: einmal als Seitenaufruf, einmal als Ereignis.

## Loslegen

**Mit der fertigen Exe** — aus den [Releases](../../releases/latest) herunterladen, irgendwo hinlegen, doppelklicken. Beim ersten Start meldet Windows „Unbekannter Herausgeber": Die Datei ist nicht signiert, das kostet Geld. Über **Weitere Informationen → Trotzdem ausführen**.

**Aus dem Quellcode:**

```
pip install requests pystray pillow
pip install pywin32        (optional, verschlüsselt das Token)
pythonw besucher_ticker.py
```

**Beim ersten Start** fragt der Ticker zwei Dinge: die Adresse deiner GoatCounter-Seite (`deinname` genügt) und ein API-Token. Das Token legst du auf deiner GoatCounter-Seite an: Benutzername oben rechts → **API** → **Add new token**. Das Häkchen bei **Read statistics** genügt — mehr Rechte braucht der Ticker nicht, und weniger Rechte können weniger Schaden anrichten, falls das Token je abhandenkommt.

## Bedienung

| | |
|---|---|
| Ziehen | Overlay verschieben — Position wird beim Loslassen gemerkt |
| Doppelklick | zwischen klein und groß umschalten |
| Rechtsklick aufs Taskleisten-Symbol | Overlay an/aus, zur Mitte holen, jetzt aktualisieren, Zeitraum, Zusatzblock, Einstellungen, Themes, Zugang ändern, Beenden |

Verschwunden? **Zur Mitte holen** im Menü. Das passiert, wenn ein zweiter Bildschirm abgezogen wurde — das Overlay ist randlos und hat keine Titelleiste zum Zurückholen.

## Datenschutz

**Der Ticker spricht mit genau einer Adresse: deiner eigenen GoatCounter-Seite.** Nur Lesezugriffe, keine Telemetrie, kein Update-Check, keine Kennung.

Das Token liegt in der `config.json` neben dem Programm — mit `pywin32` per Windows-DPAPI verschlüsselt, gebunden an dein Benutzerkonto und deinen Rechner. Ohne `pywin32` liegt es im Klartext da; das Programm läuft trotzdem, aber du solltest es wissen. In beiden Fällen gilt: Die Datei bleibt lokal und gehört in kein Backup, das andere lesen können.

## Unter der Haube

Eine Datei, rund 900 Zeilen, drei Fremdpakete (`requests`, `pystray`, `pillow`; `pywin32` optional) — alles andere ist Python-Standardbibliothek.

- Ein Durchlauf braucht 3–4 Anfragen, nacheinander. GoatCounter erlaubt 4 je Sekunde — selbst die schnellste Stufe (15 s) bleibt weit darunter, und Sekundenstufen gibt es bewusst nicht.
- Die Gesamtsumme wird höchstens alle 10 Minuten frisch geholt: Sie ändert sich träge, und jede Abfrage kostet eine Anfrage.
- Bei HTTP 429 wartet der Ticker, so lange der Server es verlangt — GoatCounter meldet das im Kopffeld `X-Rate-Limit-Reset`, ersatzweise `Retry-After`.
- `python besucher_ticker.py --selbsttest` prüft die Rechen- und Zuordnungslogik ohne Netz und ohne Fenster.

## Lizenz

MIT — siehe [LICENSE](LICENSE).

---

Besucher-Ticker ist ein privates Werkzeug und steht **in keiner Verbindung zu GoatCounter**. GoatCounter ist ein Projekt von Martin Tournoij.

Teil der [Werkstatt](https://dennismit2n.github.io/) von Dennis_mit_2n.
