# Besucher-Ticker 📊

**Your visitor numbers in view while you work.** A small floating readout for Windows that shows the numbers of your own [GoatCounter](https://www.goatcounter.com) site live: pageviews for the chosen period, visitors per page as bars, plus countries or referrers. Black, pink text, blue bars.

![Windows](https://img.shields.io/badge/Windows-10%20%7C%2011-818cf8) ![Requires](https://img.shields.io/badge/Requires-your%20own%20GoatCounter%20account-4338ca) ![Language](https://img.shields.io/badge/Language-German-6366f1) ![License](https://img.shields.io/badge/License-MIT-a5b4fc)

[Deutsche Fassung](README.de.md) — the app itself speaks German.

---

## What for

The GoatCounter dashboard is for sitting down and looking at. This readout is for running alongside: it sits in a corner of your screen and answers, without you opening any page — *how many people came, and where were they?*

## Requirement

A website of your own that counts with GoatCounter. The ticker does not count anything itself — it displays what your GoatCounter account already knows. Without an account it consequently shows nothing at all.

## What it shows

| | |
|---|---|
| **Pageviews** | the big number: pageviews for the chosen period (today / 7 days / 30 days / all time), with today's and the all-time value beside it |
| **Bars** | visitors per page — subpages are folded into their page instead of appearing separately |
| **Extra block** | countries or referrers, or off |

**The bars do not add up to the big number — on purpose.** GoatCounter tracks two different quantities: *pageviews* (how often something was opened) and *visitors* (how many people came). The ticker does not mix them, and the labels keep them apart.

Click events (GoatCounter's `data-count` events) are filtered out of the bars. Otherwise a single click would count twice: once as a pageview and once as an event.

## Getting started

**With the prebuilt exe** — download it from the [releases](../../releases/latest), put it anywhere, double-click. On first start Windows will say "unknown publisher": the file is not signed, which costs money. Continue via **More info → Run anyway**.

**From source:**

```
pip install requests pystray pillow
pip install pywin32        (optional, encrypts the token)
pythonw besucher_ticker.py
```

**On first start** the ticker asks for two things: the address of your GoatCounter site (`yourname` is enough) and an API token. You create the token on your GoatCounter site: username in the top right → **API** → **Add new token**. Ticking **Read statistics** is enough — the ticker needs no more rights, and fewer rights can do less harm should the token ever get out.

## Controls

| | |
|---|---|
| Drag | move the overlay — the position is saved when you let go |
| Double-click | switch between small and large |
| Right-click the tray icon | overlay on/off, bring to centre, refresh now, period, extra block, settings, themes, change access, quit |

Overlay gone? **Bring to centre** in the menu. That happens when a second monitor was unplugged — the overlay is borderless and has no title bar to drag it back by.

## Privacy

**The ticker talks to exactly one address: your own GoatCounter site.** Read-only requests, no telemetry, no update check, no identifier.

The token lives in `config.json` next to the program — encrypted via Windows DPAPI when `pywin32` is installed, bound to your user account and your machine. Without `pywin32` it sits there in plain text; the program still runs, but you should know. Either way: the file stays local and does not belong in any backup others can read.

## Under the hood

One file, around 900 lines, three third-party packages (`requests`, `pystray`, `pillow`; `pywin32` optional) — everything else is the Python standard library.

- One refresh takes 3–4 requests, one after another. GoatCounter allows 4 per second — even the fastest setting (15 s) stays far below that, and there are deliberately no per-second settings.
- The all-time total is fetched at most every 10 minutes: it changes slowly, and every query costs a request.
- On HTTP 429 the ticker waits as long as the server asks — GoatCounter reports that in the `X-Rate-Limit-Reset` header, with `Retry-After` as the fallback.
- `python besucher_ticker.py --selbsttest` checks the arithmetic and mapping logic without network and without a window.

## License

MIT — see [LICENSE](LICENSE).

---

Besucher-Ticker is a personal tool and is **not affiliated with GoatCounter**. GoatCounter is a project by Martin Tournoij.

Part of the [workshop](https://dennismit2n.github.io/) by Dennis_mit_2n.
