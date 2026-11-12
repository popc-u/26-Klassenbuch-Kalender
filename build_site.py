#!/usr/bin/env python3
"""Baut die GitHub-Pages-Seite mit den Kalender-Abos.

  site/<geheimer-token>/klassenbuch-apple.ics
  site/<geheimer-token>/klassenbuch-microsoft.ics

Der Token macht die Adresse unerratbar. Er kommt aus dem Secret FEED_TOKEN oder wird
aus dem SharePoint-Link abgeleitet (gleicher Link -> gleiche Adresse). Er wird im
Log maskiert, weil Action-Logs öffentlicher Repositories für alle sichtbar sind.

Vergleicht die neuen Dateien mit der veröffentlichten Version und meldet über
GITHUB_OUTPUT, ob sich etwas geändert hat (sonst wird nicht neu veröffentlicht).

Aufruf:  python build_site.py klassenbuch.xlsx site
Umgebung: SHAREPOINT_URL, FEED_TOKEN (optional), PAGES_URL, FORCE (true/false)
"""
from __future__ import annotations

import hashlib
import os
import shutil
import sys
import urllib.parse
import urllib.request
from pathlib import Path

import klassenbuch_ics

FILES = ("klassenbuch-apple.ics", "klassenbuch-microsoft.ics")


def feed_token() -> str:
    token = os.environ.get("FEED_TOKEN", "").strip()
    if token:
        return token
    link = os.environ.get("SHAREPOINT_URL", "").strip()
    if not link:
        raise SystemExit("SHAREPOINT_URL fehlt.")
    share_id = [x for x in urllib.parse.urlsplit(link).path.split("/") if x][-1]
    return hashlib.sha256(f"klassenbuch-kalender:{share_id}".encode()).hexdigest()[:24]


def comparable(ics: bytes) -> bytes:
    return b"\n".join(l for l in ics.splitlines() if not l.startswith(b"DTSTAMP:"))


def main() -> int:
    xlsx, site = sys.argv[1], Path(sys.argv[2])
    token = feed_token()
    print(f"::add-mask::{token}", flush=True)

    if site.exists():
        shutil.rmtree(site)
    feed_dir = site / token
    rc = klassenbuch_ics.main([xlsx, "--out-dir", str(feed_dir)])
    if rc != 0:
        return rc
    static = Path(__file__).with_name("site_static")
    if static.is_dir():
        shutil.copytree(static, site, dirs_exist_ok=True)

    changed = os.environ.get("FORCE", "").lower() == "true"
    base = os.environ.get("PAGES_URL", "").rstrip("/")
    if not changed and base:
        for name in FILES:
            try:
                req = urllib.request.Request(f"{base}/{token}/{name}", headers={"Cache-Control": "no-cache"})
                with urllib.request.urlopen(req, timeout=30) as r:
                    live = r.read()
            except Exception:
                changed = True
                break
            if comparable(live) != comparable((feed_dir / name).read_bytes()):
                changed = True
                break
    elif not base:
        changed = True

    print("Änderungen gegenüber dem veröffentlichten Kalender:", "ja" if changed else "nein")
    out = os.environ.get("GITHUB_OUTPUT")
    if out:
        with open(out, "a", encoding="utf-8") as f:
            f.write(f"changed={'true' if changed else 'false'}\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
