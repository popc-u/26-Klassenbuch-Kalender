#!/usr/bin/env python3
"""Lädt das Klassenbuch über den SharePoint-/OneDrive-Freigabelink herunter.

Der Link kommt aus der Umgebungsvariablen SHAREPOINT_URL (GitHub-Secret), damit er
nicht öffentlich im Repository steht.

Aufruf:  SHAREPOINT_URL="https://...sharepoint.com/:x:/g/personal/..." \
         python fetch_klassenbuch.py klassenbuch.xlsx

Exit-Codes: 0 = ok, 2 = Link fehlt, 3 = Download fehlgeschlagen / Anmeldung nötig
"""
from __future__ import annotations

import http.cookiejar
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 14_0) AppleWebKit/605.1.15 "
      "(KHTML, like Gecko) Version/17.0 Safari/605.1.15")


def candidate_urls(link: str) -> list[str]:
    """Direkt-Download-Varianten eines Freigabelinks."""
    link = link.strip()
    p = urllib.parse.urlsplit(link)
    query = urllib.parse.parse_qsl(p.query, keep_blank_values=True)
    urls = []
    if not any(k == "download" for k, _ in query):
        urls.append(urllib.parse.urlunsplit(p._replace(query=urllib.parse.urlencode(query + [("download", "1")]))))
    else:
        urls.append(link)
    # /:x:/g/personal/<user>/<token>  ->  /personal/<user>/_layouts/15/download.aspx?share=<token>
    parts = [x for x in p.path.split("/") if x]
    if len(parts) >= 4 and parts[0].startswith(":") and parts[2] == "personal":
        user, token = parts[3], parts[-1]
        urls.append(f"{p.scheme}://{p.netloc}/personal/{user}/_layouts/15/download.aspx?share={token}")
    return urls


def download(url: str) -> tuple[bytes, str, str]:
    jar = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(jar))
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "*/*"})
    with opener.open(req, timeout=60) as resp:
        return resp.read(), resp.geturl(), resp.headers.get("Content-Type", "")


def main() -> int:
    if len(sys.argv) != 2:
        print("Aufruf: python fetch_klassenbuch.py ZIELDATEI.xlsx", file=sys.stderr)
        return 2
    target = sys.argv[1]
    link = os.environ.get("SHAREPOINT_URL", "").strip()
    if not link:
        print("::error::Secret SHAREPOINT_URL fehlt (Settings → Secrets and variables → Actions).")
        return 2

    last_problem = "unbekannt"
    for attempt in range(1, 4):
        for url in candidate_urls(link):
            try:
                data, final_url, ctype = download(url)
            except urllib.error.HTTPError as e:
                last_problem = f"HTTP {e.code}"
                continue
            except Exception as e:  # Netzwerkfehler, Timeout ...
                last_problem = type(e).__name__
                continue
            host = urllib.parse.urlsplit(final_url).netloc
            if data[:4] == b"PK\x03\x04":
                with open(target, "wb") as f:
                    f.write(data)
                print(f"Klassenbuch geladen: {len(data)} Bytes.")
                return 0
            if "login" in host or "microsoftonline" in host:
                last_problem = "Anmeldung verlangt"
            else:
                last_problem = f"keine Excel-Datei erhalten ({ctype or 'unbekannter Typ'})"
        time.sleep(10 * attempt)

    if last_problem == "Anmeldung verlangt":
        print("::error::Der SharePoint-Link verlangt eine Anmeldung. Für den automatischen Abruf "
              "braucht es einen Link 'Jeder mit dem Link kann anzeigen'.")
    else:
        print(f"::error::Klassenbuch konnte nicht geladen werden: {last_problem}.")
    return 3


if __name__ == "__main__":
    sys.exit(main())
