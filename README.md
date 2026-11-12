# Klassenbuch-Kalender

Macht aus dem Klassenbuch (Excel-Datei auf SharePoint) einen Kalender, den Apple Kalender
und Microsoft Outlook abonnieren können. Jede Stunde lädt GitHub die aktuelle Excel-Datei
über den Freigabelink, erzeugt die Kalenderdateien neu und veröffentlicht sie, wenn sich
etwas geändert hat.

Im Kalender erscheinen:

- jeder Eintrag als ganztägiger Termin (mehrzeilige Zellen werden pro Zeile bzw. pro Fach aufgeteilt)
- Einträge mit Uhrzeit ("17 Uhr", "12:30") zur angegebenen Zeit
- Ferien (im Klassenbuch grün) und Sonderwochen als mehrtägige Termine
- Randnotizen mit Datumsbereich, z. B. "Besuch aus Flöha 13.3.-19.3."

## Einrichtung (einmalig)

1. **Secret**: *Settings → Secrets and variables → Actions → New repository secret*
   Name `SHAREPOINT_URL`, Wert: der Freigabelink des Klassenbuchs.
2. **Pages**: *Settings → Pages → Build and deployment → Source:* **GitHub Actions**.
3. *Actions → Klassenbuch-Kalender → Run workflow* (oder einfach eine Stunde warten).

Der Link muss ohne Anmeldung funktionieren ("Jeder mit dem Link kann anzeigen").
Verlangt er ein Schulkonto, schlägt der Download fehl – dann die Lehrperson um einen
solchen Link bitten und das Secret ersetzen.

## Kalender-Adresse

```
https://<benutzername>.github.io/klassenbuch-kalender/<token>/klassenbuch-apple.ics
https://<benutzername>.github.io/klassenbuch-kalender/<token>/klassenbuch-microsoft.ics
```

`<token>` ist eine geheime Zeichenfolge, damit niemand die Adresse erraten kann. Sie wird
aus dem SharePoint-Link abgeleitet (gleicher Link → gleiche Adresse) und in den
öffentlichen Logs als `***` ausgeblendet. Ändert sich der SharePoint-Link, ändert sich
auch die Adresse – wer die alte behalten will, legt zusätzlich das Secret `FEED_TOKEN`
mit dem alten Token an.

## Abonnieren

**Apple Kalender (Mac):** *Ablage → Neues Kalenderabonnement…* → Adresse (`webcal://…/klassenbuch-apple.ics`)
→ *Ort: iCloud* (dann auch auf iPhone/iPad) → *Automatisch aktualisieren: Jede Stunde*.

**iPhone:** *Einstellungen → Apps → Kalender → Accounts → Account hinzufügen → Andere →
Kalenderabo hinzufügen*.

**Outlook (Web, neues Outlook, Outlook.com, Schulkonto):** *Kalender → Kalender hinzufügen →
Aus dem Internet abonnieren* → `https://…/klassenbuch-microsoft.ics`.
Outlook holt abonnierte Kalender nach eigenem Zeitplan (meist alle paar Stunden); das lässt
sich dort nicht auf stündlich stellen.

Nicht zusätzlich eine .ics-Datei importieren – sonst erscheinen die Termine doppelt.

## Lokal ausprobieren

```bash
pip install -r requirements.txt
python klassenbuch_ics.py Klassenbuch.xlsx --out-dir ausgabe --list
```

## Anhalten

*Actions → Klassenbuch-Kalender → … → Disable workflow* oder das Repository löschen.
