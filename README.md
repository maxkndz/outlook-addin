# SR Standards – Outlook Add-In

Dieses Paket stellt für das **neue Outlook** ein einfaches Add-In bereit.

## Funktion
In einer neuen E-Mail erscheint ein Button **„SR Standards“**.
Beim Klick öffnet sich rechts ein Panel mit drei Vorlagen:

- Statusbericht
- Protokollversand
- Terminbestätigung

Beim Klick auf eine Vorlage wird:
- optional ein Standard-Betreff gesetzt, falls der Betreff leer ist
- der Text oben in die Mail eingefügt

---

## Dateistruktur

```text
manifest.xml
web/
  commands.html
  commands.js
  taskpane.html
  taskpane.js
  icons/
    icon-16.png
    icon-32.png
    icon-80.png