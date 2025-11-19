## Beschreibung

Dieser MCP verbindet Casabot/OpenWebUI mit **Microsoft 365 / Outlook** und stellt Tools für **E‑Mail** und **Kalender** bereit.  
Er kann sowohl zum **Nachschlagen** (z.B. “was habe ich heute im Kalender?”) als auch zum **aktiven Bearbeiten** (z.B. Mails verschicken, Meetings erstellen, Einladungen beantworten) verwendet werden.

---

## Mail

Nachfolgend eine Übersicht der wichtigsten E‑Mail‑Funktionen mit Beispiel‑Prompts, wie ein Nutzer sie formulieren könnte.

- **Neueste E‑Mails im Posteingang ansehen**  
  Beispiel: kannst du mir die neuesten E-Mails in meinem Posteingang anzeigen?

- **E‑Mails von einer bestimmten Person finden**  
  Beispiel: liste meine aktuellen E-Mails von alice@example.com auf.

- **E‑Mails nach Betreff‑Stichwörtern durchsuchen**  
  Beispiel: durchsuche meine E-Mails nach Nachrichten, deren Betreff den Begriff quarterly report enthält.

- **E‑Mails innerhalb eines Zeitraums auflisten**  
  Beispiel: zeig mir die E-Mails, die ich gestern zwischen 9 Uhr und 17 Uhr erhalten habe.

- **Alle E‑Mail‑Ordner anzeigen**  
  Beispiel: liste alle E-Mail-Ordner in meinem Postfach mit Namen und IDs auf.

- **Spezielle Ordner wie Archiv oder Gesendet finden**  
  Beispiel: welche Ordner habe ich, und wie lauten die IDs meiner Archiv- und Gesendete-Elemente-Ordner?

- **Nachrichten in einem bestimmten Ordner auflisten**  
  Beispiel: liste die 20 neuesten E-Mails in meinem Archiv-Ordner auf.

- **Nachrichten in einem benutzerdefinierten Ordner per ID ansehen**  
  Beispiel: nutze diese Ordner-ID und liste die neuesten E-Mails in diesem Ordner auf: FOLDER_ID_HERE.

- **Spezifische E‑Mail per ID in der Vorschau anzeigen**  
  Beispiel: zeig mir eine sichere Textvorschau dieser E-Mail: MESSAGE_ID_HERE.

- **Kompletten Inhalt einer E‑Mail lesen (HTML oder Text)**  
  Beispiel: hole den vollständigen HTML-Inhalt für diese Nachrichten-ID, damit ich die komplette E-Mail sehen kann: MESSAGE_ID_HERE.

- **Lange E‑Mail zusammenfassen**  
  Beispiel: hole diese Nachricht per ID und gib mir eine kurze Zusammenfassung der wichtigsten Punkte: MESSAGE_ID_HERE.

- **Einen einfachen E‑Mail‑Entwurf erstellen**  
  Beispiel: erstelle einen E-Mail-Entwurf an bob@example.com mit dem Betreff „Status-Update“ und einem kurzen Text, der meinen Fortschritt dieser Woche zusammenfasst.

- **Formatierten (HTML-)E‑Mail‑Entwurf erstellen**  
  Beispiel: erstelle einen HTML-E-Mail-Entwurf an team@example.com, der unsere Sprint-Ergebnisse mit Überschriften und Aufzählungspunkten zusammenfasst.

- **Einen E‑Mail‑Entwurf an mehrere Empfänger erstellen**  
  Beispiel: erstelle einen E-Mail-Entwurf an alice@example.com und bob@example.com mit dem Betreff „Planungsmeeting“ und einer kurzen Nachricht mit einem Terminvorschlag.

- **E‑Mail per ID löschen**  
  Beispiel: lösche diese E-Mail aus meinem Postfach anhand ihrer Nachrichten-ID: MESSAGE_ID_HERE.

- **E‑Mail ins Archiv verschieben**  
  Beispiel: verschiebe diese E-Mail (MESSAGE_ID_HERE) in meinen Archiv-Ordner.

- **E‑Mail per ID in einen bestimmten Ordner verschieben**  
  Beispiel: verschiebe die Nachricht mit dieser ID MESSAGE_ID_HERE in den Ordner mit der ID FOLDER_ID_HERE.

- **Datei an vorhandenen Entwurf anhängen**  
  Beispiel: hänge die lokale Datei report.pdf aus meinem Workspace an den E-Mail-Entwurf mit der ID DRAFT_ID_HERE an.

- **Alle Anhänge einer Nachricht auflisten**  
  Beispiel: liste alle Anhänge der E-Mail mit der ID MESSAGE_ID_HERE mit Namen und Größen auf.

- **Bestimmten Anhang herunterladen**  
  Beispiel: lade den Anhang mit dem Namen invoice.pdf aus der E-Mail mit der ID MESSAGE_ID_HERE herunter.

- **Anhang aus einem Entwurf löschen**  
  Beispiel: entferne den Anhang mit dem Namen draft-notes.docx aus dem E-Mail-Entwurf mit der ID DRAFT_ID_HERE.

- **Neue E‑Mail direkt senden (ohne gespeicherten Entwurf)**  
  Beispiel: sende eine E-Mail an alice@example.com mit dem Betreff „Kurze Frage“ und einem kurzen Text, in dem ich nach der Projektdeadline frage.

- **Vorbereiteten Entwurf versenden**  
  Beispiel: sende den vorhandenen E-Mail-Entwurf mit der ID DRAFT_ID_HERE.

- **E‑Mail beantworten**  
  Beispiel: beantworte die E-Mail mit dem Betreff „Anfrage Angebot“ und bedanke dich kurz für die Anfrage.

- **E‑Mail weiterleiten**  
  Beispiel: leite die E-Mail „Rechnung März“ an buchhaltung@testfirma.de weiter.

- **Neueste E‑Mails aus einem freigegebenen Postfach lesen**  
  Beispiel: zeig mir die 20 neuesten E-Mails im freigegebenen Postfach support@example.com.

- **Nachrichten in einem freigegebenen Postfach nach Absender durchsuchen**  
  Beispiel: liste aktuelle E-Mails von customer@example.com im freigegebenen Postfach support@example.com auf.

- **Nachrichten in einem bestimmten Ordner eines freigegebenen Postfachs auflisten**  
  Beispiel: liste die neuesten E-Mails im Ordner „Open Tickets“ des freigegebenen Postfachs support@example.com auf.

- **Spezifische Nachricht aus freigegebenem Postfach per ID in der Vorschau anzeigen**  
  Beispiel: zeig mir eine Textvorschau der Nachricht mit dieser ID aus dem freigegebenen Postfach support@example.com: MESSAGE_ID_HERE.

- **Von einer freigegebenen Postfach‑Adresse antworten**  
  Beispiel: sende eine E-Mail vom freigegebenen Postfach support@example.com an customer@example.com und bestätige, dass wir seine Anfrage erhalten haben.

---

## Kalender

Hier die Kalender‑Funktionen mit passenden Beispiel‑Prompts.

- **Termine in einem bestimmten Zeitfenster prüfen**  
  Beispiele:
  - welche Besprechungen und Termine habe ich heute in meinem Kalender?
  - welche Besprechungen sind für morgen für mich eingeplant?
  - liste alle Kalendereinträge für den Rest dieser Woche auf.
  - zeig mir alle Besprechungen, die ich für nächste Woche geplant habe.
  - liste meine Kalendereinträge zwischen dem 10. März und dem 15. März auf.
  - kannst du bitte auflisten, welche Termine ich heute zwischen 13 Uhr und 17 Uhr habe?

- **Details und Zusammenfassung eines bestimmten Termins abrufen**  
  Beispiele:
  - zeig mir die vollständigen Details (Teilnehmer, Ort, Beschreibung) für meinen nächsten Termin heute um 17 Uhr.
  - fass mir bitte Betreff, Uhrzeit und Teilnehmer der Besprechung „Projekt-Kick-off“ zusammen.

- **Einen einfachen einmaligen Termin im Primärkalender erstellen**  
  Beispiele:
  - erstelle morgen von 15 Uhr bis 16 Uhr eine Besprechung namens „Projekt-Kick-off“ mit alice@example.com und bob@example.com.
  - füge mir morgen um 9 Uhr einen 30-minütigen „Fokuszeit“-Termin in meinem Kalender hinzu.

- **Wiederkehrende Termine im Kalender erstellen**  
  Beispiele:
  - erstelle einen wöchentlich wiederkehrenden Termin „Team-Sync“ jeden Montag von 10 Uhr bis 10:30 Uhr.
  - lege einen monatlichen Termin „Reporting“ am ersten Werktag jedes Monats von 9 Uhr bis 10 Uhr an.

- **Bestehenden Termin verschieben**  
  Beispiel: verschiebe die Besprechung "Team Meeting" von morgen 14–15 Uhr auf Freitag 10–11 Uhr.

- **Betreff oder Beschreibung eines Termins ändern**  
  Beispiel: ändere den Betreff des Termins "Team Meeting" auf „Aktualisierte Projektbesprechung“ und füge eine kurze Beschreibung der Agenda hinzu.

- **Online‑/Teams‑Besprechung erstellen**  
  Beispiele:
  - erstelle für morgen von 16 Uhr bis 17 Uhr eine Teams-Besprechung namens „Kunden-Call“ mit alice@example.com und bob@example.com.
  - plane für nächsten Mittwoch um 11 Uhr eine Online-Besprechung „1:1 mit meiner Führungskraft“ und füge den Teams-Link in den Termin ein.

- **Von dir organisierte Besprechung absagen**  
  Beispiel: lösche den Termin den ich heut um 15 uhr habe aus meinem Kalender.

- **Alle meine Kalender auflisten**  
  Beispiel: liste alle Kalender in meinem Konto inklusive ihrer Namen und IDs auf.

- **Detaillierte Tagesagenda abrufen**  
  Beispiel: gib mir eine chronologische Agenda aller Besprechungen für nächsten Dienstag zwischen 8 Uhr und 18 Uhr.

- **Frühestmöglichen freien Termin‑Slot finden**  
  Beispiel: finde den frühesten 30-minütigen Zeitraum in den nächsten drei Tagen, in dem ich und alice@example.com zwischen 9 Uhr und 18 Uhr beide frei sind.

- **Verfügbarkeit von Personen prüfen**  
  Beispiele:
  - prüfe, wann ich und alice@example.com morgen zwischen 9 Uhr und 18 Uhr gleichzeitig verfügbar sind.
  - zeig mir die freien Zeitfenster von bob@example.com am Freitag zwischen 10 Uhr und 16 Uhr.
  - prüfe, ob ich nächsten Dienstag zwischen 14 Uhr und 15 Uhr noch frei bin.

- **Mögliche Besprechungszeiten mit mehreren Personen finden**  
  Beispiel: schlage mir ein paar mögliche 60-minütige Besprechungszeiten für nächste Woche vor, zu denen ich sowie alice@example.com und bob@example.com zwischen 9 Uhr und 17 Uhr verfügbar sind.

- **Auf Kalendereinladungen reagieren**  
  Beispiele:
  - akzeptiere die Einladung zur Besprechung „Projekt-Kick-off“ heute um 15 Uhr.
  - lehne die Einladung zur Besprechung „Marketing-Update“ morgen um 9 Uhr ab.
  - antworte auf die Einladung „Team-Lunch“ mit einer vorläufigen Zusage.

- **Termine in Nicht‑Primärkalendern bearbeiten**  
  Beispiele:
  - liste alle Termine für die nächsten 7 Tage im Kalender „Geburtstage“ auf.
  - zeig mir die vollständigen Details für den Termin „Kickoff Meeting“ im Kalender „Team Kalender“.
  - erstelle nächsten Mittwoch von 12 Uhr bis 13 Uhr eine Besprechung namens „Team-Lunch“ mit alice@example.com im Kalender „Team Kalender“.
  - verschiebe im Kalender „Team Kalender“ das „Team Meeting“ um eine Stunde nach hinten, behalte aber die gleiche Dauer bei.
  - lösche den Termin „Team Lunch“ aus dem Kalender „Team Kalender“.

---

## In Arbeit: Noch nicht verfügbare Funktionen

Diese Funktionen sind bereits geplant, funktionieren aber derzeit noch nicht zuverlässig im MCP‑Server.

### E‑Mail

- **Nachrichten in einem bestimmten Ordner auflisten**  
  Beispiele: 
  - zeig mir die 20 neuesten E-Mails im Ordner „Archiv“.
  - welche mails habe ich in meinem "Wichtig" Ordner?

- **E-Mail per ID in einen bestimmten Ordner verschieben**  
  Beispiel: verschiebe die E-Mail „Reisekosten Januar“ in den Ordner „Buchhaltung“.

- **Datei an vorhandenen Entwurf anhängen**  
  Beispiel: hänge die Datei „report.pdf“ an den E-Mail-Entwurf mit dem Betreff „Quartalsbericht“ an.

- **Bestimmten Anhang herunterladen**  
  Beispiel: lade den Anhang „invoice.pdf“ aus der E-Mail „Rechnung März“ herunter.

- **Anhang aus einem Entwurf löschen**  
  Beispiel: entferne den Anhang „draft-notes.docx“ aus meinem E-Mail-Entwurf „Projektplan“.

- **E-Mail-Entwurf löschen**  
  Beispiel: lösch den E-Mail-Entwurf mit dem Betreff „Testentwurf“.

- **Neueste E-Mails aus einem freigegebenen Postfach lesen**  
  Beispiel: zeig mir die 20 neuesten E-Mails im freigegebenen Postfach support@testfirma.de.

- **Nachrichten in einem freigegebenen Postfach nach Absender durchsuchen**  
  Beispiel: such alle aktuellen E-Mails von kunde@test.de im freigegebenen Postfach support@testfirma.de.

- **Nachrichten in einem bestimmten Ordner eines freigegebenen Postfachs auflisten**  
  Beispiel: liste die neuesten E-Mails im Ordner „Open Tickets“ des freigegebenen Postfachs support@testfirma.de auf.

- **Spezifische Nachricht aus freigegebenem Postfach per ID in der Vorschau anzeigen**  
  Beispiel: zeig mir eine Textvorschau der E-Mail mit der ID „98765“ aus dem freigegebenen Postfach support@testfirma.de.

- **Von einer freigegebenen Postfach-Adresse antworten**  
  Beispiel: antworte aus dem Postfach support@testfirma.de dem Absender der letzten E-Mail mit einer kurzen Bestätigung, dass wir uns um sein Anliegen kümmern.

### Kalender

- **Frühestmöglichen freien Termin-Slot finden**  
  Beispiel: finde den frühesten 30-minütigen Zeitraum in den nächsten drei Tagen, in dem ich und alice@example.com zwischen 9 Uhr und 18 Uhr beide frei sind.

- **Verfügbarkeit von Personen prüfen**  
  Beispiele:
  - prüfe, wann ich und alice@example.com morgen zwischen 9 Uhr und 18 Uhr gleichzeitig verfügbar sind.
  - zeig mir die freien Zeitfenster von bob@example.com am Freitag zwischen 10 Uhr und 16 Uhr.
  - prüfe, ob ich nächsten Dienstag zwischen 14 Uhr und 15 Uhr noch frei bin.

- **Mögliche Besprechungszeiten mit mehreren Personen finden**  
  Beispiel: schlage mir ein paar mögliche 60-minütige Besprechungszeiten für nächste Woche vor, zu denen ich sowie alice@example.com und bob@example.com zwischen 9 Uhr und 17 Uhr verfügbar sind.

- **Auf Kalendereinladungen reagieren**  
  Beispiele:
  - akzeptiere die Einladung zur Besprechung „Projekt-Kick-off“ heute um 15 Uhr.
  - lehne die Einladung zur Besprechung „Marketing-Update“ morgen um 9 Uhr ab.
  - antworte auf die Einladung „Team-Lunch“ mit einer vorläufigen Zusage.

- **Termine in Nicht-Primärkalendern bearbeiten**  
  Beispiele:
  - liste alle Termine für die nächsten 7 Tage im Kalender „Geburtstage“ auf.
  - zeig mir die vollständigen Details für den Termin „Kickoff Meeting“ im Kalender „Team Kalender“.
  - erstelle nächsten Mittwoch von 12 Uhr bis 13 Uhr eine Besprechung namens „Team-Lunch“ mit alice@example.com im Kalender „Team Kalender“.
  - verschiebe im Kalender „Team Kalender“ das „Team Meeting“ um eine Stunde nach hinten, behalte aber die gleiche Dauer bei.
  - lösche den Termin „Team Lunch“ aus dem Kalender „Team Kalender“.


