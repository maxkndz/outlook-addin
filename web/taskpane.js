/* global Office */

Office.onReady(() => {
  const btnRechnung = document.getElementById("btnRechnung");
  const btnNachtrag = document.getElementById("btnNachtrag");

  if (btnRechnung) {
    btnRechnung.addEventListener("click", () => {
      insertTemplate([
        "<p>Sehr geehrte Damen und Herren,</p>",
        "<p>im Zuge der Plausibilitätsprüfung der eingereichten Rechnung ergeben sich Unstimmigkeiten, die eine Überprüfung bzw. Freigabe derzeit nicht ermöglichen.</p>",
        "<p>Im Einzelnen ergeben sich folgende Punkte:</p>",
        "<ul>",
        "<li>Fehlende Nachweise / Anlagen</li>",
        "<li>Nicht nachvollziehbare Nachweise / Anlagen</li>",
        "<li>Abweichungen von den vereinbarten Einheitspreisen gemäß Leistungsverzeichnis / Auftrag</li>",
        "<li>Abrechnung nicht beauftragter Leistungen</li>",
        "<li>Rechnerische Fehler</li>",
        "<li>Überschreitung der Auftragssumme</li>",
        "<li>Übertragungsfehler</li>",
        "<li>Falscher Stempel</li>",
        "</ul>",
        "<p>Wir bitten Sie, die genannten Punkte zu prüfen, die Unterlage(n) entsprechend zu überarbeiten und erneut vorzulegen.</p>",
        "<p>Für Rückfragen stehen wir gerne zur Verfügung.</p>"
      ].join(""));
    });
  }

  if (btnNachtrag) {
    btnNachtrag.addEventListener("click", () => {
      insertTemplate([
        "<p>Sehr geehrte Damen und Herren,</p>",
        "<p>im Zuge der Plausibilitätsprüfung des eingereichten Nachtrags ergeben sich Unstimmigkeiten, die eine Überprüfung bzw. Freigabe derzeit nicht ermöglichen.</p>",
        "<p>Im Einzelnen ergeben sich folgende Punkte:</p>",
        "<ul>",
        "<li>Fehlende Nachweise / Anlagen</li>",
        "<li>Nicht nachvollziehbare Nachweise / Anlagen</li>",
        "<li>Abweichungen von den vereinbarten Einheitspreisen</li>",
        "<li>Unangemessene Preisansätze</li>",
        "<li>Fehlende Nachtragsbegründung / Unzureichende Nachtragsbegründung</li>",
        "<li>Leistungen bereits beauftragt</li>",
        "<li>Nicht nachvollziehbare Mengenansätze</li>",
        "<li>Rechnerische Fehler</li>",
        "<li>Übertragungsfehler</li>",
        "<li>Falscher Stempel</li>",
        "</ul>",
        "<p>Wir bitten Sie, die genannten Punkte zu prüfen, die Unterlage(n) entsprechend zu überarbeiten und erneut vorzulegen.</p>",
        "<p>Für Rückfragen stehen wir gerne zur Verfügung.</p>"
      ].join(""));
    });
  }
});

function setStatus(message) {
  const status = document.getElementById("status");
  if (status) {
    status.textContent = message;
  }
}

function insertTemplate(htmlTemplate) {
  const item = Office.context.mailbox.item;

  if (!item || !item.body) {
    setStatus("Keine bearbeitbare E-Mail gefunden.");
    return;
  }

  item.body.prependAsync(
    htmlTemplate,
    { coercionType: Office.CoercionType.Html },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        setStatus("Text eingefügt.");
      } else {
        setStatus("Fehler beim Einfügen des Textes.");
      }
    }
  );
}