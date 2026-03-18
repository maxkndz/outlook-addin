/* global Office */

Office.onReady(() => {
  const btnStatusbericht = document.getElementById("btnStatusbericht");
  const btnProtokoll = document.getElementById("btnProtokoll");
  const btnTermin = document.getElementById("btnTermin");

  if (btnStatusbericht) {
    btnStatusbericht.addEventListener("click", () => {
      insertTemplate(
        "Statusbericht – [Projektname]",
        [
          "<p>Sehr geehrte Damen und Herren,</p>",
          "<p>anbei übersenden wir Ihnen den aktuellen Statusbericht zum Projekt <strong>[Projektname]</strong>.</p>",
          "<p>Für Rückfragen stehen wir gerne zur Verfügung.</p>",
          "<p>Mit freundlichen Grüßen<br>[Name]</p>"
        ].join("")
      );
    });
  }

  if (btnProtokoll) {
    btnProtokoll.addEventListener("click", () => {
      insertTemplate(
        "Protokollversand – [Projektname]",
        [
          "<p>Sehr geehrte Damen und Herren,</p>",
          "<p>anbei erhalten Sie das Protokoll vom <strong>[Datum]</strong> zur weiteren Verwendung.</p>",
          "<p>Bitte prüfen Sie die Inhalte und geben Sie uns bei Bedarf Rückmeldung.</p>",
          "<p>Mit freundlichen Grüßen<br>[Name]</p>"
        ].join("")
      );
    });
  }

  if (btnTermin) {
    btnTermin.addEventListener("click", () => {
      insertTemplate(
        "Terminbestätigung – [Projektname]",
        [
          "<p>Sehr geehrte Damen und Herren,</p>",
          "<p>hiermit bestätigen wir den Termin am <strong>[Datum]</strong> um <strong>[Uhrzeit]</strong>.</p>",
          "<p>Ort: <strong>[Ort]</strong></p>",
          "<p>Mit freundlichen Grüßen<br>[Name]</p>"
        ].join("")
      );
    });
  }
});

function setStatus(message) {
  const status = document.getElementById("status");
  if (status) {
    status.textContent = message;
  }
}

function insertTemplate(subjectText, htmlTemplate) {
  const item = Office.context.mailbox.item;

  if (!item || !item.body || !item.subject) {
    setStatus("Keine bearbeitbare E-Mail gefunden.");
    return;
  }

  item.subject.getAsync((subjectResult) => {
    const currentSubject =
      subjectResult.status === Office.AsyncResultStatus.Succeeded
        ? (subjectResult.value || "")
        : "";

    const continueWithBody = () => {
      item.body.prependAsync(
        htmlTemplate,
        { coercionType: Office.CoercionType.Html },
        (bodyResult) => {
          if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
            setStatus("Text eingefügt.");
          } else {
            setStatus("Fehler beim Einfügen des Textes.");
          }
        }
      );
    };

    if (!currentSubject.trim()) {
      item.subject.setAsync(subjectText, (setSubjectResult) => {
        if (setSubjectResult.status === Office.AsyncResultStatus.Succeeded) {
          continueWithBody();
        } else {
          setStatus("Betreff konnte nicht gesetzt werden.");
        }
      });
    } else {
      continueWithBody();
    }
  });
}