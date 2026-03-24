/* global Office */

const rechnungMap = {
  r1: "Die Rechnungsstruktur",
  r2: "Die Vollständigkeit der Nachweise / Anlagen",
  r3: "Die Nachvollziehbarkeit der Nachweise / Anlagen",
  r4: "Die Einhaltung vertraglicher Rahmenbedingungen",
  r5: "Die Übereinstimmung der abgerechneten Einheitspreise mit dem Leistungsverzeichnis / Auftrag",
  r6: "Auf Rechnungspositionen, für die kein Auftrag vorliegt. Diese müssen gestrichen werden.",
  r7: "Die rechnerische Richtigkeit sowie die korrekte Übertragung der Werte",
  r8: "Die Einhaltung der vereinbarten Auftragssumme",
  r9: "Die Vollständigkeit und Richtigkeit der formalen Angaben"
};

const nachtragMap = {
  n1: "Die Vollständigkeit der Nachweise / Anlagen",
  n2: "Die Nachvollziehbarkeit der Nachweise / Anlagen",
  n3: "Die Übereinstimmung der angesetzten Einheitspreise mit den vertraglichen Vereinbarungen",
  n4: "Die Angemessenheit der Preisansätze",
  n5: "Die Nachvollziehbarkeit der Nachtragsbegründung",
  n6: "Auf Leistungen, die bereits vom bestehenden Auftrag umfasst sind",
  n7: "Die Herleitung und Nachvollziehbarkeit der Mengenansätze",
  n8: "Die rechnerische Richtigkeit sowie die korrekte Übertragung der Werte",
  n9: "Die Vollständigkeit und Richtigkeit der formalen Angaben"
};

Office.onReady(() => {
  const typeRechnung = document.getElementById("typeRechnung");
  const typeNachtrag = document.getElementById("typeNachtrag");
  const btnInsert = document.getElementById("btnInsert");

  if (typeRechnung) typeRechnung.addEventListener("change", toggleSections);
  if (typeNachtrag) typeNachtrag.addEventListener("change", toggleSections);

  if (btnInsert) {
    btnInsert.addEventListener("click", () => {
      const selectedType = document.querySelector('input[name="mailType"]:checked')?.value;

      if (selectedType === "nachtrag") {
        insertNachtrag();
      } else {
        insertRechnung();
      }
    });
  }

  toggleSections();
});

function toggleSections() {
  const isNachtrag = document.getElementById("typeNachtrag")?.checked;
  const rechnungSection = document.getElementById("rechnungSection");
  const nachtragSection = document.getElementById("nachtragSection");

  if (isNachtrag) {
    rechnungSection?.classList.add("hidden");
    nachtragSection?.classList.remove("hidden");
  } else {
    nachtragSection?.classList.add("hidden");
    rechnungSection?.classList.remove("hidden");
  }
}

function setStatus(message) {
  const status = document.getElementById("status");
  if (status) status.textContent = message;
}

function getCheckedMapped(map) {
  const reasons = [];

  Object.keys(map).forEach((key) => {
    const checkbox = document.getElementById(key);
    const note = document.getElementById(`${key}_note`)?.value?.trim() || "";

    if (checkbox && checkbox.checked) {
      let text = map[key];

      if (note) {
        text += ` (Ergebnis der Stichprobe(n): ${note})`;
      }

      reasons.push(text);
    }
  });

  return reasons;
}

function buildHtmlMail(intro, reasons, closing) {
  let html = "";

  html += '<div style="font-family:Calibri, Arial, sans-serif; font-size:11pt; line-height:1.35; color:#222222;">';
  html += '<p style="margin:0 0 12pt 0;">Sehr geehrte Damen und Herren,</p>';
  html += `<p style="margin:0 0 12pt 0;">${escapeHtml(intro)}</p>`;
  html += '<p style="margin:0 0 12pt 0;">Die durchgeführten Stichproben ergeben nachfolgenden Prüf- bzw. Überarbeitungsbedarf (ohne Anspruch auf Vollständigkeit).</p>';
  html += '<p style="margin:0 0 6pt 0;">Bitte prüfen Sie:</p>';

  html += '<ul style="margin:0 0 12pt 24px; padding:0;">';
  reasons.forEach((reason) => {
    html += `<li style="margin:0 0 6pt 0;">${escapeHtml(reason)}</li>`;
  });
  html += '</ul>';

  html += `<p style="margin:0 0 12pt 0;">${escapeHtml(closing)}</p>`;
  html += '<p style="margin:0;">Für Rückfragen stehen wir gerne zur Verfügung.</p>';
  html += '</div>';

  return html;
}

function escapeHtml(text) {
  return text
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function buildSubject() {
  const useSubject = document.getElementById("useSubject")?.checked;
  if (!useSubject) return null;

  const parts = [
    (document.getElementById("proj")?.value || "").trim(),
    (document.getElementById("verg")?.value || "").trim(),
    (document.getElementById("auftr")?.value || "").trim(),
    (document.getElementById("subjLast")?.value || "").trim()
  ].filter(Boolean);

  return parts.length ? parts.join(" | ") : null;
}

function insertRechnung() {
  const reasons = getCheckedMapped(rechnungMap);

  if (reasons.length === 0) {
    setStatus("Bitte mindestens einen Prüfpunkt auswählen.");
    return;
  }

  const intro = "im Zuge der Plausibilitätsprüfung der eingereichten Rechnung ergeben sich Unstimmigkeiten, die einer Weitergabe an den Bauherrn und damit der Beauftragung derzeit entgegenstehen.";
  const closing = "Bitte überarbeiten Sie Ihre Rechnungsprüfung auch über die genannten Punkte hinaus und legen Sie die Prüfung erneut vor.";

  const html = buildHtmlMail(intro, reasons, closing);
  const subject = buildSubject();

  insertTemplate(html, subject);
}

function insertNachtrag() {
  const reasons = getCheckedMapped(nachtragMap);

  if (reasons.length === 0) {
    setStatus("Bitte mindestens einen Prüfpunkt auswählen.");
    return;
  }

  const intro = "im Zuge der Plausibilitätsprüfung des eingereichten Nachtrags ergeben sich Unstimmigkeiten, die einer Weitergabe an den Bauherrn und damit der Beauftragung derzeit entgegenstehen.";
  const closing = "Bitte überarbeiten Sie Ihre Nachtragsprüfung auch über die genannten Punkte hinaus und legen Sie die Prüfung erneut vor.";

  const html = buildHtmlMail(intro, reasons, closing);
  const subject = buildSubject();

  insertTemplate(html, subject);
}

function insertTemplate(htmlTemplate, subjectText) {
  const item = Office.context.mailbox.item;

  if (!item || !item.body) {
    setStatus("Keine bearbeitbare E-Mail gefunden.");
    return;
  }

  item.body.prependAsync(
    htmlTemplate,
    { coercionType: Office.CoercionType.Html },
    (bodyResult) => {
      if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
        setStatus("Fehler beim Einfügen des Textes.");
        return;
      }

      if (subjectText && item.subject) {
        item.subject.setAsync(subjectText, (subjectResult) => {
          if (subjectResult.status === Office.AsyncResultStatus.Succeeded) {
            setStatus("Text eingefügt und Betreff gesetzt.");
          } else {
            setStatus("Text eingefügt. Betreff konnte nicht gesetzt werden.");
          }
        });
      } else {
        setStatus("Text eingefügt.");
      }
    }
  );
}