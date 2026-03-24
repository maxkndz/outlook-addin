/* global Office */

const rechnungMap = {
  r1: "Bitte prüfen Sie die Vollständigkeit der Nachweise / Anlagen",
  r2: "Bitte prüfen Sie die Nachvollziehbarkeit der Nachweise / Anlagen",
  r3: "Bitte prüfen Sie die Übereinstimmung der abgerechneten Einheitspreise mit dem Leistungsverzeichnis / Auftrag",
  r4: "Bitte prüfen Sie, ob ausschließlich beauftragte Leistungen abgerechnet wurden",
  r5: "Bitte prüfen Sie die rechnerische Richtigkeit sowie die korrekte Übertragung der Werte",
  r6: "Bitte prüfen Sie die Einhaltung der vereinbarten Auftragssumme",
  r7: "Bitte prüfen Sie die Vollständigkeit und Richtigkeit der formalen Angaben (z. B. Stempel)"
};

const nachtragMap = {
  n1: "Bitte prüfen Sie die Vollständigkeit der Nachweise / Anlagen",
  n2: "Bitte prüfen Sie die Nachvollziehbarkeit der Nachweise / Anlagen",
  n3: "Bitte prüfen Sie die Übereinstimmung der angesetzten Einheitspreise mit den vertraglichen Vereinbarungen",
  n4: "Bitte prüfen Sie die Angemessenheit der Preisansätze",
  n5: "Bitte prüfen Sie die Nachvollziehbarkeit der Nachtragsbegründung",
  n6: "Bitte prüfen Sie, ob Leistungen bereits vom bestehenden Auftrag umfasst sind",
  n7: "Bitte prüfen Sie die Herleitung und Nachvollziehbarkeit der Mengenansätze",
  n8: "Bitte prüfen Sie die rechnerische Richtigkeit sowie die korrekte Übertragung der Werte",
  n9: "Bitte prüfen Sie die Vollständigkeit und Richtigkeit der formalen Angaben (z. B. Stempel)"
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
    const el = document.getElementById(key);
    if (el && el.checked) {
      reasons.push(map[key]);
    }
  });

  return reasons;
}

function buildHtmlMail(intro, reasons, closing) {
  let html = "";

  html += "<p>Sehr geehrte Damen und Herren,</p>";
  html += "<p>&nbsp;</p>";

  html += `<p>${escapeHtml(intro)}</p>`;
  html += "<p>Die durchgeführten Stichproben ergeben nachfolgenden Prüf- bzw. Überarbeitungsbedarf (ohne Anspruch auf Vollständigkeit):</p>";

  html += "<ul>";
  reasons.forEach((reason) => {
    html += `<li>${escapeHtml(reason)}</li>`;
  });
  html += "</ul>";

  html += `<p>${escapeHtml(closing)}</p>`;
  html += "<p>&nbsp;</p>";
  html += "<p>Für Rückfragen stehen wir gerne zur Verfügung.</p>";

  return html;
}

function escapeHtml(text) {
  return text
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function buildSubject(type) {
  const useSubject = document.getElementById("useSubject")?.checked;
  if (!useSubject) return null;

  const projekt = (document.getElementById("proj")?.value || "").trim();
  const vergabe = (document.getElementById("verg")?.value || "").trim();
  const auftragnehmer = (document.getElementById("auftr")?.value || "").trim();
  const laufendeNummer = (document.getElementById("lfdnr")?.value || "").trim();

  const suffixPrefix = type === "nachtrag" ? "Nachtrag" : "Rechnung";
  const suffix = laufendeNummer ? `${suffixPrefix} ${laufendeNummer}` : `${suffixPrefix} XY`;

  return `${projekt} | ${vergabe} | ${auftragnehmer} | ${suffix}`;
}

function insertRechnung() {
  const reasons = getCheckedMapped(rechnungMap);

  if (reasons.length === 0) {
    setStatus("Bitte mindestens einen Prüfpunkt auswählen.");
    return;
  }

  const intro = "im Zuge der Plausibilitätsprüfung der eingereichten Rechnung ergeben sich Unstimmigkeiten, die eine Weitergabe an den Bauherrn und damit eine Freigabe derzeit nicht ermöglichen.";
  const closing = "Wir bitten Sie, die genannten Punkte sowie darüberhinausgehende Aspekte entsprechend zu prüfen, die Unterlagen zu überarbeiten und erneut vorzulegen.";

  const html = buildHtmlMail(intro, reasons, closing);
  const subject = buildSubject("rechnung");

  insertTemplate(html, subject);
}

function insertNachtrag() {
  const reasons = getCheckedMapped(nachtragMap);

  if (reasons.length === 0) {
    setStatus("Bitte mindestens einen Prüfpunkt auswählen.");
    return;
  }

  const intro = "im Zuge der Plausibilitätsprüfung des eingereichten Nachtrags ergeben sich Unstimmigkeiten, die eine Weitergabe an den Bauherrn und damit die Beauftragung derzeit nicht ermöglichen.";
  const closing = "Wir bitten Sie, die genannten Punkte sowie darüberhinausgehende Aspekte entsprechend zu prüfen, die Unterlagen zu überarbeiten und erneut vorzulegen.";

  const html = buildHtmlMail(intro, reasons, closing);
  const subject = buildSubject("nachtrag");

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