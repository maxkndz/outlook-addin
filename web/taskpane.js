/* global Office */

const rechnungFormalMap = {
  rf1: "das Vorhandensein des Prüfblattes zur Rechnungsprüfung",
  rf2: "die kumulierte Darstellung der Rechnung",
  rf3: "die korrekte Angabe des Leistungszeitraums",
  rf4: "die Vollständigkeit und Richtigkeit der Prüfstempel",
  rf5: "die korrekte Darstellung der bereits geleisteten Zahlungen",
  rf6: "die Einhaltung der vereinbarten Auftragssumme",
  rf7: "die Vollständigkeit der Nachweise / Anlagen"
};

const rechnungInhaltlichMap = {
  ri1: "die Vollständigkeit und Richtigkeit des Prüfblattes zur Rechnungsprüfung",
  ri2: "die Einhaltung vertraglicher Rahmenbedingungen",
  ri3: "die Übereinstimmung der abgerechneten Einheitspreise mit dem Leistungsverzeichnis / Auftrag",
  ri4: "die Abweichungen in Mengen und Massen (max. + 10%), bezogen auf den vereinbarten Leistungsumfang",
  ri5: "die rechnerische Richtigkeit sowie die korrekte Übertragung der Werte",
  ri6: "die Nachvollziehbarkeit der Nachweise / Anlagen"
};

const nachtragFormalMap = {
  nf1: "das Vorhandensein des Prüfblattes zur Nachtragsprüfung",
  nf2: "die Vollständigkeit und Richtigkeit der Prüfstempel",
  nf3: "die Vollständigkeit der Nachweise / Anlagen"
};

const nachtragInhaltlichMap = {
  ni1: "die Vollständigkeit und Richtigkeit des Prüfblattes zur Nachtragsprüfung",
  ni2: "die Nachvollziehbarkeit der Nachtragsbegründung",
  ni3: "die Leistungen, die bereits vom bestehenden Auftrag umfasst sind",
  ni4: "die Übereinstimmung der angesetzten Einheitspreise mit den vertraglichen Vereinbarungen bei Mehrmengen",
  ni5: "die Angemessenheit der Preisansätze",
  ni6: "die Herleitung und Nachvollziehbarkeit der Mengenansätze",
  ni7: "die rechnerische Richtigkeit sowie die korrekte Übertragung der Werte",
  ni8: "die Nachvollziehbarkeit der Nachweise / Anlagen"
};

const SR_START_MARKER = "SR_STANDARDS_BLOCK_START";
const SR_END_MARKER = "SR_STANDARDS_BLOCK_END";

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
        text += ` (Hinweis aus Stichprobe(n): ${note})`;
      }

      reasons.push(text);
    }
  });

  return reasons;
}

function buildHiddenMarker(markerText) {
  return `<span style="display:none;mso-hide:all;font-size:1px;line-height:1px;color:#ffffff;">${markerText}</span>`;
}

function buildHtmlMail(intro, formalReasons, inhaltlichReasons, closing) {
  let html = "";

  html += buildHiddenMarker(SR_START_MARKER);
  html += '<div data-sr-standards="block" style="font-family:Calibri, Arial, sans-serif; font-size:11pt; line-height:1.35; color:#222222;">';
  html += '<p style="margin:0 0 12pt 0;">Sehr geehrte Damen und Herren,</p>';
  html += `<p style="margin:0 0 12pt 0;">${escapeHtml(intro)}</p>`;

  if (formalReasons.length > 0) {
    html += '<p style="margin:0 0 6pt 0;">Bitte prüfen Sie folgende formale Aspekte (ohne Anspruch auf Vollständigkeit):</p>';
    html += '<ul style="margin:0 0 12pt 24px; padding:0;">';
    formalReasons.forEach((reason) => {
      html += `<li style="margin:0 0 6pt 0;">${escapeHtml(reason)}</li>`;
    });
    html += '</ul>';
  }

  if (inhaltlichReasons.length > 0) {
    html += '<p style="margin:0 0 6pt 0;">Bitte prüfen Sie folgende inhaltliche Aspekte (ohne Anspruch auf Vollständigkeit):</p>';
    html += '<ul style="margin:0 0 12pt 24px; padding:0;">';
    inhaltlichReasons.forEach((reason) => {
      html += `<li style="margin:0 0 6pt 0;">${escapeHtml(reason)}</li>`;
    });
    html += '</ul>';
  }

  html += `<p style="margin:0 0 12pt 0;">${escapeHtml(closing)}</p>`;
  html += '<p style="margin:0;">Für Rückfragen stehen wir gerne zur Verfügung.</p>';
  html += '</div>';
  html += buildHiddenMarker(SR_END_MARKER);

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
  const formalReasons = getCheckedMapped(rechnungFormalMap);
  const inhaltlichReasons = getCheckedMapped(rechnungInhaltlichMap);

  if (formalReasons.length === 0 && inhaltlichReasons.length === 0) {
    setStatus("Bitte mindestens einen Prüfpunkt auswählen.");
    return;
  }

  const intro = "im Zuge der Plausibilitätsprüfung der eingereichten Rechnung ergeben sich Unstimmigkeiten, die einer Weitergabe an den Bauherrn und damit einer Freigabe entgegenstehen.";
  const closing = "Bitte überarbeiten Sie Ihre Rechnungsprüfung, wenn erforderlich auch über die genannten Punkte hinaus und legen Sie die Prüfung erneut vor.";

  const html = buildHtmlMail(intro, formalReasons, inhaltlichReasons, closing);
  const subject = buildSubject();

  insertTemplate(html, subject);
}

function insertNachtrag() {
  const formalReasons = getCheckedMapped(nachtragFormalMap);
  const inhaltlichReasons = getCheckedMapped(nachtragInhaltlichMap);

  if (formalReasons.length === 0 && inhaltlichReasons.length === 0) {
    setStatus("Bitte mindestens einen Prüfpunkt auswählen.");
    return;
  }

  const intro = "im Zuge der Plausibilitätsprüfung des eingereichten Nachtrags ergeben sich Unstimmigkeiten, die einer Weitergabe an den Bauherrn und damit der Beauftragung derzeit entgegenstehen.";
  const closing = "Bitte überarbeiten Sie Ihre Nachtragsprüfung, wenn erforderlich auch über die genannten Punkte hinaus und legen Sie die Prüfung erneut vor.";

  const html = buildHtmlMail(intro, formalReasons, inhaltlichReasons, closing);
  const subject = buildSubject();

  insertTemplate(html, subject);
}

function replaceSrBlockInHtml(bodyHtml, blockHtml) {
  const escapedStart = SR_START_MARKER.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const escapedEnd = SR_END_MARKER.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

  const markerRegex = new RegExp(
    `${escapedStart}[\\s\\S]*?${escapedEnd}`,
    "i"
  );

  if (markerRegex.test(bodyHtml)) {
    return bodyHtml.replace(markerRegex, blockHtml);
  }

  return blockHtml + bodyHtml;
}

function insertTemplate(htmlTemplate, subjectText) {
  const item = Office.context.mailbox.item;

  if (!item || !item.body) {
    setStatus("Keine bearbeitbare E-Mail gefunden.");
    return;
  }

  item.body.getAsync(Office.CoercionType.Html, (getResult) => {
    if (getResult.status !== Office.AsyncResultStatus.Succeeded) {
      setStatus("Fehler beim Auslesen des E-Mail-Textes.");
      return;
    }

    const currentHtml = getResult.value || "";
    const updatedHtml = replaceSrBlockInHtml(currentHtml, htmlTemplate);

    item.body.setAsync(
      updatedHtml,
      { coercionType: Office.CoercionType.Html },
      (setResult) => {
        if (setResult.status !== Office.AsyncResultStatus.Succeeded) {
          setStatus("Fehler beim Aktualisieren des Textes.");
          return;
        }

        if (subjectText && item.subject) {
          item.subject.setAsync(subjectText, (subjectResult) => {
            if (subjectResult.status === Office.AsyncResultStatus.Succeeded) {
              setStatus("Text aktualisiert und Betreff gesetzt.");
            } else {
              setStatus("Text aktualisiert. Betreff konnte nicht gesetzt werden.");
            }
          });
        } else {
          setStatus("Text aktualisiert.");
        }
      }
    );
  });
}