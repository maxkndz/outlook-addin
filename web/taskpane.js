/* global Office */

Office.onReady(() => {
  const typeRechnung = document.getElementById("typeRechnung");
  const typeNachtrag = document.getElementById("typeNachtrag");
  const btnInsert = document.getElementById("btnInsert");

  if (typeRechnung) {
    typeRechnung.addEventListener("change", toggleSections);
  }

  if (typeNachtrag) {
    typeNachtrag.addEventListener("change", toggleSections);
  }

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
  if (status) {
    status.textContent = message;
  }
}

function getCheckedReasons(prefix, count) {
  const reasons = [];

  for (let i = 1; i <= count; i++) {
    const checkbox = document.getElementById(`${prefix}${i}`);
    const label = document.querySelector(`label[for="${prefix}${i}"]`);

    if (checkbox && checkbox.checked && label) {
      reasons.push(label.textContent.trim());
    }
  }

  return reasons;
}

function buildHtmlMail(greetingIntro, reasons, closingText) {
  let html = "";
  html += "<p>Sehr geehrte Damen und Herren,</p>";
  html += `<p>${greetingIntro}</p>`;
  html += "<p>Im Einzelnen ergeben sich folgende Punkte:</p>";
  html += "<ul>";

  reasons.forEach((reason) => {
    html += `<li>${escapeHtml(reason)}</li>`;
  });

  html += "</ul>";
  html += `<p>${closingText}</p>`;
  html += "<p>Für Rückfragen stehen wir gerne zur Verfügung.</p>";

  return html;
}

function escapeHtml(text) {
  return text
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function insertRechnung() {
  const reasons = getCheckedReasons("r", 8);

  if (reasons.length === 0) {
    setStatus("Bitte mindestens einen Grund auswählen.");
    return;
  }

  const intro = "im Zuge der Plausibilitätsprüfung der eingereichten Rechnung ergeben sich Unstimmigkeiten, die eine Überprüfung bzw. Freigabe derzeit nicht ermöglichen.";
  const closing = "Wir bitten Sie, die genannten Punkte zu prüfen, die Unterlage(n) entsprechend zu überarbeiten und erneut vorzulegen.";

  const html = buildHtmlMail(intro, reasons, closing);
  insertTemplate(html);
}

function insertNachtrag() {
  const reasons = getCheckedReasons("n", 10);

  if (reasons.length === 0) {
    setStatus("Bitte mindestens einen Grund auswählen.");
    return;
  }

  const intro = "im Zuge der Plausibilitätsprüfung des eingereichten Nachtrags ergeben sich Unstimmigkeiten, die eine Überprüfung bzw. Freigabe derzeit nicht ermöglichen.";
  const closing = "Wir bitten Sie, die genannten Punkte zu prüfen, die Unterlage(n) entsprechend zu überarbeiten und erneut vorzulegen.";

  const html = buildHtmlMail(intro, reasons, closing);
  insertTemplate(html);
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