/* Mail Agent – knop "Aanvullen met Robaws" in het antwoordvenster (Salto Works)
 * Zet een concept-antwoord met feiten uit Robaws bovenaan de mail die de gebruiker aan het schrijven is.
 * Hergebruikt het concept dat MAILBOX al klaarzette (backend-actie "aanvullen"). Verstuurt nooit iets. */
"use strict";

const API = "https://api.eu.peliqan.io/2704/mail/api";
const KEY_SETTING = "mailAgentKey";
const NOTE_ID = "saltodata";

function esc(s) {
  return String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
function get(fn) {   // Office-async als belofte
  return new Promise((resolve) => {
    try { fn((r) => resolve(r && r.status === Office.AsyncResultStatus.Succeeded ? r.value : null)); }
    catch (e) { resolve(null); }
  });
}
function lukt(fn) {  // Office-async -> true/false
  return new Promise((resolve) => {
    try { fn((r) => resolve(!!r && r.status === Office.AsyncResultStatus.Succeeded)); }
    catch (e) { resolve(false); }
  });
}
function note(item, type, message) {
  const M = Office.MailboxEnums.ItemNotificationMessageType;
  const data = { type: M[type], message: String(message).slice(0, 150) };
  if (type === "InformationalMessage") { data.icon = "Icon.16x16"; data.persistent = false; }
  return get((cb) => item.notificationMessages.replaceAsync(NOTE_ID, data, cb));
}
// "Met vriendelijke groeten," behouden, maar de naam eronder weglaten: de handtekening van Outlook volgt
function zonderNaam(tekst, naam) {
  const t = String(tekst || "").trimEnd();
  const lijnen = t.split("\n");
  if (naam && lijnen.length > 1 && lijnen[lijnen.length - 1].trim().toLowerCase() === naam.trim().toLowerCase()) lijnen.pop();
  return lijnen.join("\n");
}
function simpleHtml(t) {
  return String(t || "").split(/\n{2,}/).map((p) => "<p>" + esc(p).replace(/\n/g, "<br>") + "</p>").join("") + "<p></p>";
}

async function aanvullen(event) {
  const item = Office.context.mailbox.item;
  try {
    let key = "";
    try { key = Office.context.roamingSettings.get(KEY_SETTING) || ""; } catch (e) { /* geen instellingen */ }
    if (!key) {
      await note(item, "ErrorMessage", "Sleutel ontbreekt: open één keer de Mail Agent bij een mail en vul de sleutel in (⚙).");
      return;
    }
    await note(item, "ProgressIndicator", "Concept met Robaws-gegevens wordt klaargezet…");
    const [to, subject, body] = await Promise.all([
      get((cb) => item.to.getAsync(cb)),
      get((cb) => item.subject.getAsync(cb)),
      get((cb) => item.body.getAsync(Office.CoercionType.Text, cb)),
    ]);
    const ontvanger = (to || [])[0] || {};
    const profiel = Office.context.mailbox.userProfile || {};
    const res = await fetch(API, {
      method: "POST", headers: { "Content-Type": "text/plain;charset=UTF-8" },
      body: JSON.stringify({
        key, action: "aanvullen", user: profiel.emailAddress || "", user_name: profiel.displayName || "",
        conversation_id: item.conversationId || "", from_email: (ontvanger.emailAddress || "").toLowerCase(),
        from_name: ontvanger.displayName || "", subject: String(subject || "").replace(/^\s*(re|antw|aw)\s*:\s*/i, ""),
        body: String(body || "").slice(0, 5000), item_id: "compose",
      }),
    });
    let j = {};
    try { j = await res.json(); } catch (e) { /* lege body */ }
    if (!res.ok || j.ok === false || !j.tekst) {
      await note(item, "ErrorMessage", j.fout || ("Concept maken mislukt (" + res.status + ")."));
      return;
    }
    const html = simpleHtml(zonderNaam(j.tekst, profiel.displayName));
    const opts = { coercionType: Office.CoercionType.Html };
    const ingevoegd = (await lukt((cb) => item.body.prependAsync(html, opts, cb)))
                   || (await lukt((cb) => item.body.setSelectedDataAsync(html, opts, cb)));
    if (!ingevoegd) {
      await note(item, "ErrorMessage", "Invoegen lukte niet. Gebruik de Mail Agent bij de mail en kies ‘Kopiëren’.");
      return;
    }
    const n = (j.na_te_kijken || []).length;
    await note(item, "InformationalMessage",
      `Concept ingevoegd${j.klant_gekend ? " met gegevens uit Robaws" : ""}. Vul [ … ] aan${n ? ` (${n} punten)` : ""} en verstuur zelf.`);
  } catch (e) {
    await note(item, "ErrorMessage", "Geen verbinding met Salto Data. Probeer opnieuw.");
  } finally {
    if (event && event.completed) event.completed();
  }
}

Office.onReady(() => {
  if (Office.actions && Office.actions.associate) Office.actions.associate("aanvullen", aanvullen);
});
window.aanvullen = aanvullen;   // oudere Outlook-versies zoeken de functie globaal
