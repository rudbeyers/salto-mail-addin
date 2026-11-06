/* Mail Agent – knoppen zonder zijpaneel (Salto Works):
 *   "Antwoord met Robaws" in de leesweergave (1.2) en "Aanvullen met Robaws" in het antwoordvenster (1.1).
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
    const token = await ssoToken(false);
    if (!key && !token) {
      await note(item, "ErrorMessage", "Niet aangemeld: open één keer de Mail Agent bij een mail om je aan te melden.");
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
        key, token, action: "aanvullen", user: profiel.emailAddress || "", user_name: profiel.displayName || "",
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

// ------------------------------------------------------------------ knop "Antwoord met Robaws" in de leesweergave (1.2)
// Eén klik vanuit de geopende mail: concept ophalen (bewaard concept van MAILBOX = meteen en gratis, anders Sonnet)
// en meteen een ingevuld antwoordvenster openen. Lukt openen niet (sommige Outlook-apps), dan een duidelijke melding.
// Nieuwe Outlook voor Windows meldt bij displayReplyForm "gelukt" maar opent niets (getest 3/11/2026, op het web werkt
// het wel). Daar zet de server het antwoord als CONCEPT in het gesprek (backend-actie concept_maken, v3.8).
function nieuweOutlook() {
  try { return String(Office.context.mailbox.diagnostics.hostName || "").toLowerCase() === "newoutlookwindows"; }
  catch (e) { return false; }
}
function openConcept(id) {   // proberen het concept te openen; lukt het niet, dan staat het toch in het gesprek
  try {
    const ews = Office.context.mailbox.convertToEwsId(id, Office.MailboxEnums.RestVersion.v2_0);
    Office.context.mailbox.displayMessageForm(ews);
  } catch (e) { /* niet erg */ }
}

function openAntwoord(item, html) {
  const poging = (data) => new Promise((resolve) => {
    try {
      if (!item.displayReplyFormAsync) return resolve(false);
      item.displayReplyFormAsync(data, (r) => resolve(r.status === Office.AsyncResultStatus.Succeeded));
    } catch (e) { resolve(false); }
  });
  return poging(html).then((ok) => ok || poging({ htmlBody: html })).then((ok) => {
    if (ok) return true;
    try { item.displayReplyForm(html); return true; } catch (e) { return false; }
  });
}

async function antwoordMetRobaws(event) {
  const item = Office.context.mailbox.item;
  try {
    let key = "";
    try { key = Office.context.roamingSettings.get(KEY_SETTING) || ""; } catch (e) { /* geen instellingen */ }
    const token = await ssoToken(false);
    if (!key && !token) {
      await note(item, "ErrorMessage", "Niet aangemeld: open één keer de Mail Agent bij een mail om je aan te melden.");
      return;
    }
    const nieuw = nieuweOutlook();
    await note(item, "ProgressIndicator", nieuw ? "Concept-antwoord met Robaws-gegevens wordt in dit gesprek gezet…"
                                                : "Antwoord met Robaws-gegevens wordt klaargezet…");
    const body = await get((cb) => item.body.getAsync(Office.CoercionType.Text, cb));
    const from = item.from || item.sender || {};
    const profiel = Office.context.mailbox.userProfile || {};
    const d = item.dateTimeCreated ? new Date(item.dateTimeCreated) : null;
    const res = await fetch(API, {
      method: "POST", headers: { "Content-Type": "text/plain;charset=UTF-8" },
      body: JSON.stringify({
        key, token, action: nieuw ? "concept_maken" : "antwoord", user: profiel.emailAddress || "", user_name: profiel.displayName || "",
        item_id: item.itemId || "", internet_message_id: item.internetMessageId || "",
        subject: item.subject || "", from_email: (from.emailAddress || "").toLowerCase(), from_name: from.displayName || "",
        date: d ? d.toLocaleString("nl-BE", { dateStyle: "short", timeStyle: "short" }) : "",
        body: String(body || "").slice(0, 5000),
      }),
    });
    let j = {};
    try { j = await res.json(); } catch (e) { /* lege body */ }
    if (!res.ok || j.ok === false || !j.tekst) {
      await note(item, "ErrorMessage", j.fout || ("Concept maken mislukt (" + res.status + ")."));
      return;
    }
    if (nieuw) {
      if (j.concept_id) openConcept(j.concept_id);
      const n = (j.na_te_kijken || []).length;
      await note(item, "InformationalMessage", (j.bestond ? "Het concept-antwoord stond al klaar" : "Concept-antwoord staat klaar")
        + ` in dit gesprek (en bij Concepten)${n ? ` – ${n} punten nakijken` : ""}. Vul [ … ] aan en verstuur zelf.`);
      return;
    }
    const html = simpleHtml(zonderNaam(j.tekst, profiel.displayName));
    if (await openAntwoord(item, html)) {
      const n = (j.na_te_kijken || []).length;
      await note(item, "InformationalMessage",
        `Antwoord geopend${n ? ` – ${n} punten zelf nakijken` : ""}. Vul [ … ] aan en verstuur zelf.`);
    } else {
      await note(item, "ErrorMessage",
        "Deze Outlook liet het openen niet toe. Klik Beantwoorden en daarna ‘Aanvullen met Robaws’ (het concept staat klaar).");
    }
  } catch (e) {
    await note(item, "ErrorMessage", "Geen verbinding met Salto Data. Probeer opnieuw.");
  } finally {
    if (event && event.completed) event.completed();
  }
}

Office.onReady(() => {
  if (Office.actions && Office.actions.associate) {
    Office.actions.associate("aanvullen", aanvullen);
    Office.actions.associate("antwoordMetRobaws", antwoordMetRobaws);
  }
});
window.aanvullen = aanvullen;   // oudere Outlook-versies zoeken de functie globaal
window.antwoordMetRobaws = antwoordMetRobaws;
