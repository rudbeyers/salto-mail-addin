/* Mail Agent – Outlook-taakvenster (Salto Works)
 * Praat met de Peliqan-endpoint /mail/api. Body als text/plain → geen CORS-preflight. */
"use strict";

const API = "https://api.eu.peliqan.io/2704/mail/api";
const KEY_SETTING = "mailAgentKey";
const MAX_ATT_BYTES = 8 * 1024 * 1024;

const $ = (id) => document.getElementById(id);
const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

let mail = null;      // gegevens van de geopende mail
let state = { client: null, projects: [], lookup: null, extract: null };
let loadSeq = 0;

// ------------------------------------------------------------------ helpers
function getKey() {
  try { return Office.context.roamingSettings.get(KEY_SETTING) || ""; } catch (e) { return ""; }
}
function saveKey(k) {
  return new Promise((resolve) => {
    Office.context.roamingSettings.set(KEY_SETTING, k);
    Office.context.roamingSettings.saveAsync(() => resolve());
  });
}
function status(text, kind) {
  const el = $("status");
  el.textContent = text || "";
  el.className = "msg" + (kind ? " " + kind : "");
}
async function api(action, payload) {
  const body = Object.assign({ key: getKey(), action, user: userEmail(), item_id: mail ? mail.item_id : "" }, payload || {});
  let res;
  try {
    res = await fetch(API, { method: "POST", headers: { "Content-Type": "text/plain;charset=UTF-8" }, body: JSON.stringify(body) });
  } catch (e) {
    throw new Error("Geen verbinding met Salto Data. Probeer opnieuw.");
  }
  let j = {};
  try { j = await res.json(); } catch (e) { /* lege of foute body */ }
  if (res.status === 401) { showSettings("Sleutel ontbreekt of is fout."); }
  if (!res.ok || j.ok === false) throw new Error(j.fout || ("Fout " + res.status));
  return j;
}
function userEmail() {
  try { return Office.context.mailbox.userProfile.emailAddress || ""; } catch (e) { return ""; }
}
function busy(btn, on, label) {
  if (!btn) return;
  if (on) { btn.dataset.label = btn.textContent; btn.textContent = label || "Bezig…"; btn.disabled = true; }
  else { btn.textContent = btn.dataset.label || btn.textContent; btn.disabled = false; }
}
function show(id, on) { $(id).hidden = !on; }

// ------------------------------------------------------------------ mail lezen
function readMail() {
  const item = Office.context.mailbox.item;
  if (!item) return Promise.resolve(null);
  return new Promise((resolve) => {
    item.body.getAsync(Office.CoercionType.Text, (r) => {
      const from = item.from || item.sender || {};
      const d = item.dateTimeCreated ? new Date(item.dateTimeCreated) : null;
      resolve({
        item_id: item.itemId || "",
        subject: item.subject || "",
        from_email: (from.emailAddress || "").toLowerCase(),
        from_name: from.displayName || "",
        date: d ? d.toLocaleString("nl-BE", { dateStyle: "short", timeStyle: "short" }) : "",
        body: r.status === Office.AsyncResultStatus.Succeeded ? (r.value || "") : "",
        attachments: (item.attachments || []).filter((a) => !a.isInline && a.attachmentType === "file"),
      });
    });
  });
}
function mailPayload(limit) {
  return { subject: mail.subject, from_email: mail.from_email, from_name: mail.from_name, date: mail.date,
           body: limit ? mail.body.slice(0, limit) : mail.body };
}

// ------------------------------------------------------------------ laden
async function load() {
  const seq = ++loadSeq;
  ["secSender", "secNewClient", "secProject", "secActions"].forEach((s) => show(s, false));
  show("loading", true);
  $("loading").textContent = "Mail wordt bekeken…";
  status("");
  state = { client: null, projects: [], lookup: null, extract: null };
  mail = await readMail();
  if (seq !== loadSeq) return;
  if (!mail) { $("loading").textContent = "Open een mail om te beginnen."; return; }
  if (!getKey()) { show("loading", false); showSettings("Vul eerst de sleutel in."); return; }

  $("inTitle").value = mail.subject;
  $("inDeadline").value = "";
  $("inTaskNote").value = "";
  $("inAttNote").value = "";
  renderAttachments();

  try {
    const r = await api("lookup", mailPayload(3000));
    if (seq !== loadSeq) return;
    state.lookup = r;
    show("loading", false);
    if (r.gevonden) setClient(r.klant, r.hoe, r.projecten, r.voorstel_project_id, r.voorstel_waarom, r.offertes, r.werkbonnen);
    else showUnknown();
  } catch (e) {
    if (seq !== loadSeq) return;
    $("loading").textContent = e.message;
  }
}

function setClient(klant, hoe, projecten, voorstelId, waarom, offertes, werkbonnen) {
  state.client = klant;
  state.projects = projecten || [];
  show("secSender", true);
  show("unknown", false);
  show("secNewClient", false);
  $("clientCard").hidden = false;
  $("clientCard").innerHTML =
    `<div class="name">${esc(klant.naam)}</div>` +
    (klant.adres ? `<div class="sub">${esc(klant.adres)}</div>` : "") +
    ((klant.email || klant.tel) ? `<div class="sub">${esc([klant.email, klant.tel].filter(Boolean).join(" · "))}</div>` : "") +
    (hoe ? `<span class="badge">herkend op ${esc(hoe)}</span>` : "");
  fillProjects(voorstelId, waarom);
  renderContext(offertes || [], werkbonnen || []);
  show("secProject", true);
  show("secActions", true);
}

function fillProjects(voorstelId, waarom) {
  const sel = $("selProject");
  sel.innerHTML = "";
  const none = document.createElement("option");
  none.value = ""; none.textContent = state.projects.length ? "— geen project —" : "— klant heeft nog geen project —";
  sel.appendChild(none);
  state.projects.forEach((p) => {
    const o = document.createElement("option");
    o.value = p.id;
    o.textContent = [p.nr, p.naam !== p.nr ? p.naam : "", p.status ? `(${p.status})` : ""].filter(Boolean).join(" ");
    sel.appendChild(o);
  });
  sel.value = voorstelId || "";
  $("projectWhy").textContent = voorstelId ? `Voorstel: ${waarom}` : "";
}

function renderContext(offertes, werkbonnen) {
  const rows = [];
  offertes.forEach((o) => rows.push(`<div class="item"><b>${esc(o.nr)}</b> ${esc(o.titel)}<div class="sub">offerte · ${esc(o.status)} · ${esc(o.datum)}</div></div>`));
  werkbonnen.forEach((w) => rows.push(`<div class="item"><b>${esc(w.nr)}</b> ${esc(w.titel)}<div class="sub">werkbon · ${esc(w.status)} · ${esc(w.datum)}</div></div>`));
  $("ctx").innerHTML = rows.join("") || '<div class="muted small">Nog geen offertes of werkbonnen.</div>';
  $("ctxDetails").hidden = false;
}

function showUnknown() {
  show("secSender", true);
  $("clientCard").hidden = true;
  show("unknown", true);
  $("unknownText").textContent = `${mail.from_name || mail.from_email} is nog geen klant in Robaws.`;
  $("inSearchClient").value = "";
  $("searchResults").innerHTML = "";
  // acties kunnen ook zonder klant (taak zonder project)
  state.projects = [];
  fillProjects(null, "");
  $("ctxDetails").hidden = true;
  show("secProject", false);
  show("secActions", true);
}

// ------------------------------------------------------------------ zoeken
let searchTimer = null;
function onSearch() {
  clearTimeout(searchTimer);
  const q = $("inSearchClient").value.trim();
  if (q.length < 2) { $("searchResults").innerHTML = ""; return; }
  searchTimer = setTimeout(async () => {
    try {
      const r = await api("search", { q });
      if ($("inSearchClient").value.trim() !== q) return;
      const box = $("searchResults");
      box.innerHTML = "";
      (r.klanten || []).forEach((k) => box.appendChild(resultBtn(`👤 ${k.naam}`, k.adres, () => pickClient(k))));
      (r.projecten || []).forEach((p) => box.appendChild(resultBtn(`📁 ${p.nr} ${p.naam}`, [p.klant, p.status].filter(Boolean).join(" · "),
        () => pickClient({ id: p.klant_id, naam: p.klant || "klant " + p.klant_id }, p.id))));
      if (!box.children.length) box.innerHTML = '<div class="muted small">Niets gevonden.</div>';
    } catch (e) { status(e.message, "err"); }
  }, 300);
}
function resultBtn(title, sub, fn) {
  const b = document.createElement("button");
  b.className = "item";
  b.innerHTML = `${esc(title)}${sub ? `<div class="sub">${esc(sub)}</div>` : ""}`;
  b.addEventListener("click", fn);
  return b;
}
async function pickClient(k, projectId) {
  status("Klant wordt geladen…");
  try {
    // context van de gekozen klant ophalen via lookup op naam is onbetrouwbaar → zoek projecten via search
    const r = await api("search", { q: k.naam });
    const projs = (r.projecten || []).filter((p) => String(p.klant_id) === String(k.id));
    const full = (r.klanten || []).find((c) => String(c.id) === String(k.id)) || k;
    setClient(full, "jouw keuze", projs, projectId || (projs.length === 1 ? projs[0].id : null),
      projectId ? "gekozen in de zoekresultaten" : (projs.length === 1 ? "enige project van deze klant" : ""), [], []);
    $("ctxDetails").hidden = true;
    status("");
  } catch (e) { status(e.message, "err"); }
}

// ------------------------------------------------------------------ nieuwe klant
async function onExtract() {
  const btn = $("btnExtract");
  busy(btn, true, "Handtekening lezen…");
  status("");
  try {
    const r = await api("client_extract", mailPayload());
    state.extract = r.klant || {};
    document.querySelectorAll("#secNewClient [data-k]").forEach((inp) => { inp.value = state.extract[inp.dataset.k] || ""; });
    const d = r.dubbels || [];
    const warn = $("dupWarn");
    if (d.length) {
      warn.hidden = false;
      warn.innerHTML = "⚠️ Bestaat deze klant al?";
      d.forEach((c) => {
        const b = document.createElement("button");
        b.textContent = `Gebruik ${c.naam} (${c.waarom})`;
        b.addEventListener("click", () => { show("secNewClient", false); pickClient(c); });
        warn.appendChild(document.createElement("br"));
        warn.appendChild(b);
      });
    } else { warn.hidden = true; }
    show("secNewClient", true);
    $("btnCreateClient").textContent = d.length ? "✓ Toch nieuwe klant maken" : "✓ Maak klant in Robaws";
  } catch (e) { status(e.message, "err"); }
  busy(btn, false);
}
async function onCreateClient() {
  const btn = $("btnCreateClient");
  const klant = {};
  document.querySelectorAll("#secNewClient [data-k]").forEach((inp) => { klant[inp.dataset.k] = inp.value.trim(); });
  if (!klant.naam) { status("Vul minstens een naam in.", "err"); return; }
  busy(btn, true, "Klant aanmaken…");
  try {
    const r = await api("client_create", { klant });
    show("secNewClient", false);
    setClient(r.klant, "net aangemaakt", [], null, "", [], []);
    $("ctxDetails").hidden = true;
    status(`✅ Klant ${r.klant.naam} aangemaakt in Robaws.`, "ok");
  } catch (e) { status(e.message, "err"); }
  busy(btn, false);
}

// ------------------------------------------------------------------ taak
async function onTask() {
  const btn = $("btnTask");
  busy(btn, true, "Taak aanmaken…");
  status("");
  try {
    const r = await api("task", Object.assign(mailPayload(), {
      title: $("inTitle").value.trim() || mail.subject,
      deadline: $("inDeadline").value || null,
      note: $("inTaskNote").value.trim(),
      project_id: $("selProject").value || null,
    }));
    status(`✅ Taak “${r.titel}” staat in Robaws.`, "ok");
  } catch (e) { status(e.message, "err"); }
  busy(btn, false);
}

// ------------------------------------------------------------------ aan project hangen
function renderAttachments() {
  const box = $("attList");
  box.innerHTML = "";
  if (!mail.attachments.length) { box.innerHTML = '<div class="muted small">Geen bijlagen.</div>'; return; }
  mail.attachments.forEach((a) => {
    const lab = document.createElement("label");
    lab.className = "item";
    const tooBig = a.size > MAX_ATT_BYTES;
    lab.innerHTML = `<input type="checkbox" value="${esc(a.id)}" ${tooBig ? "disabled" : "checked"}> ` +
      `<span>📎 ${esc(a.name)} <span class="sub">${Math.max(1, Math.round(a.size / 1024))} kB${tooBig ? " · te groot" : ""}</span></span>`;
    box.appendChild(lab);
  });
}
function attachmentBase64(id) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentContentAsync(id, (r) => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) return reject(new Error("Bijlage niet leesbaar"));
      if (r.value.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) return reject(new Error("Bijlagetype niet ondersteund"));
      resolve(r.value.content);
    });
  });
}
async function onAttach() {
  const pid = $("selProject").value;
  if (!pid) {
    if ($("secProject").hidden) {
      status("Zoek hierboven eerst de klant of het project (bv. naam of projectnummer) en klik het aan.", "err");
      $("inSearchClient").focus();
    } else if (!state.projects.length) {
      status("Deze klant heeft nog geen project in Robaws. Maak er eerst een aan, of maak een taak.", "err");
    } else {
      status("Kies hierboven bij Project het project waar de mail bij hoort.", "err");
      $("selProject").focus();
    }
    return;
  }
  const btn = $("btnAttach");
  busy(btn, true, "Mail toevoegen…");
  status("");
  const fouten = [];
  try {
    await api("attach", Object.assign(mailPayload(), { project_id: pid, note: $("inAttNote").value.trim() }));
    const ids = [...document.querySelectorAll("#attList input:checked")].map((i) => i.value);
    let n = 0;
    for (const id of ids) {
      const a = mail.attachments.find((x) => x.id === id);
      btn.textContent = `Bijlage ${++n}/${ids.length}…`;
      try {
        const b64 = await attachmentBase64(id);
        await api("attach_file", { project_id: pid, name: a.name, contentType: a.contentType, base64: b64 });
      } catch (e) { fouten.push(a.name); }
    }
    const pnaam = $("selProject").selectedOptions[0].textContent;
    if (fouten.length) status(`⚠️ Mail staat op ${pnaam}, maar niet gelukt: ${fouten.join(", ")}`, "err");
    else status(`✅ Mail${ids.length ? ` + ${ids.length} bijlage${ids.length > 1 ? "n" : ""}` : ""} staat op ${pnaam}.`, "ok");
  } catch (e) { status(e.message, "err"); }
  busy(btn, false);
}

// ------------------------------------------------------------------ instellingen
function showSettings(msg) {
  show("settings", true);
  $("inKey").value = getKey();
  $("settingsMsg").textContent = msg || "";
  $("settingsMsg").className = "msg" + (msg ? " err" : "");
}
async function onSaveKey() {
  const btn = $("btnSaveKey");
  busy(btn, true, "Testen…");
  await saveKey($("inKey").value.trim());
  try {
    const r = await api("ping");
    $("settingsMsg").textContent = `✅ Verbonden (versie ${r.versie}).`;
    $("settingsMsg").className = "msg ok";
    setTimeout(() => { show("settings", false); load(); }, 700);
  } catch (e) {
    $("settingsMsg").textContent = e.message;
    $("settingsMsg").className = "msg err";
  }
  busy(btn, false);
}

// ------------------------------------------------------------------ start
Office.onReady(() => {
  $("btnSettings").addEventListener("click", () => ($("settings").hidden ? showSettings() : show("settings", false)));
  $("btnCloseSettings").addEventListener("click", () => show("settings", false));
  $("btnSaveKey").addEventListener("click", onSaveKey);
  $("btnExtract").addEventListener("click", onExtract);
  $("btnCancelClient").addEventListener("click", () => show("secNewClient", false));
  $("btnCreateClient").addEventListener("click", onCreateClient);
  $("btnTask").addEventListener("click", onTask);
  $("btnAttach").addEventListener("click", onAttach);
  $("inSearchClient").addEventListener("input", onSearch);
  document.querySelectorAll(".tab").forEach((t) => t.addEventListener("click", () => {
    document.querySelectorAll(".tab").forEach((x) => x.classList.toggle("active", x === t));
    document.querySelectorAll(".tabpane").forEach((p) => { p.hidden = p.id !== t.dataset.tab; });
  }));
  // vastgezet venster: opnieuw laden bij een andere mail
  try {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => load());
  } catch (e) { /* oudere Outlook */ }
  load();
});
