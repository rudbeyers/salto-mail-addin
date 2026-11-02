# Mail Agent – Outlook-add-in (Salto Works)

Mail Agent koppelt je mails aan Robaws via Salto Data (Peliqan). Ze werkt op twee plaatsen.

**In het taakvenster bij een geopende mail** (vastpinbaar):
- **Analyse:** het oordeel van MAILBOX (Dringend, Kans, Opvolgen of Info), met de reden en een voorstel.
- **✎ Antwoord:** een concept-antwoord met feiten uit Robaws. Bij Dringend en Kans staat het al klaar. Met **Openen als antwoord** of **Kopiëren** zet je het in je mail.
- **✓ Taak:** maakt een Robaws-taak, al ingevuld met het voorstel. De volledige mail komt eronder als opmerking.
- **📎 Project:** de mail komt als opmerking op het project, de bijlagen als documenten.
- **Nieuwe klant uit handtekening:** leest de gegevens uit de mail, controleert op dubbels en maakt de klant pas na bevestiging aan.

**In de leesweergave**, via de knop **Antwoord met Robaws** (manifest 1.2): één klik, en er opent meteen een antwoord met het concept. Staat het concept al klaar (MAILBOX), dan gaat dat zonder nieuwe AI-kost. Lukt openen niet in een Outlook-versie, dan meldt de knop dat je via Beantwoorden → Aanvullen met Robaws kunt gaan.

**In het antwoordvenster**, via de knop **Aanvullen met Robaws** (manifest 1.1): het concept-antwoord komt bovenaan je mail. Je leest na en verstuurt zelf.

## Hoe het werkt

- Dit is de voorkant: statische bestanden op GitHub Pages (`taskpane.html/.js/.css`, `manifest.xml`).
- De achterkant is de Peliqan API-handler **SALTODATA · 30 Mail-addin**, bereikbaar op `POST https://api.eu.peliqan.io/2704/mail/api`.
- Elke vraag draagt de sleutel (Peliqan-secret `SALTODATA Mail`). Die wordt per gebruiker bewaard in Outlook (roamingSettings), ook voor de knop in het antwoordvenster.
- Logging gebeurt in `saltodata.actions` (module `mail`). Analyse en concept-antwoorden staan in `saltodata.mail_items`.

## Installeren (test, enkel voor jezelf)

1. Open https://aka.ms/olksideload. Dat opent Outlook op het web met het venster **Invoegtoepassingen**.
2. Kies **Mijn invoegtoepassingen** → **Aangepaste invoegtoepassingen** → **Een aangepaste invoegtoepassing toevoegen** → **Toevoegen vanaf URL…**
3. Plak `https://rudbeyers.github.io/salto-mail-addin/manifest.xml` en bevestig.
4. Open een mail en klik op **Mail Agent** (in het lint, of onder ⋯ → Apps).
5. Vul bij ⚙ de sleutel in en klik **Bewaren en testen**.

## Een nieuwe versie uitrollen

> **Let op:** een persoonlijke installatie (via aka.ms/olksideload → *Toevoegen vanaf URL*) leest de manifest maar één keer in
> en wordt nooit vernieuwd. Ze heeft voorrang op de uitrol uit het beheercentrum (zelfde Id). Nieuwe knoppen in het lint
> verschijnen dan pas nadat je de persoonlijke kopie verwijdert (Mijn invoegtoepassingen → Aangepaste → ⋯ → Verwijderen) of
> opnieuw toevoegt met een nieuwe URL (`manifest.xml?v=…`). Voor klanten: altijd via het beheercentrum, met een link (geen bestand).


- **Alleen JS/CSS gewijzigd:** verhoog `?v=` achter `taskpane.js`/`taskpane.css` in `taskpane.html`, of achter `commands.js` in `commands.html`.
- **`taskpane.html` of `commands.html` zelf gewijzigd:** verhoog ook `?v=` in `manifest.xml` (Taskpane.Url, SourceLocation, Commands.Url) en de `<Version>`. Doe daarna in het M365-beheercentrum **Update add-in**.
- Zonder update van het manifest: rechtsklik in het paneel → **Vernieuwen**, of Ctrl+F5 in Outlook online.
