# Mail Agent – Outlook-add-in (Salto Works)

Taakvenster in Outlook dat een mail koppelt aan Robaws via Salto Data (Peliqan):

- **Afzender herkennen**: op mailadres, domein of naam, met de projecten, offertes en werkbonnen van de klant.
- **Mail als taak**: maakt een Robaws-taak. De volledige mail komt eronder als opmerking.
- **Mail aan project hangen**: de mail komt als opmerking op het project, de bijlagen als documenten.
- **Nieuwe klant uit handtekening**: leest de gegevens uit de mail, controleert op dubbels en maakt pas na bevestiging de klant aan.

## Hoe het werkt

- Dit is de voorkant: statische bestanden op GitHub Pages (`taskpane.html/.js/.css`, `manifest.xml`).
- De achterkant is de Peliqan API-handler **MAIL · 20 Outlook API**, bereikbaar op `POST https://api.eu.peliqan.io/2704/mail/api`.
- Elke vraag draagt de testsleutel (Peliqan-secret `MAIL Testsleutel`). Die wordt per gebruiker bewaard in Outlook (roamingSettings).
- Logging gebeurt in `mail.mail_log`. De koppeling van gebruiker naar Robaws-gebruiker staat in `mail.mail_config` (`user_map`, `default_user_id`).

## Installeren (test, enkel voor jezelf)

1. Open https://aka.ms/olksideload. Dat opent Outlook op het web met het venster **Invoegtoepassingen**.
2. Kies **Mijn invoegtoepassingen** → **Aangepaste invoegtoepassingen** → **Een aangepaste invoegtoepassing toevoegen** → **Toevoegen vanaf URL…**
3. Plak `https://rudbeyers.github.io/salto-mail-addin/manifest.xml` en bevestig.
4. Open een mail en klik op **Mail Agent** (in het lint, of onder ⋯ → Apps).
5. Vul bij ⚙ de sleutel in en klik **Bewaren en testen**.

Na een update aan de bestanden: herlaad het taakvenster (rechtsklik → Vernieuwen) of verhoog `?v=` in `taskpane.html`.
