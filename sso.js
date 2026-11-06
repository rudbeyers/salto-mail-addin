// Aanmelden met Microsoft via Nested App Authentication (MSAL). De server (app 30 v3.9) controleert het token en haalt
// de gebruiker daaruit. Lukt aanmelden niet, dan geeft ssoToken "" en valt de add-in terug op de gedeelde sleutel.
const SSO_CLIENT = "edf608f8-3a98-498f-af40-866eac8474ab";
const SSO_SCOPES = ["api://" + SSO_CLIENT + "/access_as_user"];
let _ssoPca = null;
let ssoFout = "";
let _ssoPopupMislukt = false;   // na één mislukt aanmeldvenster niet opnieuw proberen tot het paneel herladen wordt
async function ssoToken(interactief) {
  try {
    if (!window.msal) { ssoFout = "aanmeldbibliotheek niet geladen"; return ""; }
    if (!Office.context.requirements.isSetSupported("NestedAppAuth", "1.1")) { ssoFout = "deze Outlook ondersteunt het nog niet"; return ""; }
    if (!_ssoPca) {
      _ssoPca = await msal.createNestablePublicClientApplication({
        auth: { clientId: SSO_CLIENT, authority: "https://login.microsoftonline.com/common" } });
    }
    try {
      return (await _ssoPca.acquireTokenSilent({ scopes: SSO_SCOPES })).accessToken || "";
    } catch (e) {
      if (!interactief || _ssoPopupMislukt) { ssoFout = ssoFout || "eerst één keer aanmelden via het paneel"; return ""; }
      try {
        return (await _ssoPca.acquireTokenPopup({ scopes: SSO_SCOPES })).accessToken || "";
      } catch (e2) {
        _ssoPopupMislukt = true;
        throw e2;
      }
    }
  } catch (e) {
    ssoFout = (e && (e.errorCode || e.message)) || "onbekende fout";
    console.log("SSO:", e);
    return "";
  }
}
