/* Kuwait Hospital – New Physician Onboarding: settings
   This is the ONLY file to edit when the flow addresses or IDs change.
   Nothing here is secret: all checks run in Power Automate and Entra ID. */
window.KH_CONFIG = {

  /* Power Automate "When an HTTP request is received" URLs
     (open each flow → click the trigger → copy HTTP URL) */
  flows: {
    requestCode: "https://default7571e4919c534562a9b1e1b1091538.7e.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/18/workflows/412d033b02754d258ec06d2ddfbccd1d/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=ZRFzomWbkgz7_hB84wbVVSKebvG8-jToSLxtVAAbAl8",   // PO-2 Request code
    verifyCode:  "https://default7571e4919c534562a9b1e1b1091538.7e.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/30/workflows/e1619b2dff014affb7a836a89ddfb4a2/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=J8vKs86HjjFSWvKiOlw0rsgr0pc9jTk_CiUJ8eSDVzQ",   // PO-3 Verify code
    submit:      "https://default7571e4919c534562a9b1e1b1091538.7e.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/00/workflows/04549eef851e4341960e706a944b03d6/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=e4ONXgN2a8hIr-WnRQL_tLpnU7R2cdXyfkIl4zsNebQ"                            // PO-4 Submit form  (after PO-3)
  },

  /* Largest single file a doctor can upload, in MB */
  maxFileMB: 10,

  /* Largest total of all files in one submission, in MB
     (Power Automate accepts about 100 MB per request) */
  maxTotalMB: 70,

  /* Staff page sign-in (Entra ID app "KH Physician Onboarding Portal") */
  msal: {
    clientId: "93ed6021-44af-40cf-a751-76baa93a0701",
    tenantId: "7571e491-9c53-4562-a9b1-e1b10915387e",
    redirectUri: "https://kuwaithospital.github.io/dr-onboarding-portal/hr.html"
  },

  /* SharePoint site and lists */
  sharepoint: {
    host: "kwh.sharepoint.com",
    sitePath: "/sites/PhysicianOnboarding",
    invitees: "Onboarding Invitees",
    submissions: "Physician Submissions"
  }
};
