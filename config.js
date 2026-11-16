/* Kuwait Hospital – New Physician Onboarding: settings
   This is the ONLY file to edit when the flow addresses or IDs change.
   Nothing here is secret: all checks run in Power Automate and Entra ID. */
window.KH_CONFIG = {

  /* Power Automate "When an HTTP request is received" URLs (filled in later) */
  flows: {
    requestCode: "",   // PO-2 Request code
    verifyCode:  "",   // PO-3 Verify code
    submit:      ""    // PO-4 Submit form
  },

  /* Largest file a doctor can upload, in MB */
  maxFileMB: 10,

  /* Staff page sign-in (Entra ID app "KH Physician Onboarding Portal") */
  msal: {
    clientId: "93ed6021-44af-40cf-a751-76baa93a0701",
    tenantId: "7571e491-9c53-4562-a9b1-e1b10915387e",
    redirectUri: "https://kuwaithospital.github.io/dr-onboarding-portal/hr.html"
  },

  /* SharePoint site and lists */
  sharepoint: {
    host: "kuwaithospital.sharepoint.com",        // check against your SharePoint address
    sitePath: "/sites/PhysicianOnboarding",
    invitees: "Onboarding Invitees",
    submissions: "Physician Submissions"
  }
};
