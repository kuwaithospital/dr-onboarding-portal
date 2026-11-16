/* Kuwait Hospital – New Physician Onboarding: settings
   This is the ONLY file to edit when the flow addresses or IDs change.
   Paste each value between the quotes. Nothing here is secret: all checks run in Power Automate and Entra ID. */
window.KH_CONFIG = {

  /* Power Automate "When an HTTP request is received" URLs (copy from each flow's trigger after saving it) */
  flows: {
    requestCode: "",   // PO-2 Request code
    verifyCode:  "",   // PO-3 Verify code
    submit:      ""    // PO-4 Submit form
  },

  /* Largest file a doctor can upload, in MB */
  maxFileMB: 10,

  /* Staff page sign-in (Entra ID app registration "KH Physician Onboarding Portal") */
  msal: {
    clientId: "",      // Application (client) ID
    tenantId: "",      // Directory (tenant) ID
    redirectUri: "https://onboarding.kuwaithospital.com.kw/hr.html"
  },

  /* SharePoint site and lists */
  sharepoint: {
    host: "kuwaithospital.sharepoint.com",        // check against your SharePoint address
    sitePath: "/sites/PhysicianOnboarding",
    invitees: "Onboarding Invitees",
    submissions: "Physician Submissions"
  }
};
