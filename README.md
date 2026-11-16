# Kuwait Hospital – New Physician Onboarding Portal

Static web pages for the New Physician Onboarding process. **No data is stored here.**
All answers and documents go to Kuwait Hospital SharePoint through Power Automate.

| File | Purpose |
| --- | --- |
| `index.html`, `app.js` | Doctor page: email → 6-digit code → form → thank-you |
| `hr.html`, `hr.js` | Staff page: Microsoft 365 sign-in, register doctors, resend invitations |
| `config.js` | **The only file to edit**: flow URLs, Entra IDs, SharePoint names |
| `style.css` | Shared styles (Kuwait Hospital navy, Fira Sans / Tajawal) |
| `logo-white.png`, `logo-navy.png`, `qr.png` | Images used by the pages and the emails |
| `vendor/msal-browser.min.js` | Microsoft sign-in library (MSAL.js 2.39.0, MIT licence) |
| `CNAME` | Custom address `onboarding.kuwaithospital.com.kw` |

## Setup
1. Build the SharePoint site, Entra app and five flows (see the Implementation Guide).
2. Paste the three flow URLs, the client ID and the tenant ID into `config.js`.
3. Settings → Pages → Deploy from branch `main`, folder `/ (root)`; custom domain `onboarding.kuwaithospital.com.kw`; Enforce HTTPS.
4. DNS: CNAME `onboarding` → `kuwaithospital.github.io`.

Doctor link: `https://onboarding.kuwaithospital.com.kw/` (add `?email=name@example.com` to pre-fill the email)
Staff link: `https://onboarding.kuwaithospital.com.kw/hr.html`

Prepared and developed by Mohamed Sirajudeen, IT – Kuwait Hospital.
