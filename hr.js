/* Kuwait Hospital – New Physician Onboarding (staff page)
   Sign-in: Microsoft Entra ID (only users assigned to the app can sign in).
   Data: SharePoint list "Onboarding Invitees" through Microsoft Graph, as the signed-in user. */
(() => {
const CFG = window.KH_CONFIG || {};
const M = CFG.msal || {}, SP = CFG.sharepoint || {};
const SCOPES = ['User.Read', 'Sites.ReadWrite.All'];
const GRAPH = 'https://graph.microsoft.com/v1.0';

const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const busy = (b, on) => { b.disabled = on; b.classList.toggle('busy', on); };
const showErr = (el, msg) => { el.innerHTML = msg; el.classList.add('show'); };
const hide = el => el.classList.remove('show');
const emailOK = e => /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(e);
const ini = n => String(n || '?').split(/\s+/).filter(Boolean).slice(0, 2).map(w => w[0]).join('').toUpperCase();
const fmtD = s => s ? new Date(s).toLocaleDateString('en-GB', {day:'numeric', month:'short', year:'numeric'}) : '';

$('#today').textContent = new Date().toLocaleDateString('en-GB', {day:'numeric', month:'long', year:'numeric'});
$('#spLink').href = `https://${SP.host}${SP.sitePath}`;

if (!M.clientId || !M.tenantId || !window.msal) {
  showErr($('#sErr'), 'Sign-in is not configured yet. IT: fill in clientId and tenantId in config.js.');
  $('#msBtn').disabled = true;
  return;
}

const app = new msal.PublicClientApplication({
  auth: {clientId: M.clientId, authority: `https://login.microsoftonline.com/${M.tenantId}`, redirectUri: M.redirectUri || location.href.split('#')[0].split('?')[0]},
  cache: {cacheLocation: 'sessionStorage'}
});

let account = null, siteId = '', listId = '', items = [], role = '';

async function token() {
  try { return (await app.acquireTokenSilent({scopes: SCOPES, account})).accessToken; }
  catch (e) { await app.acquireTokenRedirect({scopes: SCOPES, account}); throw e; }
}
async function graph(path, opts = {}) {
  const t = await token();
  const r = await fetch(GRAPH + path, {...opts, headers: {Authorization: 'Bearer ' + t, 'Content-Type': 'application/json', ...(opts.headers || {})}});
  if (!r.ok) { let m = ''; try { m = (await r.json()).error.message; } catch (_) {} const e = new Error(m || ('HTTP ' + r.status)); e.status = r.status; throw e; }
  return r.status === 204 ? {} : r.json();
}

function roleFrom(acc) {
  const r = (acc.idTokenClaims && acc.idTokenClaims.roles) || [];
  if (r.includes('Portal.Admin')) return 'Admin';
  if (r.includes('Portal.HR')) return 'HR';
  if (r.includes('Portal.Viewer')) return 'Viewer';
  return '';
}

async function start() {
  const res = await app.handleRedirectPromise().catch(e => { showErr($('#sErr'), friendly(e)); return null; });
  account = (res && res.account) || app.getAllAccounts()[0] || null;
  if (!account) return;
  role = roleFrom(account);
  if (!role) { showErr($('#sErr'), 'You are signed in, but your account has no role on this page. Please ask IT to assign you the HR or Viewer role.'); return; }
  $('#sLogin').style.display = 'none'; $('#sApp').style.display = '';
  const name = account.name || account.username;
  $('#meAv').textContent = ini(name); $('#meN').textContent = name; $('#meE').textContent = account.username; $('#meR').textContent = role;
  $('#entraLink').style.display = role === 'Admin' ? '' : 'none';
  $('#fInv').style.display = role === 'Viewer' ? 'none' : '';
  await load();
}

function friendly(e) {
  const m = String(e && (e.errorMessage || e.message) || e);
  if (/AADSTS50105/.test(m)) return 'Your account has not been given access to this page. Please contact the IT administrator.';
  if (/AADSTS50020|AADSTS50034/.test(m)) return 'Only Kuwait Hospital (@kuwaithospital.com.kw) accounts can sign in to this page.';
  return 'Sign-in failed. Please try again or contact IT. (' + esc(m.split('\n')[0].slice(0, 160)) + ')';
}

$('#msBtn').addEventListener('click', () => { busy($('#msBtn'), true); app.loginRedirect({scopes: SCOPES, prompt: 'select_account'}); });
$('#outBtn').addEventListener('click', () => app.logoutRedirect({account}));

/* ---------- SharePoint ---------- */
async function ensureIds() {
  if (!siteId) siteId = (await graph(`/sites/${SP.host}:${SP.sitePath}`)).id;
  if (!listId) {
    const r = await graph(`/sites/${siteId}/lists?$filter=displayName eq '${SP.invitees.replace(/'/g, "''")}'&$select=id,displayName`);
    if (!r.value || !r.value.length) throw new Error(`List "${SP.invitees}" not found on the site.`);
    listId = r.value[0].id;
  }
}

async function load() {
  const L = $('#loading'); L.style.display = ''; L.textContent = 'Loading…';
  try {
    await ensureIds();
    items = [];
    let url = `/sites/${siteId}/lists/${listId}/items?$expand=fields($select=Title,Email,Status,InviteCount,LastInviteSent,Reference)&$top=500`;
    while (url) {
      const r = await graph(url);
      items.push(...(r.value || []));
      url = r['@odata.nextLink'] ? r['@odata.nextLink'].replace(GRAPH, '') : '';
    }
    items.sort((a, b) => new Date(b.createdDateTime) - new Date(a.createdDateTime));
    L.style.display = 'none';
    render();
  } catch (e) {
    L.innerHTML = e.status === 403 ? 'You signed in, but you don\'t have permission to the SharePoint list. Please ask IT to add you to the PO HR or PO Viewers group.' : 'Could not load the list: ' + esc(e.message);
  }
}

function render() {
  const q = $('#search').value.trim().toLowerCase();
  const rows = items.filter(i => !q || (i.fields.Title || '').toLowerCase().includes(q) || (i.fields.Email || '').toLowerCase().includes(q));
  $('#count').textContent = `${rows.length} of ${items.length} doctors`;
  $('#invRows').innerHTML = rows.map(i => {
    const f = i.fields, st = f.Status || 'Invited';
    const act = (st === 'Submitted' || role === 'Viewer') ? '' : `<button class="rm" style="color:var(--accent)" type="button" data-id="${esc(i.id)}" aria-label="Resend invite to ${esc(f.Email)}">Resend</button>`;
    return `<tr><td data-l="Doctor">${esc(f.Title)}</td><td data-l="Email">${esc(f.Email)}</td><td data-l="Registered">${fmtD(i.createdDateTime)}</td><td data-l="Status"><span class="st ${esc(st)}">${esc(st)}</span>${f.Reference ? `<br><small style="color:var(--muted)">${esc(f.Reference)}</small>` : ''}</td><td data-l="">${act}</td></tr>`;
  }).join('') || `<tr><td colspan="5" style="color:var(--muted)">${items.length ? 'No match.' : 'No doctors registered yet.'}</td></tr>`;
}
$('#search').addEventListener('input', render);
$('#refresh').addEventListener('click', load);

$('#fInv').addEventListener('submit', async e => {
  e.preventDefault();
  const n = $('#invN').value.trim(), em = $('#invE').value.trim().toLowerCase(), ok = $('#invOk'), er = $('#invErr'), b = $('#invBtn');
  hide(ok); hide(er);
  if (!n) { showErr(er, 'Enter the doctor\'s name.'); $('#invN').focus(); return; }
  if (!emailOK(em)) { showErr(er, 'Enter a valid email address.'); $('#invE').focus(); return; }
  if (items.some(i => (i.fields.Email || '').toLowerCase() === em)) { showErr(er, 'This email is already registered. Use Resend if the doctor needs the invitation again.'); return; }
  busy(b, true);
  try {
    await graph(`/sites/${siteId}/lists/${listId}/items`, {method: 'POST', body: JSON.stringify({fields: {Title: n, Email: em}})});
    $('#invN').value = ''; $('#invE').value = '';
    showErr(ok, `Saved. An invitation with the form link and QR code will be emailed to <b>${esc(em)}</b> within a minute.`);
    await load();
  } catch (x) { showErr(er, 'Could not save: ' + esc(x.message)); }
  busy(b, false);
});

$('#invRows').addEventListener('click', async e => {
  const b = e.target.closest('[data-id]'); if (!b) return;
  const it = items.find(i => i.id === b.dataset.id); if (!it) return;
  b.disabled = true; b.textContent = 'Sending…';
  try {
    await graph(`/sites/${siteId}/lists/${listId}/items/${it.id}/fields`, {method: 'PATCH', body: JSON.stringify({ResendInvite: true})});
    showErr($('#invOk'), `Invitation will be sent again to <b>${esc(it.fields.Email)}</b> within a minute.`);
    b.textContent = 'Sent';
  } catch (x) { showErr($('#invErr'), 'Could not resend: ' + esc(x.message)); b.disabled = false; b.textContent = 'Resend'; }
});

start();
})();
