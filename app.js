/* Kuwait Hospital – New Physician Onboarding (doctor page)
   Talks only to the three Power Automate flows named in config.js. */
{
const CFG = window.KH_CONFIG || {};
const FLOWS = CFG.flows || {};
const MAX_MB = CFG.maxFileMB || 10;

const $ = s => document.querySelector(s), $$ = s => [...document.querySelectorAll(s)];
const esc = s => String(s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const wait = ms => new Promise(r => setTimeout(r, ms));
const busy = (b, on) => { b.disabled = on; b.classList.toggle('busy', on); };
const emailOK = e => /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(e);
const mmss = s => `${Math.floor(s / 60)}:${String(s % 60).padStart(2, '0')}`;
const showErr = (el, msg) => { el.textContent = msg; el.classList.add('show'); };
const hideErr = el => el.classList.remove('show');

$('#today').textContent = new Date().toLocaleDateString('en-GB', {day:'numeric', month:'long', year:'numeric'});

/* ---------- calls to Power Automate ---------- */
async function post(url, body) {
  if (!url) throw new Error('not-configured');
  const r = await fetch(url, {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(body)});
  let json = {};
  try { json = await r.json(); } catch (_) {}
  return {status: r.status, json};
}
const NET_MSG = 'We could not reach the server. Check your internet connection and try again.';
const CFG_MSG = 'The form is not configured yet. Please contact Kuwait Hospital HR.';

/* ---------- state ---------- */
let email = '', token = '', expTimer = null;

function showGate(id) {
  ['#gEmail', '#gCode', '#formWrap'].forEach(s => $(s).style.display = s === id ? '' : 'none');
  window.scrollTo({top: 0});
}

/* ---------- step 1: request a code ---------- */
const prefill = new URLSearchParams(location.search).get('email');
if (prefill && emailOK(prefill)) $('#em').value = prefill.trim().toLowerCase();

async function requestCode() {
  const res = await post(FLOWS.requestCode, {email});
  if (res.status >= 500) throw new Error('server');
}

function startTimers() {
  let left = 60, exp = 600;
  const rs = $('#resend');
  rs.disabled = true; rs.textContent = `Resend code in ${left}s`;
  $('#exp').textContent = 'Code expires in 10:00';
  clearInterval(expTimer);
  expTimer = setInterval(() => {
    left--; exp--;
    if (left <= 0) { rs.disabled = false; rs.textContent = 'Resend code'; } else rs.textContent = `Resend code in ${left}s`;
    $('#exp').textContent = exp > 0 ? `Code expires in ${mmss(exp)}` : 'Code expired. Request a new one.';
    if (exp <= 0 && left <= 0) clearInterval(expTimer);
  }, 1000);
}

$('#fEmail').addEventListener('submit', async e => {
  e.preventDefault();
  const em = $('#em').value.trim().toLowerCase(), er = $('#emErr'), b = $('#sendBtn');
  if (!emailOK(em)) { showErr(er, 'Please enter a valid email address.'); return; }
  hideErr(er); busy(b, true);
  email = em;
  try { await requestCode(); }
  catch (err) { busy(b, false); showErr(er, err.message === 'not-configured' ? CFG_MSG : NET_MSG); return; }
  busy(b, false);
  $('#codeMsg').innerHTML = `If <b style="color:var(--ink)">${esc(em)}</b> is registered for onboarding, a 6-digit code has been sent to it. Check your inbox and spam folder.`;
  clearOtp(); startTimers(); showGate('#gCode'); boxes[0].focus();
});

$('#resend').addEventListener('click', async () => {
  const er = $('#codeErr');
  try { await requestCode(); clearOtp(); startTimers(); showErr(er, 'A new code has been sent if this email is registered.'); er.classList.replace('err', 'info'); }
  catch (err) { er.classList.replace('info', 'err'); showErr(er, NET_MSG); }
});
$('#chg').addEventListener('click', () => { clearInterval(expTimer); showGate('#gEmail'); $('#em').focus(); });

/* ---------- step 2: verify the code ---------- */
const boxes = $$('#otp input');
function clearOtp() { boxes.forEach(x => x.value = ''); const er = $('#codeErr'); hideErr(er); er.classList.replace('info', 'err'); }
boxes.forEach((x, i) => {
  x.addEventListener('input', () => {
    x.value = x.value.replace(/\D/g, '').slice(-1);
    if (x.value && i < 5) boxes[i + 1].focus();
    if (boxes.every(b => b.value)) $('#fCode').requestSubmit();
  });
  x.addEventListener('keydown', e => { if (e.key === 'Backspace' && !x.value && i > 0) boxes[i - 1].focus(); });
  x.addEventListener('paste', e => {
    const t = (e.clipboardData.getData('text') || '').replace(/\D/g, '').slice(0, 6);
    if (!t.length) return;
    e.preventDefault();
    t.split('').forEach((c, j) => boxes[j] && (boxes[j].value = c));
    boxes[Math.min(t.length, 5)].focus();
    if (t.length === 6) $('#fCode').requestSubmit();
  });
});

$('#fCode').addEventListener('submit', async e => {
  e.preventDefault();
  const code = boxes.map(b => b.value).join(''), er = $('#codeErr'), b = $('#verBtn');
  er.classList.replace('info', 'err');
  if (code.length < 6) { showErr(er, 'Enter all 6 digits.'); return; }
  busy(b, true);
  let res;
  try { res = await post(FLOWS.verifyCode, {email, code}); }
  catch (err) { busy(b, false); showErr(er, err.message === 'not-configured' ? CFG_MSG : NET_MSG); return; }
  busy(b, false);
  const j = res.json || {};
  if (res.status === 200 && j.token) {
    token = j.token; clearInterval(expTimer);
    $('#sessEm').textContent = email;
    const who = $('#sbWho'); if (who) who.textContent = 'Verified as ' + email;
    if (j.name && !$('#nameEn').value) $('#nameEn').value = j.name;
    update(); showGate('#formWrap');
    return;
  }
  boxes.forEach(x => x.value = ''); boxes[0].focus();
  if (res.status === 429 || j.error === 'locked') showErr(er, 'Too many attempts. Please request a new code.');
  else if (j.error === 'expired') showErr(er, 'This code is invalid or has expired. Please request a new code.');
  else if (typeof j.attemptsLeft === 'number') showErr(er, `Incorrect code. ${j.attemptsLeft} attempt${j.attemptsLeft === 1 ? '' : 's'} left.`);
  else showErr(er, 'Incorrect code. Please try again.');
});

/* ---------- step 3: the form ---------- */
const f = $('#f');
const store = {licence:[], goodStanding:[], education:[], certificates:[], experience:[], resume:[], patientContacts:[]};
const extOK = {word:/\.(docx?|rtf)$/i, pdf:/\.pdf$/i, image:/\.(png|jpe?g|gif|bmp|webp|heic|tiff?)$/i};
const allowed = (name, types) => types === 'any' || types.split(',').some(t => extOK[t].test(name));
const size = n => n > 1048576 ? (n / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB';

function renderFiles(q) {
  const k = q.dataset.file, ul = q.querySelector('.files');
  ul.innerHTML = store[k].map((x, i) => `<li><span>${esc(x.name)}</span><small>${size(x.size)}</small><button type="button" data-i="${i}" aria-label="Remove ${esc(x.name)}">×</button></li>`).join('');
}
function addFiles(q, list) {
  const k = q.dataset.file, max = +q.dataset.max, types = q.dataset.types, msg = q.querySelector('.msg');
  let err = '';
  for (const file of list) {
    if (!allowed(file.name, types)) { err = `"${file.name}" is not an allowed file type.`; continue; }
    if (file.size > MAX_MB * 1048576) { err = `"${file.name}" is larger than ${MAX_MB} MB.`; continue; }
    if (max === 1) store[k] = [file];
    else if (store[k].length >= max) { err = `You can upload up to ${max} files.`; break; }
    else store[k].push(file);
  }
  renderFiles(q);
  if (err) { msg.textContent = err; q.classList.add('bad'); }
  else { q.classList.remove('bad'); if (q.dataset.msg) msg.textContent = q.dataset.msg; }
  update();
}
$$('[data-file]').forEach(q => {
  const inp = q.querySelector('input[type=file]'), drop = q.querySelector('.drop');
  inp.addEventListener('change', () => { addFiles(q, inp.files); inp.value = ''; });
  ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, () => drop.classList.remove('over')));
  drop.addEventListener('drop', e => { e.preventDefault(); addFiles(q, e.dataTransfer.files); });
  q.querySelector('.files').addEventListener('click', e => {
    const b = e.target.closest('button'); if (!b) return;
    store[q.dataset.file].splice(+b.dataset.i, 1); renderFiles(q); update();
  });
});

function filled(q) {
  if (q.dataset.radio !== undefined) return !!f.querySelector(`input[name=${q.dataset.radio}]:checked`);
  if (q.dataset.file) return store[q.dataset.file].length > 0;
  if (q.dataset.three !== undefined) return [...q.querySelectorAll('input')].every(i => i.value.trim());
  const el = q.querySelector('input,textarea');
  if (q.dataset.num !== undefined) return el.value.trim() !== '' && !isNaN(+el.value) && +el.value >= 0;
  return el.value.trim() !== '';
}
const reqs = $$('.q[data-req]');
function update() {
  const n = reqs.filter(filled).length;
  $('#pct').textContent = `${n} of ${reqs.length}`;
  $('#pbar').style.width = (n / reqs.length * 100) + '%';
}
f.addEventListener('input', e => { const q = e.target.closest('.q'); if (q && filled(q) && !q.dataset.file) q.classList.remove('bad'); update(); });
f.addEventListener('change', e => { const q = e.target.closest('.q'); if (q && q.dataset.radio && filled(q)) q.classList.remove('bad'); update(); });
update();

const toB64 = file => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(String(r.result).split(',')[1]); r.onerror = rej; r.readAsDataURL(file); });
const pack = async k => Promise.all(store[k].map(async x => ({name:x.name, type:x.type || 'application/octet-stream', size:x.size, content: await toB64(x)})));

f.addEventListener('submit', async e => {
  e.preventDefault();
  const er = $('#subErr'); hideErr(er);
  const bad = reqs.filter(q => !filled(q)).concat($$('.q[data-pattern]').filter(q => { const x = q.querySelector('input').value.trim(); return x && !new RegExp(q.dataset.pattern).test(x); }));
  reqs.forEach(q => q.classList.toggle('bad', bad.includes(q)));
  $$('.q[data-pattern]').forEach(q => { if (bad.includes(q)) q.classList.add('bad'); });
  if (bad.length) { bad[0].scrollIntoView({behavior:'smooth', block:'center'}); const el = bad[0].querySelector('input,textarea'); el && el.focus({preventScroll:true}); return; }
  const v = n => (f.elements[n].value || '').trim();
  const data = {
    fullNameEn: v('fullNameEn'), fullNameAr: v('fullNameAr'), specialty: v('specialty'),
    designation: f.querySelector('input[name=designation]:checked').value,
    yearsExperience: +v('yearsExperience'), qualifications: v('qualifications'),
    keyServices: v('keyServices'), symptoms: v('symptoms'), minPatientAge: +v('minPatientAge'),
    civilId: v('civilId'), passportNumber: v('passportNumber').toUpperCase(),
    languages: v('languages'), socialMedia: v('socialMedia'),
    top3Services: [v('top1'), v('top2'), v('top3')]
  };
  const totalMB = Object.values(store).flat().reduce((t, x) => t + x.size, 0) / 1048576, MAX_TOTAL = CFG.maxTotalMB || 70;
  if (totalMB > MAX_TOTAL) { showErr(er, `Your files add up to ${totalMB.toFixed(0)} MB; the limit for one submission is ${MAX_TOTAL} MB. Please use smaller files, for example compressed PDFs.`); return; }
  const btn = $('#go'); busy(btn, true); $('#goT').textContent = 'Uploading, please wait…';
  let res;
  try {
    const body = {token, ...data, licence: await pack('licence'), goodStanding: await pack('goodStanding'), education: await pack('education'), certificates: await pack('certificates'), experience: await pack('experience'), resume: await pack('resume'), patientContacts: await pack('patientContacts')};
    res = await post(FLOWS.submit, body);
  } catch (err) {
    busy(btn, false); $('#goT').textContent = 'Submit onboarding form';
    showErr(er, err.message === 'not-configured' ? CFG_MSG : NET_MSG + ' Your answers are still here.');
    return;
  }
  busy(btn, false); $('#goT').textContent = 'Submit onboarding form';
  const j = res.json || {};
  if (res.status === 200 && j.ok !== false) { showDone(data, j.reference || ''); return; }
  if (res.status === 401) { showErr(er, 'Your verification has ended. Open the link again and verify your email with a new code. Your answers on this page will be lost if you close it.'); return; }
  if (res.status === 409) { showErr(er, 'A form has already been submitted for this email. Please contact Kuwait Hospital HR if you need to change it.'); return; }
  if (res.status === 413) { showErr(er, 'The files are too large to send together. Please use smaller files (for example, compressed PDFs) and try again.'); return; }
  showErr(er, 'Something went wrong while saving your form. Please try again in a few minutes. Your answers are still here.');
});

function showDone(d, ref) {
  const fl = k => store[k].length ? store[k].map(x => x.name).join('\n') : '—';
  const rows = [['Name (English)', d.fullNameEn], ['Name (Arabic)', d.fullNameAr], ['Specialty', d.specialty], ['Designation', d.designation],
    ['Civil ID number', d.civilId || '—'], ['Passport number', d.passportNumber],
    ['Years of experience', d.yearsExperience], ['Minimum patient age', d.minPatientAge === 0 ? 'All ages' : d.minPatientAge + ' years'],
    ['Languages', d.languages], ['Top 3 services', d.top3Services.join(' · ')],
    ['Medical licence', fl('licence')], ['Good Standing', fl('goodStanding')], 
    ['Educational documents', fl('education')], ['Certificates', fl('certificates')], ['Experience letters', fl('experience')], ['Resume / CV', fl('resume')], ['Patient contacts', fl('patientContacts')]];
  $('#sum').innerHTML = rows.map(([a, b]) => `<div><dt>${a}</dt><dd ${a.includes('Arabic') ? 'lang="ar" dir="rtl" style="text-align:left"' : ''}>${esc(b)}</dd></div>`).join('');
  $('#ref').textContent = ref ? 'Reference ' + ref : '';
  $('#ref').style.display = ref ? '' : 'none';
  $('#tyH').textContent = 'Thank you, ' + d.fullNameEn;
  $('#tyMail').textContent = email;
  $('#sess').style.display = 'none';
  const who = $('#sbWho'); if (who) who.textContent = '';
  f.style.display = 'none'; $('.intro').style.display = 'none'; $('#done').style.display = 'block';
  window.scrollTo({top: 0, behavior: 'smooth'});
}

/* ---------- sticky title bar ---------- */
const bar = $('#sbar'), cov = $('header.cover');
const onScroll = () => { const show = cov.getBoundingClientRect().bottom < 70; bar.classList.toggle('show', show); bar.setAttribute('aria-hidden', show ? 'false' : 'true'); document.body.classList.toggle('sb-on', show); };
window.addEventListener('scroll', onScroll, {passive:true}); window.addEventListener('resize', onScroll); onScroll();
}
