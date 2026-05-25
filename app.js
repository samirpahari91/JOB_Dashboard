/* =========================================================
   FitForge — Job-Match & ATS Resume Studio
   Pure client-side. No frameworks, no servers.
   ========================================================= */
(() => {
'use strict';

// =========================================================
// 1. EXTERNAL LIBRARY BOOTSTRAP (race-safe)
// =========================================================
// Set pdf.js worker as soon as it's available, not on window 'load'
// (a fast user click could fire before 'load' on a cached page).
function ensurePdfWorker() {
  const lib = window.pdfjsLib;
  if (lib && lib.GlobalWorkerOptions && !lib.GlobalWorkerOptions.workerSrc) {
    lib.GlobalWorkerOptions.workerSrc =
      'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  }
}
ensurePdfWorker();
window.addEventListener('load', ensurePdfWorker);

// Wait for a global (PDF/DOCX parsers may still be loading from CDN)
function waitForGlobal(name, timeoutMs = 8000) {
  return new Promise((resolve, reject) => {
    if (window[name]) return resolve(window[name]);
    const t0 = Date.now();
    const id = setInterval(() => {
      if (window[name]) { clearInterval(id); resolve(window[name]); }
      else if (Date.now() - t0 > timeoutMs) {
        clearInterval(id);
        reject(new Error(`${name} failed to load. Check your internet connection or CDN access.`));
      }
    }, 80);
  });
}

// =========================================================
// 2. STATE
// =========================================================
const state = {
  cv: { name: '', text: '' },
  transcript: { name: '', text: '' },
  jd: '',
  analysis: null,
  resume: '',
  cover: '',
  lastAnalysisHash: '',
};

// =========================================================
// 3. DOM HELPERS
// =========================================================
const $  = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
const el = (tag, props = {}, ...kids) => {
  const n = Object.assign(document.createElement(tag), props);
  for (const k of kids) {
    if (k == null || k === false) continue;
    n.append(k && k.nodeType ? k : document.createTextNode(String(k)));
  }
  return n;
};

// =========================================================
// 4. TOAST + LOADER
// =========================================================
let toastTimer;
function toast(msg, err = false) {
  const t = $('#toast');
  t.textContent = msg;
  t.classList.toggle('err', err);
  t.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => (t.hidden = true), 3400);
}
let loaderTimer;
function showLoader(text = 'Working…') {
  $('#loaderText').textContent = text;
  $('#loader').hidden = false;
  // Safety: never trap the user forever
  clearTimeout(loaderTimer);
  loaderTimer = setTimeout(() => {
    if (!$('#loader').hidden) { hideLoader(); toast('That took too long — please try again.', true); }
  }, 25000);
}
function hideLoader() { clearTimeout(loaderTimer); $('#loader').hidden = true; }

// =========================================================
// 5. THEME
// =========================================================
const themeBtn = $('#themeBtn');
const savedTheme = localStorage.getItem('ff-theme');
if (savedTheme === 'light' || savedTheme === 'dark') document.body.dataset.theme = savedTheme;
themeBtn.addEventListener('click', () => {
  const next = document.body.dataset.theme === 'dark' ? 'light' : 'dark';
  document.body.dataset.theme = next;
  try { localStorage.setItem('ff-theme', next); } catch (_) {}
});

// =========================================================
// 6. FILE PARSING
// =========================================================
const MAX_FILE_BYTES = 10 * 1024 * 1024; // 10 MB sanity cap

async function parseFile(file) {
  if (file.size > MAX_FILE_BYTES) throw new Error('File is over 10 MB. Try exporting a smaller PDF or TXT.');
  const name = (file.name || '').toLowerCase();
  if (name.endsWith('.txt'))  return (await file.text()).trim();
  if (name.endsWith('.docx')) return (await parseDocx(file)).trim();
  if (name.endsWith('.pdf'))  return (await parsePdf(file)).trim();
  // Try by MIME as a last resort
  if (file.type === 'text/plain') return (await file.text()).trim();
  throw new Error('Unsupported file type. Please upload a PDF, DOCX, or TXT file.');
}

async function parseDocx(file) {
  const mammoth = await waitForGlobal('mammoth');
  const buf = await file.arrayBuffer();
  const { value } = await mammoth.extractRawText({ arrayBuffer: buf });
  return value || '';
}

async function parsePdf(file) {
  const pdfjsLib = await waitForGlobal('pdfjsLib');
  ensurePdfWorker();
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf, isEvalSupported: false }).promise;
  let text = '';
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    text += content.items.map(it => it.str).join(' ') + '\n\n';
  }
  return text;
}

// =========================================================
// 7. UPLOAD WIRING
// =========================================================
function wireUpload({ dropId, inputId, statusId, target, label }) {
  const drop   = $('#' + dropId);
  const input  = $('#' + inputId);
  const status = $('#' + statusId);

  const openPicker = () => input.click();
  drop.addEventListener('click', openPicker);
  drop.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openPicker(); }
  });

  ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => {
    e.preventDefault(); drop.classList.add('dragging');
  }));
  ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => {
    e.preventDefault(); drop.classList.remove('dragging');
  }));
  drop.addEventListener('drop', e => {
    const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) handleFile(f);
  });
  input.addEventListener('change', e => {
    const f = e.target.files && e.target.files[0];
    if (f) handleFile(f);
  });

  async function handleFile(file) {
    try {
      showLoader(`Reading ${label}…`);
      const text = await parseFile(file);
      if (!text || text.length < 5) {
        throw new Error('Could not extract text from this file. If it is a scanned PDF (text-as-images), please paste the text into a TXT file.');
      }
      state[target] = { name: file.name, text };
      status.hidden = false;
      status.classList.remove('err');
      status.innerHTML = '';
      status.append(
        svgCheck(),
        el('span', {}, el('strong', {}, file.name), ` · ${formatBytes(file.size)} · ${wordCount(text)} words`),
        (() => {
          const b = el('button', { className: 'link-btn' }, 'Remove');
          b.addEventListener('click', () => clearUpload(target, status));
          return b;
        })()
      );
      updateStepper(); updateAnalyzeButton();
    } catch (err) {
      console.error(err);
      status.hidden = false;
      status.classList.add('err');
      status.innerHTML = '';
      status.append(el('span', {}, '⚠️ ' + (err && err.message ? err.message : 'Failed to read file.')));
    } finally {
      hideLoader();
      input.value = '';
    }
  }
}
function svgCheck() {
  const ns = 'http://www.w3.org/2000/svg';
  const svg = document.createElementNS(ns, 'svg');
  svg.setAttribute('viewBox', '0 0 24 24'); svg.setAttribute('width', '16'); svg.setAttribute('height', '16');
  svg.setAttribute('fill', 'none'); svg.setAttribute('stroke', 'currentColor'); svg.setAttribute('stroke-width', '2.4');
  const p = document.createElementNS(ns, 'path'); p.setAttribute('d', 'M20 6 9 17l-5-5');
  svg.appendChild(p); return svg;
}

function clearUpload(target, status) {
  state[target] = { name: '', text: '' };
  status.hidden = true; status.innerHTML = '';
  updateStepper(); updateAnalyzeButton();
}

wireUpload({ dropId: 'cvDrop', inputId: 'cvFile', statusId: 'cvStatus', target: 'cv',         label: 'CV' });
wireUpload({ dropId: 'trDrop', inputId: 'trFile', statusId: 'trStatus', target: 'transcript', label: 'transcript' });

// JD textarea
const jdTextEl = $('#jdText');
jdTextEl.addEventListener('input', () => {
  state.jd = jdTextEl.value;
  $('#jdCount').textContent = `${wordCount(state.jd)} words`;
  updateStepper(); updateAnalyzeButton();
});
$('#clearJd').addEventListener('click', () => {
  jdTextEl.value = ''; state.jd = '';
  $('#jdCount').textContent = '0 words';
  updateStepper(); updateAnalyzeButton();
});

// =========================================================
// 8. STEPPER + BUTTON STATE
// =========================================================
function updateStepper() {
  const steps = $$('.stepper .step');
  const done = {
    1: !!state.cv.text,
    2: !!state.transcript.text,
    3: state.jd.trim().length > 30,
    4: !!state.analysis,
    5: !!state.resume || !!state.cover,
  };
  // First unfinished step is "active". Step 2 is optional, so completing 1 → active = 3.
  const order = [1, 2, 3, 4, 5];
  let firstUndone = -1;
  for (const k of order) {
    if (k === 2) continue; // skip optional step in active calc
    if (!done[k]) { firstUndone = k; break; }
  }
  steps.forEach(s => {
    const k = +s.dataset.step;
    s.classList.toggle('done', !!done[k]);
    s.classList.toggle('active', k === firstUndone);
  });
}
function updateAnalyzeButton() {
  $('#analyzeBtn').disabled = !(state.cv.text && state.jd.trim().length > 30);
}

// =========================================================
// 9. NLP UTILITIES
// =========================================================
const STOP = new Set(`a an and are as at be been being but by can could did do does for from had has have having he her him his i if in into is it its itself me my of on one or our ours she so some such than that the their them then there these they this those to was we were what when where which while who whom why will with would you your yours about above after again against all am any because before below between both during each few further here how just more most no nor not now off only other out over own same should some too under until up very what whats whatever via per upon also will would should may might must shall ours ourselves themselves you'll you've they're we're it's that's there's here's well even still ever already mr ms mrs role roles team teams company companies job jobs work works`.split(/\s+/));

// Skill vocabulary used for skill-match signal. Each entry is { canonical, regex }.
const SKILL_LIST = [
  // Languages
  'python','java','javascript','typescript','c++','c#','golang','go','rust','ruby','php','scala','kotlin','swift','r','matlab','perl','bash','shell','sql','nosql','html','css','sass',
  // Frameworks / web
  'react','angular','vue','svelte','next.js','nuxt','node.js','node','express','django','flask','fastapi','spring','rails','laravel','.net','asp.net',
  // Data / ML
  'pandas','numpy','scikit-learn','sklearn','tensorflow','pytorch','keras','xgboost','llm','nlp','machine learning','deep learning','computer vision','statistics','a/b testing','time series','etl','elt','data warehouse','dbt','airflow','spark','hadoop','kafka',
  // DB
  'mongodb','postgres','postgresql','mysql','redis','sqlite','snowflake','bigquery','redshift','databricks','oracle','dynamodb','elasticsearch',
  // Cloud / devops
  'aws','azure','gcp','google cloud','kubernetes','docker','terraform','ansible','jenkins','gitlab','github actions','ci/cd','linux','rest','graphql','grpc','microservices',
  // BI
  'tableau','power bi','looker','excel','vba','jupyter','sas','spss',
  // GIS
  'gis','arcgis','arcgis pro','arcgis online','arcgis enterprise','arcpy','qgis','postgis','geoserver','mapbox','leaflet','openlayers','spatial analysis','remote sensing','geodatabase','geospatial','cartography',
  // PM / soft
  'agile','scrum','kanban','jira','confluence','stakeholder management','project management','communication','leadership','mentoring','collaboration','problem solving','critical thinking',
  // QA / sec
  'unit testing','integration testing','playwright','cypress','selenium','jest','mocha','pytest','penetration testing','owasp','compliance','hipaa',
  // Marketing / biz
  'seo','sem','google analytics','crm','salesforce','hubspot','copywriting','content strategy','product management','market research','financial modeling','accounting','quickbooks',
  // Health
  'clinical research','epidemiology','biostatistics','phlebotomy','ehr','epic','cerner',
];

// Precompile skill regexes once. Use \b boundaries that work in Safari (no lookbehind).
const SKILL_REGEX = SKILL_LIST.map(s => ({
  name: s,
  // Anchor with non-letter-digit boundaries on both sides, handled via manual char class.
  re: new RegExp(`(^|[^a-z0-9+#.])${escapeRegex(s)}(?=$|[^a-z0-9+#.])`, 'i'),
}));

function tokens(text) {
  return (text || '')
    .toLowerCase()
    .replace(/[^a-z0-9+#./\-\s]/g, ' ')
    .split(/\s+/)
    .filter(w => w && !STOP.has(w) && w.length > 1 && !/^\d+$/.test(w));
}

// n-grams (1..3) with stop-word filtering at every position
function termFreq(text, maxNgram = 3) {
  const toks = tokens(text);
  const counts = new Map();
  for (let n = 1; n <= maxNgram; n++) {
    for (let i = 0; i + n <= toks.length; i++) {
      const slice = toks.slice(i, i + n);
      if (slice.some(t => STOP.has(t))) continue;
      // skip if all tokens are tiny numbers
      if (slice.every(t => /^\d+$/.test(t))) continue;
      const gram = slice.join(' ');
      if (gram.length < 3) continue;
      counts.set(gram, (counts.get(gram) || 0) + 1);
    }
  }
  return counts;
}

function extractSkills(text) {
  const out = new Set();
  if (!text) return out;
  for (const { name, re } of SKILL_REGEX) {
    if (re.test(text)) out.add(name);
  }
  return out;
}

// Match a phrase as a whole "word" (Safari-safe; no lookbehind)
function phrasePresent(text, phrase) {
  if (!text || !phrase) return false;
  const re = new RegExp(`(^|[^a-z0-9+#.])${escapeRegex(phrase)}(?=$|[^a-z0-9+#.])`, 'i');
  return re.test(text);
}

function topJdKeywords(jd, limit = 30) {
  const tf = termFreq(jd, 3);
  const arr = [];
  for (const [term, count] of tf) {
    const ngram = term.split(' ').length;
    // Penalize 1-grams that are extremely common English nouns we forgot to stop
    const score = count * (1 + (ngram - 1) * 0.7);
    arr.push({ term, count, score });
  }
  arr.sort((a, b) => b.score - a.score);
  return arr.slice(0, limit);
}

// =========================================================
// 10. QUALIFICATIONS / TITLE / EXPERIENCE EXTRACTORS
// =========================================================
function extractRequiredQuals(jd) {
  const quals = [];
  const lines = jd.split(/\r?\n|•|·/).map(l => l.trim()).filter(Boolean);
  const reqHints = /(required|must have|minimum|at least|you have|we require|qualifications)/i;

  // years experience
  const yrRe = /(\d{1,2})\+?\s*(?:to\s*\d{1,2}\s*)?(?:years?|yrs?)\s+(?:of\s+)?([a-z][a-z\s/+#.\-]{2,60})?/ig;
  let m;
  while ((m = yrRe.exec(jd)) !== null) {
    const yrs = +m[1];
    if (yrs > 40) continue; // sanity
    const ctx = (m[2] || 'relevant experience').trim().replace(/\s+/g, ' ').slice(0, 60);
    quals.push({ type: 'experience', text: `${yrs}+ years of ${ctx}`, years: yrs, key: `exp:${yrs}` });
  }
  // degrees
  const degRe = /\b(bachelor|bachelor['']s|master|master['']s|ph\.?d|doctorate|associate)s?\s+(?:degree\s+)?(?:in\s+([a-z][a-z\s,/&+\-]{2,60}))?/ig;
  while ((m = degRe.exec(jd)) !== null) {
    const lvl = m[1].toLowerCase().replace(/['']s$/, '');
    const field = (m[2] || '').trim().replace(/\s+/g, ' ').slice(0, 60);
    quals.push({ type: 'degree', text: `${capitalize(lvl)}'s${field ? ' in ' + field : ' degree'}`, level: lvl, field, key: `deg:${lvl}:${field}` });
  }
  // certifications
  const certRe = /\b(certified|certification|certificate|licensed?)\s+(?:in\s+)?([a-z][a-z\s,/&+\-]{2,60})/ig;
  while ((m = certRe.exec(jd)) !== null) {
    const what = m[2].trim();
    if (what.length < 3) continue;
    quals.push({ type: 'cert', text: `Certification: ${what}`, key: `cert:${what.toLowerCase()}` });
  }
  // requirement-like bullets
  for (const line of lines) {
    if (line.length < 8 || line.length > 220) continue;
    if (reqHints.test(line) && !/responsibilit/i.test(line)) {
      const clean = line.replace(/^[\-*•●]\s*/, '').trim();
      if (clean.length < 8) continue;
      quals.push({ type: 'general', text: clean, key: 'gen:' + clean.slice(0, 40).toLowerCase() });
    }
  }
  const seen = new Set();
  return quals.filter(q => seen.has(q.key) ? false : (seen.add(q.key), true)).slice(0, 12);
}

function maxYears(text) {
  if (!text) return 0;
  const re = /(\d{1,2})\+?\s*(?:years?|yrs?)\b/ig;
  let m, best = 0;
  while ((m = re.exec(text)) !== null) {
    const n = +m[1];
    if (n > 0 && n <= 40 && n > best) best = n;
  }
  return best;
}

function qualMatchesEvidence(qual, evidence) {
  const ev = (evidence || '').toLowerCase();
  if (qual.type === 'experience') {
    return maxYears(evidence) >= (qual.years || 0);
  }
  if (qual.type === 'degree') {
    const lvl = qual.level || '';
    const lvlHit = new RegExp(`\\b${escapeRegex(lvl)}\\b|\\bb\\.?s\\.?\\b|\\bm\\.?s\\.?\\b|\\bph\\.?d\\.?\\b|\\bmba\\b`, 'i').test(evidence) ||
                   /\b(bachelor|master|doctorate|associate)\b/i.test(evidence);
    if (!lvlHit) return false;
    if (!qual.field) return true;
    const fieldToks = tokens(qual.field).filter(t => t.length > 2);
    return fieldToks.length === 0 || fieldToks.some(t => ev.includes(t));
  }
  if (qual.type === 'cert') {
    const target = (qual.text || '').toLowerCase().replace(/^certification:\s*/, '').trim();
    if (!target) return false;
    // require at least the first significant word
    const firstWord = target.split(/\s+/).find(w => w.length > 2);
    return firstWord ? ev.includes(firstWord) : false;
  }
  // general — need 2+ content tokens overlap, or 40% of content tokens
  const t = tokens(qual.text).filter(w => w.length > 3).slice(0, 8);
  if (t.length === 0) return false;
  const hits = t.filter(w => ev.includes(w)).length;
  return hits >= Math.max(2, Math.ceil(t.length * 0.4));
}

// Best-effort job title detector — looks at first few short lines
function detectJobTitle(jd) {
  const TITLE_NOISE = /(about|description|overview|responsibilit|requirement|qualif|summary|posting|company|location|department|salary|benefit|equal opportunity)/i;
  const lines = jd.split(/\r?\n/).map(l => l.trim());
  // explicit "Position: X" / "Job Title: X"
  for (const l of lines.slice(0, 25)) {
    const m = l.match(/^(?:position|role|job\s*title|title)\s*[:\-]\s*(.{4,80})$/i);
    if (m) return cleanTitle(m[1]);
  }
  // first short non-noise line in the top of the JD
  for (const l of lines.slice(0, 8)) {
    if (l.length < 4 || l.length > 80) continue;
    if (TITLE_NOISE.test(l)) continue;
    if (!/[A-Za-z]/.test(l)) continue;
    if (l.split(/\s+/).length > 12) continue;
    return cleanTitle(l);
  }
  return '';
}
function cleanTitle(s) {
  return s.replace(/[—–\-•|].*$/, '').replace(/\s+/g, ' ').trim();
}

function jobTitleAlignment(jdTitle, cvText) {
  if (!jdTitle) return { title: '', score: 50, hits: [] };
  const titleToks = tokens(jdTitle).filter(t => t.length > 2);
  const cv = (cvText || '').toLowerCase();
  const hits = titleToks.filter(t => cv.includes(t));
  const score = titleToks.length ? Math.round((hits.length / titleToks.length) * 100) : 50;
  return { title: jdTitle, score, hits };
}

// Company name — conservative. Returns '' if not confident.
// Uses [ \t]+ for inter-word spacing (no newlines) and stops at any sentence punctuation.
function detectCompany(jd) {
  if (!jd) return '';
  // Look for explicit "Company: X"
  let m = jd.match(/^[ \t]*company[ \t]*[:\-][ \t]*(.{2,60})$/im);
  if (m) return cleanEntity(m[1]);
  // "at <Capitalized Phrase>" — narrow capture, stops at end-of-clause or end-of-line
  m = jd.match(/\bat[ \t]+([A-Z][A-Za-z0-9&\-]+(?:[ \t]+[A-Z][A-Za-z0-9&\-]+){0,3})(?=[ \t]*[.,!?;:\n]|[ \t]*$)/);
  if (m) {
    const candidate = cleanEntity(m[1]);
    if (!/^(senior|junior|the|our|this|that|least|most|all|every|once|some|any)\b/i.test(candidate)) return candidate;
  }
  // "join <Capitalized>" / "with <Capitalized>"
  m = jd.match(/\b(?:join|with)[ \t]+([A-Z][A-Za-z0-9&\-]+(?:[ \t]+[A-Z][A-Za-z0-9&\-]+){0,2})(?=[ \t]*[.,!?;:\n]|[ \t]*$)/);
  if (m) {
    const candidate = cleanEntity(m[1]);
    if (!/^(senior|junior|the|our|this|that|least|most|all|every|once|some|any)\b/i.test(candidate)) return candidate;
  }
  return '';
}
function cleanEntity(s) {
  return s.replace(/[.,;:]$/, '').replace(/[ \t]+/g, ' ').trim();
}

// =========================================================
// 11. ATS FORMAT SCORE
// =========================================================
function atsFormatScore(text) {
  const reasons = [];
  let score = 100;
  const t = text || '';
  const lines = t.split(/\r?\n/);

  if (!t || t.length < 400) { score -= 25; reasons.push('CV text is short — ATS may have parsed it poorly.'); }
  const odd = (t.match(/[^\x09\x0A\x0D\x20-\x7E£€©®•·–—…]/g) || []).length;
  const oddRatio = odd / Math.max(t.length, 1);
  if (oddRatio > 0.04) { score -= 12; reasons.push('Many non-standard characters detected — often from icons or multi-column layouts.'); }

  const SECTIONS = ['summary','experience','education','skills','certifications','projects'];
  const present = SECTIONS.filter(s => new RegExp(`(^|\\n)\\s*${s}\\b`, 'i').test(t));
  if (present.length < 3) { score -= 15; reasons.push(`Missing standard section headings (found: ${present.join(', ') || 'none'}).`); }
  else reasons.push(`Found standard headings: ${present.map(capitalize).join(', ')}.`);

  const hasEmail = /[\w.+\-]+@[\w\-]+\.[\w.\-]+/.test(t);
  const hasPhone = /(\+?\d[\d\s().\-]{7,}\d)/.test(t);
  if (!hasEmail) { score -= 8;  reasons.push('No email address detected.'); }
  if (!hasPhone) { score -= 5;  reasons.push('No phone number detected.'); }

  const wc = wordCount(t);
  if (wc < 200)  { score -= 10; reasons.push('Under 200 words — likely missing detail.'); }
  if (wc > 1400) { score -= 5;  reasons.push('Over 1400 words — consider trimming.'); }

  const bullets = lines.filter(l => /^\s*[\-*•●]/.test(l)).length;
  if (bullets < 4) { score -= 6; reasons.push('Few bullet points — ATS prefers bulleted achievements.'); }
  else reasons.push(`${bullets} bullet points detected — good ATS readability.`);

  return { score: clamp(score, 0, 100), reasons };
}

// =========================================================
// 12. MAIN ANALYSIS
// =========================================================
function analyzeMatch() {
  const cv = state.cv.text || '';
  const tr = state.transcript.text || '';
  const jd = state.jd || '';
  const evidence = cv + '\n' + tr;

  // Keywords
  const jdTop = topJdKeywords(jd, 30);
  const matched = [], missing = [];
  let wSum = 0, wMatched = 0;
  for (const { term, count } of jdTop) {
    wSum += count;
    if (phrasePresent(cv, term)) { matched.push(term); wMatched += count; }
    else missing.push(term);
  }
  const keywordScore = wSum ? Math.round((wMatched / wSum) * 100) : 0;

  // Skills
  const jdSkills = [...extractSkills(jd)];
  const cvSkills = extractSkills(evidence);
  const matchedSkills = jdSkills.filter(s => cvSkills.has(s));
  const missingSkills = jdSkills.filter(s => !cvSkills.has(s));
  const skillScore = jdSkills.length ? Math.round((matchedSkills.length / jdSkills.length) * 100) : 60;

  // Qualifications
  const quals = extractRequiredQuals(jd);
  const qualResults = quals.map(q => ({ ...q, found: qualMatchesEvidence(q, evidence) }));
  const qualScore = qualResults.length
    ? Math.round((qualResults.filter(q => q.found).length / qualResults.length) * 100)
    : 70;

  // Education / transcript
  let educationScore = 60;
  let transcriptImpact = '';
  const cvEdu = /\b(bachelor|master|phd|ph\.d|doctorate|associate|b\.s|m\.s|degree)\b/i.test(cv);
  const jdEdu = /\b(bachelor|master|phd|ph\.d|doctorate|associate|degree)\b/i.test(jd);
  if (jdEdu) educationScore = cvEdu ? 80 : 40;
  if (tr) {
    const gpaMatch = tr.match(/gpa[:\s]*([0-4]\.\d{1,2})/i);
    const gpaNum = gpaMatch ? parseFloat(gpaMatch[1]) : null;
    const jdTermsTop = jdTop.map(j => j.term).slice(0, 15);
    const courseHits = jdTermsTop.filter(t => tr.toLowerCase().includes(t));
    const bits = [];
    let bump = 0;
    if (gpaNum !== null) {
      const b = gpaNum >= 3.5 ? 12 : gpaNum >= 3.0 ? 8 : 3;
      bump += b;
      bits.push(`GPA ${gpaNum} detected (+${b}).`);
    }
    if (courseHits.length) {
      const b = Math.min(15, courseHits.length * 3);
      bump += b;
      bits.push(`Coursework overlap on ${courseHits.length} JD term${courseHits.length === 1 ? '' : 's'}: ${courseHits.slice(0, 6).join(', ')} (+${b}).`);
    }
    educationScore = clamp(educationScore + bump, 0, 100);
    transcriptImpact = bits.length ? bits.join(' ') : 'Transcript parsed but no clear GPA or course overlap detected.';
  }

  // Title
  const jdTitle = detectJobTitle(jd);
  const title = jobTitleAlignment(jdTitle, cv);

  // Experience
  const cvYrs = maxYears(cv);
  const jdYrs = maxYears(jd);
  let expScore = 70;
  if (jdYrs) expScore = cvYrs >= jdYrs ? 95 : Math.max(20, Math.round((cvYrs / jdYrs) * 90));

  // ATS format on the CV
  const atsCv = atsFormatScore(cv);

  // Section completeness
  const cvSecs = ['summary','experience','education','skills'].filter(s => new RegExp(`(^|\\n)\\s*${s}\\b`, 'i').test(cv)).length;
  const sectionScore = Math.round((cvSecs / 4) * 100);

  // Overall (weighted)
  const W = { keyword: 0.22, skill: 0.22, qual: 0.16, education: 0.12, title: 0.10, experience: 0.10, ats: 0.05, section: 0.03 };
  const overall = Math.round(
    keywordScore   * W.keyword +
    skillScore     * W.skill   +
    qualScore      * W.qual    +
    educationScore * W.education +
    title.score    * W.title +
    expScore       * W.experience +
    atsCv.score    * W.ats +
    sectionScore   * W.section
  );

  // Confidence
  const confidence = clamp(
    40 +
    (cv.length > 800 ? 25 : 10) +
    (jd.length > 600 ? 20 : 8) +
    (tr ? 15 : 0),
    0, 100
  );

  // Reasons
  const reasons = [
    `Matched ${matched.length} of ${jdTop.length} top JD keywords (${keywordScore}/100).`,
    `Detected ${jdSkills.length} JD-relevant skills; your evidence covers ${matchedSkills.length} (${skillScore}/100).`,
    `${qualResults.filter(q => q.found).length}/${qualResults.length || 0} explicit qualifications evidenced in your CV${tr ? '/transcript' : ''}.`,
    jdYrs ? `JD asks for ${jdYrs}+ years; your evidence signals about ${cvYrs} year${cvYrs === 1 ? '' : 's'} (${expScore}/100).` : `No explicit years requirement parsed — experience scored neutrally.`,
    title.title ? `Detected JD title “${title.title}” — title-word overlap ${title.score}/100.` : `Job title not clearly parsed from the JD.`,
    `ATS formatting on your CV scored ${atsCv.score}/100.`,
    tr ? `Transcript boosted the education signal.` : `No transcript uploaded — education signal relies on the CV.`,
  ];

  // Recommendations
  const recs = [];
  if (missing.length)        recs.push(`Weave in missing keywords where truthful: ${missing.slice(0, 8).join(', ')}.`);
  if (missingSkills.length)  recs.push(`Highlight (or upskill on) JD-emphasized skills: ${missingSkills.slice(0, 8).join(', ')}.`);
  if (atsCv.score < 75)      recs.push('Re-export your CV as single-column, no-icons, no-tables for better ATS parsing.');
  if (sectionScore < 75)     recs.push('Add standard sections: Summary, Experience, Education, Skills.');
  if (title.score < 50 && title.title) recs.push(`Mirror the JD title (“${title.title}”) in your Summary if accurate.`);
  qualResults.filter(q => !q.found).slice(0, 4).forEach(q => recs.push(`Address missing qualification: ${q.text}.`));
  if (recs.length === 0) recs.push('You are well aligned. Submit with confidence and personalize the cover letter.');

  // Verdict
  let verdict, verdictTitle, verdictBlurb;
  if (overall >= 78) {
    verdict = 'good';
    verdictTitle = 'Strong match — apply.';
    verdictBlurb = 'Your CV aligns well across keywords, skills, and qualifications. Tailor lightly and submit.';
  } else if (overall >= 60) {
    verdict = 'ok';
    verdictTitle = 'Moderate match — review and improve.';
    verdictBlurb = 'Worth applying after a focused tailoring pass on the missing keywords and qualifications.';
  } else {
    verdict = 'weak';
    verdictTitle = 'Weak match — not recommended without significant tailoring.';
    verdictBlurb = 'Gaps are large enough that you should upskill, reframe transferable experience, or look at closer roles.';
  }

  return {
    overall, verdict, verdictTitle, verdictBlurb, confidence,
    signals: [
      { key: 'keyword',    label: 'Keyword overlap',    score: keywordScore, tip: 'Top JD terms found in your CV.' },
      { key: 'skill',      label: 'Skills match',       score: skillScore,   tip: 'JD-mentioned skills present in your CV / transcript.' },
      { key: 'qual',       label: 'Required quals',     score: qualScore,    tip: 'Stated requirements (years, degree, certs, bullets) evidenced in your CV.' },
      { key: 'education',  label: 'Education / transcript', score: educationScore, tip: 'Degree level and field match, plus transcript GPA and coursework boost.' },
      { key: 'title',      label: 'Job title alignment', score: title.score, tip: 'Overlap between JD title words and your CV vocabulary.' },
      { key: 'experience', label: 'Experience level',   score: expScore,     tip: 'Years of experience inferred from JD vs CV.' },
      { key: 'ats',        label: 'ATS formatting',     score: atsCv.score,  tip: 'How well your CV will parse through automated tracking systems.' },
    ],
    extras: { sectionScore, cvAtsReasons: atsCv.reasons, jdTitle },
    keywords: { matched, missing, jdTop },
    skills:   { matched: matchedSkills, missing: missingSkills, required: jdSkills },
    quals: qualResults,
    title,
    reasons, recs, transcriptImpact,
  };
}

// =========================================================
// 13. RENDER RESULTS
// =========================================================
function renderResults(a) {
  const sec = $('#step-results'); sec.hidden = false;

  // Verdict ring
  $('#overallScore').textContent = a.overall;
  const ringFg = $('#ringFg');
  ringFg.style.strokeDashoffset = String(100 - a.overall);
  ringFg.setAttribute('stroke', a.verdict === 'good' ? '#22c55e' : a.verdict === 'ok' ? '#f59e0b' : '#ef4444');
  const badge = $('#verdictBadge');
  badge.className = 'verdict-badge ' + a.verdict;
  badge.textContent = a.verdict === 'good' ? 'Strong match' : a.verdict === 'ok' ? 'Moderate match' : 'Weak match';
  $('#verdictTitle').textContent = a.verdictTitle;
  $('#verdictBlurb').textContent = a.verdictBlurb;
  $('#confFill').style.width = a.confidence + '%';
  $('#confLabel').textContent = a.confidence + '%';

  // Signals
  const sig = $('#signals'); sig.innerHTML = '';
  a.signals.forEach(s => {
    const tone = s.score >= 75 ? 'good' : s.score >= 55 ? 'ok' : 'weak';
    const node = el('div', { className: `signal ${tone}` });

    const top = el('div', { className: 'signal-top' },
      el('strong', {}, s.label),
      (() => {
        const tip = el('span', { className: 'tip-trigger', tabIndex: 0 }, 'ⓘ');
        tip.setAttribute('data-tip', s.tip);
        tip.setAttribute('aria-label', s.tip);
        return tip;
      })()
    );
    const scoreLine = el('div', { className: 'signal-score' }, String(s.score),
      el('span', { style: 'font-size:14px;color:var(--text-muted);font-weight:600;' }, '/100')
    );
    const bar = el('div', { className: 'signal-bar' });
    const fill = el('div'); fill.style.width = s.score + '%';
    bar.appendChild(fill);
    const sub = el('small', {}, toneText(tone));

    node.append(top, scoreLine, bar, sub);
    sig.appendChild(node);
  });

  // Keywords
  const mc = $('#matchedChips'), msc = $('#missingChips');
  mc.innerHTML = ''; msc.innerHTML = '';
  if (a.keywords.matched.length === 0) mc.appendChild(el('span', { className: 'muted small' }, 'No top JD keywords matched yet.'));
  if (a.keywords.missing.length === 0) msc.appendChild(el('span', { className: 'muted small' }, 'Nothing major missing — nice.'));
  a.keywords.matched.forEach(k => mc.appendChild(chip(k)));
  a.keywords.missing.forEach(k => msc.appendChild(chip(k)));
  $('#matchedCount').textContent = a.keywords.matched.length;
  $('#missingCount').textContent = a.keywords.missing.length;

  // Quals
  const ql = $('#qualList'); ql.innerHTML = '';
  if (!a.quals.length) ql.appendChild(el('li', { className: 'muted' }, 'No explicit requirements parsed from the JD.'));
  a.quals.forEach(q => ql.appendChild(el('li', { className: q.found ? 'found' : 'missing' }, q.text)));

  // Recs
  const rl = $('#recList'); rl.innerHTML = '';
  a.recs.forEach(r => rl.appendChild(el('li', {}, r)));

  // Reasoning
  const rs = $('#reasonList'); rs.innerHTML = '';
  a.reasons.forEach(r => rs.appendChild(el('li', {}, r)));
  if (a.extras.cvAtsReasons.length) {
    rs.appendChild(el('li', { className: 'muted' }, '— CV ATS-format notes —'));
    a.extras.cvAtsReasons.forEach(r => rs.appendChild(el('li', { className: 'muted' }, r)));
  }

  // Transcript impact
  $('#transcriptImpact').textContent = a.transcriptImpact ||
    (state.transcript.text
      ? 'Transcript uploaded but did not change the education signal materially.'
      : 'No transcript uploaded — education signal scored from the CV only.');

  // Clear stale outputs from a previous analysis
  state.resume = ''; state.cover = '';
  $('#resumeOut').value = ''; $('#coverOut').value = '';
  $('#resumeOutCard').hidden = true;
  $('#coverOutCard').hidden = true;
  $('#outputs').hidden = true;

  sec.scrollIntoView({ behavior: 'smooth', block: 'start' });
  updateStepper();

  // Save to history with dedup
  const hash = quickHash(state.cv.text + '|' + state.jd);
  if (hash !== state.lastAnalysisHash) {
    saveHistory({
      when: new Date().toISOString(),
      score: a.overall, verdict: a.verdict,
      cvName: state.cv.name || 'CV',
      jdSnippet: state.jd.slice(0, 90),
      hash,
    });
    state.lastAnalysisHash = hash;
    renderHistory();
  }
}

function toneText(t) { return t === 'good' ? 'Strong area' : t === 'ok' ? 'Acceptable — could improve' : 'Needs attention'; }
function chip(text) { const c = el('span', { className: 'chip' }); c.textContent = text; return c; }

// =========================================================
// 14. TABS
// =========================================================
$$('.tab').forEach(t => t.addEventListener('click', () => {
  $$('.tab').forEach(x => { x.classList.remove('active'); x.setAttribute('aria-selected', 'false'); });
  t.classList.add('active'); t.setAttribute('aria-selected', 'true');
  $$('.tab-pane').forEach(p => p.classList.toggle('active', p.dataset.pane === t.dataset.tab));
}));

// =========================================================
// 15. ANALYZE
// =========================================================
$('#analyzeBtn').addEventListener('click', async () => {
  if (!state.cv.text || state.jd.length < 30) { toast('Add your CV and a job description first.', true); return; }
  showLoader('Analyzing match…');
  try {
    await new Promise(r => setTimeout(r, 250)); // tiny delay so the UI feels considered
    state.analysis = analyzeMatch();
    renderResults(state.analysis);
    toast('Analysis complete.');
  } catch (e) {
    console.error(e);
    toast('Analysis failed: ' + (e.message || 'unknown error'), true);
  } finally {
    hideLoader();
  }
});

// =========================================================
// 16. CV SECTION PARSING (for generation)
// =========================================================
const SECTION_NAMES = {
  summary:        ['summary','objective','profile','professional summary'],
  experience:     ['experience','work experience','professional experience','employment history','employment'],
  education:      ['education','academic background'],
  skills:         ['skills','technical skills','core competencies','key skills'],
  certifications: ['certifications','licenses','licenses & certifications','licenses and certifications'],
  projects:       ['projects','selected projects','key projects'],
};

function isSectionHeading(line) {
  // A heading is a short, mostly-uppercase or title-cased line with no trailing sentence punctuation.
  const t = line.trim();
  if (t.length === 0 || t.length > 50) return null;
  if (/[.,;:?!]$/.test(t)) return null;
  const lower = t.toLowerCase().replace(/[^a-z& ]+/g, '').trim();
  for (const [key, syns] of Object.entries(SECTION_NAMES)) {
    if (syns.includes(lower)) return key;
  }
  return null;
}

function parseCvSections(cv) {
  const out = { contact: '', name: '', summary: '', experience: '', education: '', skills: '', certifications: '', projects: '', other: '' };
  if (!cv) return out;
  const lines = cv.split(/\r?\n/);

  // Header: first non-empty lines until we hit a section heading or 6 lines
  const headerLines = [];
  let i = 0;
  for (; i < lines.length && headerLines.length < 6; i++) {
    const l = lines[i].trim();
    if (!l) continue;
    if (isSectionHeading(l)) break;
    headerLines.push(l);
  }
  out.contact = headerLines.join('\n');
  // Name guess: first header line if it looks like a name (no @, no digits)
  const firstName = headerLines.find(l => l && !/[@\d]/.test(l) && l.length < 60);
  if (firstName) out.name = firstName;

  // Section parsing from i onward
  const buf = { summary: [], experience: [], education: [], skills: [], certifications: [], projects: [], other: [] };
  let current = 'other';
  for (; i < lines.length; i++) {
    const l = lines[i];
    const head = isSectionHeading(l);
    if (head) { current = head; continue; }
    buf[current].push(l);
  }
  for (const k of Object.keys(buf)) out[k] = buf[k].join('\n').replace(/\n{3,}/g, '\n\n').trim();
  return out;
}

// =========================================================
// 17. RESUME GENERATION (deterministic)
// =========================================================
const ACTION_VERBS = ['Led','Built','Designed','Implemented','Delivered','Drove','Owned','Launched','Optimized','Automated','Architected','Streamlined','Coordinated','Established','Mentored'];

// Deterministic pick — same input → same verb
function pickVerb(text, fallbackKey) {
  const lc = (text || '').toLowerCase();
  if (/\bteam|stakehold|cross|mentor\b/.test(lc)) return 'Led';
  if (/\btest|qa|verif|valid\b/.test(lc))         return 'Validated';
  if (/\bdata|analy|report|dashboard|insight\b/.test(lc)) return 'Analyzed';
  if (/\bdesign|architect|build|develop\b/.test(lc)) return 'Designed';
  if (/\bautomat|script|pipeline|tool\b/.test(lc)) return 'Automated';
  if (/\boptim|reduce|improv|perform|efficien\b/.test(lc)) return 'Optimized';
  // deterministic fallback
  const seed = quickHash(fallbackKey || lc);
  return ACTION_VERBS[seed % ACTION_VERBS.length];
}

// Recognize a true bullet line: starts with bullet marker, or is short and starts with a known action verb
function looksLikeBullet(line) {
  const t = line.trim();
  if (!t) return false;
  if (/^[\-*•●▪►·]\s+/.test(t)) return true;
  if (t.length > 200) return false;
  const firstWord = t.split(/\s+/)[0];
  if (!firstWord) return false;
  // Match a verb-ish first word (-ed / -ing endings, or known action verbs)
  if (/^(?:[A-Z][a-z]+(?:ed|ing|s))\b/.test(firstWord)) {
    // Avoid treating "Education" or "Experience" headings as bullets
    if (isSectionHeading(t)) return false;
    return true;
  }
  return false;
}

function strengthenBullet(line, weaveKeyword) {
  let s = line.replace(/^[\-*•●▪►·\s]+/, '').trim();
  if (!s) return '';
  if (/^(?:responsible for|worked on|helped|assisted|involved in|tasked with|in charge of|duties included)/i.test(s)) {
    s = s.replace(/^(?:responsible for|worked on|helped|assisted|involved in|tasked with|in charge of|duties included)\s*/i, '');
    s = `${pickVerb(s, s)} ${s}`;
  } else {
    s = s.charAt(0).toUpperCase() + s.slice(1);
  }
  if (weaveKeyword && !new RegExp(`\\b${escapeRegex(weaveKeyword)}\\b`, 'i').test(s)) {
    s = s.replace(/[.!?]?$/, '') + `, leveraging ${weaveKeyword}.`;
  } else if (!/[.!?]$/.test(s)) {
    s += '.';
  }
  return '• ' + s;
}

function generateResume() {
  const a = state.analysis;
  if (!a) { toast('Run an analysis first.', true); return; }
  const cv = state.cv.text;
  const parts = parseCvSections(cv);
  const jdTitle = (a.title && a.title.title) || 'Target Role';
  const missing = a.keywords.missing.slice(0, 10);
  const matched = a.keywords.matched.slice(0, 14);

  // Summary
  const baseSummary = (parts.summary || '').replace(/\s+/g, ' ').trim().slice(0, 600);
  const matchedSkillsTitled = a.skills.matched.slice(0, 6).map(titleCase);
  const synthSummary = `Results-driven professional targeting the ${jdTitle} role, with proven experience across ${[...new Set([...matched.slice(0, 6), ...matchedSkillsTitled])].slice(0, 8).join(', ') || 'core competencies relevant to this position'}. Track record of delivering measurable outcomes, partnering cross-functionally, and applying ${matchedSkillsTitled.slice(0, 3).join(', ') || 'relevant tools and methodologies'} to drive results.`;
  const summary = baseSummary || synthSummary;

  // Skills
  const allSkillsSet = new Set([
    ...a.skills.matched.map(skillDisplay),
    ...[...extractSkills(cv)].map(skillDisplay),
  ]);
  const allSkills = [...allSkillsSet];
  const groupedSkills = allSkills.length
    ? `• Core: ${allSkills.slice(0, 8).join(', ')}` +
      (allSkills.length > 8 ? `\n• Additional: ${allSkills.slice(8, 20).join(', ')}` : '')
    : '• (List your top tools and skills here)';

  // Experience — deterministic weaving
  const expLines = (parts.experience || '').split(/\r?\n/);
  const strongExp = [];
  let woven = 0;
  const weaveBudget = Math.min(3, missing.length);
  for (let idx = 0; idx < expLines.length; idx++) {
    const line = expLines[idx];
    const t = line.trim();
    if (!t) { strongExp.push(''); continue; }
    if (looksLikeBullet(line)) {
      // Deterministically decide whether this bullet receives a keyword weave:
      // weave the first weaveBudget bullets that don't already contain the keyword
      let weave = null;
      if (woven < weaveBudget) {
        const candidate = missing[woven];
        if (candidate && !new RegExp(`\\b${escapeRegex(candidate)}\\b`, 'i').test(t)) {
          weave = candidate;
          woven++;
        }
      }
      strongExp.push(strengthenBullet(line, weave));
    } else {
      strongExp.push(t); // preserve job titles, dates, company lines
    }
  }
  const experience = strongExp.join('\n').replace(/\n{3,}/g, '\n\n').trim() || '(Add your experience bullets here.)';

  const education      = parts.education      || '(Add education details here.)';
  const certifications = parts.certifications || '';
  const projects       = parts.projects       || '';

  // Compose
  const contactHeader = (parts.contact || parts.name || 'Your Name').trim();
  const sections = [];
  sections.push(contactHeader, '');
  sections.push('SUMMARY', summary, '');
  sections.push('SKILLS', groupedSkills, '');
  sections.push('EXPERIENCE', experience, '');
  sections.push('EDUCATION', education);
  if (certifications) sections.push('', 'CERTIFICATIONS', certifications);
  if (projects)       sections.push('', 'PROJECTS', projects);

  const resume = sections.join('\n').replace(/\n{3,}/g, '\n\n').trim();
  state.resume = resume;

  // ATS score on the generated resume
  const atsGen = atsFormatScore(resume);
  const newCoverage = a.keywords.jdTop.filter(({ term }) => phrasePresent(resume, term)).length;
  const newCovPct = a.keywords.jdTop.length ? Math.round((newCoverage / a.keywords.jdTop.length) * 100) : 80;
  const finalAts = clamp(Math.round(atsGen.score * 0.55 + newCovPct * 0.45), 0, 100);

  // Render
  $('#outputs').hidden = false;
  $('#resumeOutCard').hidden = false;
  $('#resumeOut').value = resume;
  $('#atsPill').textContent = `ATS score: ${finalAts}/100`;
  $('#cvExcerpt').textContent = cv.slice(0, 1400) + (cv.length > 1400 ? '\n…' : '');

  const diffs = [];
  diffs.push(`Reframed summary toward the ${jdTitle} role.`);
  if (woven > 0) diffs.push(`Wove ${woven} missing JD keyword${woven === 1 ? '' : 's'} into experience bullets where they fit naturally.`);
  diffs.push(`Reordered to ATS-standard sections: Summary → Skills → Experience → Education.`);
  diffs.push(`Replaced weak openers (e.g., "responsible for") with stronger action verbs.`);
  if (allSkills.length) diffs.push(`Consolidated ${allSkills.length} skill terms into a single Skills block.`);
  diffs.push(`Removed non-ATS characters and enforced single-column layout.`);
  const dl = $('#resumeDiff'); dl.innerHTML = '';
  diffs.forEach(d => dl.appendChild(el('li', {}, d)));

  const ae = $('#atsExplain'); ae.innerHTML = '';
  ae.appendChild(el('li', {}, `Formatting subscore: ${atsGen.score}/100 (weighted 55%).`));
  ae.appendChild(el('li', {}, `Keyword coverage subscore: ${newCovPct}/100 from ${newCoverage}/${a.keywords.jdTop.length} JD terms (weighted 45%).`));
  atsGen.reasons.forEach(r => ae.appendChild(el('li', { className: 'muted' }, r)));

  $('#resumeOutCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
  toast('ATS resume generated.');
  updateStepper();
}

function skillDisplay(s) {
  // SQL / GIS / GCP / NLP / SEO / SAS / VBA / CSS / HTML / API / etc. stay uppercase.
  const upper = new Set(['sql','gis','gcp','nlp','seo','sas','vba','css','html','api','aws','crm','llm','ehr','sem','etl','elt','qa','ci/cd']);
  if (upper.has(s)) return s.toUpperCase();
  return s.split(/\s+/).map(w => w ? w[0].toUpperCase() + w.slice(1) : '').join(' ');
}

// =========================================================
// 18. COVER LETTER GENERATION
// =========================================================
function generateCover() {
  const a = state.analysis;
  if (!a) { toast('Run an analysis first.', true); return; }
  const cv = state.cv.text;
  const parts = parseCvSections(cv);
  const jdTitleRaw = (a.title && a.title.title) || '';
  const jdTitle = jdTitleRaw || 'the role';
  const company = detectCompany(state.jd);

  const today = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
  const topSkills = [...new Set([
    ...a.skills.matched.map(skillDisplay),
    ...a.keywords.matched.slice(0, 5),
  ])].slice(0, 5);
  const evidencedQuals = a.quals.filter(q => q.found).slice(0, 2).map(q => q.text.replace(/^Certification:\s*/, ''));
  const candName = (parts.name && parts.name.length < 60) ? parts.name : 'Your Name';

  // Build paragraphs carefully — avoid awkward phrasing
  const greetingCompany = company || 'your team';
  const intro = `I am writing to express my strong interest in the ${jdTitleRaw ? jdTitleRaw + ' position' : 'position'}${company ? ' at ' + company : ''}. After reviewing the job description, I am confident that my background aligns closely with the responsibilities and qualifications you describe.`;

  const skillsLine = topSkills.length
    ? `Throughout my career, I have built hands-on experience with ${topSkills.join(', ')}.`
    : `Throughout my career, I have built hands-on experience directly relevant to this position.`;

  const qualsLine = evidencedQuals.length
    ? ` I also bring ${evidencedQuals.join(' and ')}, which map directly to the requirements in your posting.`
    : '';

  const impactSeed = (topSkills.join(' ') + ' ' + jdTitle).toLowerCase();
  const impactVerb = pickVerb(impactSeed, 'impact').toLowerCase();
  const impactLine = `In recent roles I have ${impactVerb} measurable outcomes through ${a.skills.matched.slice(0, 2).map(skillDisplay).join(' and ') || 'focused, structured execution'}, and I am eager to bring that same impact to ${greetingCompany}.`;

  const fitLine = `What draws me specifically to ${greetingCompany} is the opportunity to contribute where ${a.keywords.matched.slice(0, 3).join(', ') || 'relevant expertise'} can meaningfully move the needle. I take pride in clear communication, collaborative problem-solving, and shipping work that holds up under real-world constraints — qualities I see reflected in the role description.`;

  const close = `I would welcome the chance to discuss how my experience could support your team's goals. Thank you for your time and consideration; I look forward to the opportunity to speak further.`;

  const body =
`${today}

Hiring Manager
${company || 'Hiring Team'}

Dear Hiring Manager,

${intro}

${skillsLine}${qualsLine} ${impactLine}

${fitLine}

${close}

Sincerely,
${candName}`;

  state.cover = body;
  $('#outputs').hidden = false;
  $('#coverOutCard').hidden = false;
  $('#coverOut').value = body;
  $('#coverOutCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
  toast('Cover letter generated.');
  updateStepper();
}

$('#genResumeBtn').addEventListener('click', () => { try { generateResume(); } catch (e) { console.error(e); toast(e.message || 'Failed to generate resume.', true); } });
$('#genCoverBtn') .addEventListener('click', () => { try { generateCover();  } catch (e) { console.error(e); toast(e.message || 'Failed to generate cover letter.', true); } });

// Keep state in sync with edits
$('#resumeOut').addEventListener('input', e => state.resume = e.target.value);
$('#coverOut') .addEventListener('input', e => state.cover  = e.target.value);

// =========================================================
// 19. COPY / DOWNLOAD
// =========================================================
function copyText(text) {
  if (navigator.clipboard && navigator.clipboard.writeText) {
    return navigator.clipboard.writeText(text).catch(() => fallbackCopy(text));
  }
  return Promise.resolve(fallbackCopy(text));
}
function fallbackCopy(text) {
  const ta = document.createElement('textarea');
  ta.value = text;
  ta.style.position = 'fixed'; ta.style.opacity = '0';
  document.body.appendChild(ta); ta.focus(); ta.select();
  try { document.execCommand('copy'); } catch (_) {}
  ta.remove();
}
function downloadText(text, filename) {
  triggerDownload(new Blob([text], { type: 'text/plain;charset=utf-8' }), filename);
}
function downloadDoc(text, filename) {
  const html = `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
<head><meta charset='utf-8'><title>${escapeHtml(filename)}</title>
<style>body{font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:1.4;color:#111;}pre{font-family:Calibri,Arial,sans-serif;white-space:pre-wrap;margin:0;}</style>
</head><body><pre>${escapeHtml(text)}</pre></body></html>`;
  triggerDownload(new Blob(['﻿', html], { type: 'application/msword' }), filename);
}
function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

$('#copyResume') .addEventListener('click', () => copyText($('#resumeOut').value).then(() => toast('Resume copied.')));
$('#dlResume')   .addEventListener('click', () => downloadText($('#resumeOut').value, 'ATS_Resume.txt'));
$('#dlResumeDocx').addEventListener('click', () => downloadDoc($('#resumeOut').value, 'ATS_Resume.doc'));
$('#copyCover')  .addEventListener('click', () => copyText($('#coverOut').value).then(() => toast('Cover letter copied.')));
$('#dlCover')    .addEventListener('click', () => downloadText($('#coverOut').value, 'Cover_Letter.txt'));
$('#dlCoverDocx').addEventListener('click', () => downloadDoc($('#coverOut').value, 'Cover_Letter.doc'));

// =========================================================
// 20. HISTORY
// =========================================================
function loadHistory() {
  try { return JSON.parse(localStorage.getItem('ff-history') || '[]'); } catch { return []; }
}
function saveHistory(entry) {
  const arr = loadHistory().filter(e => e.hash !== entry.hash);
  arr.unshift(entry);
  try { localStorage.setItem('ff-history', JSON.stringify(arr.slice(0, 5))); } catch (_) {}
}
function renderHistory() {
  const list = $('#historyList');
  const items = loadHistory();
  if (!items.length) {
    list.innerHTML = '';
    list.appendChild(el('p', { className: 'muted' }, 'Your last 5 analyses will appear here (saved locally in your browser).'));
    return;
  }
  list.innerHTML = '';
  items.forEach(it => {
    const tone = it.verdict === 'good' ? '#22c55e' : it.verdict === 'ok' ? '#f59e0b' : '#ef4444';
    const card = el('div', { className: 'history-item', tabIndex: 0 });
    const score = el('span', { className: 'hi-score' }, String(it.score));
    score.style.color = tone;
    card.append(
      score,
      el('h5', {}, it.cvName || 'CV'),
      el('p', {}, (it.jdSnippet || '') + (it.jdSnippet && it.jdSnippet.length >= 89 ? '…' : '')),
      el('p', { className: 'muted small' }, new Date(it.when).toLocaleString())
    );
    list.appendChild(card);
  });
}
$('#clearHist').addEventListener('click', () => {
  try { localStorage.removeItem('ff-history'); } catch (_) {}
  renderHistory();
  toast('History cleared.');
});
renderHistory();

// =========================================================
// 21. SAMPLE DATA
// =========================================================
const SAMPLE_CV = `Alex Morgan
alex.morgan@example.com · (555) 123-4567 · linkedin.com/in/alexmorgan · github.com/alexmorgan

SUMMARY
Senior data analyst with 6 years of experience turning operational data into measurable outcomes for SaaS and fintech teams. Skilled in SQL, Python, and dashboarding; comfortable partnering with product and engineering to ship analytics that drive decisions.

EXPERIENCE
Senior Data Analyst — Brightline SaaS (2021 – Present)
- Responsible for the executive KPI dashboard used by 80+ stakeholders weekly.
- Worked on churn analysis that informed a pricing change projected to retain $2.4M ARR.
- Built ETL pipelines in Python and Airflow feeding Snowflake; reduced reporting latency from 24h to 45m.
- Helped onboard two junior analysts and instituted a peer-review process for SQL.

Data Analyst — Northwave Bank (2018 – 2021)
- Delivered fraud-detection reporting that cut false positives by 22%.
- Automated weekly regulatory reports using SQL Server and Python (pandas).
- Collaborated with risk and compliance teams to define KPI definitions.

EDUCATION
B.S. Statistics, University of Washington — 2018

SKILLS
SQL, Python, pandas, Airflow, Snowflake, Tableau, Power BI, Git, A/B testing, statistics

CERTIFICATIONS
Google Data Analytics Certificate — 2021`;

const SAMPLE_TRANSCRIPT = `University of Washington — Official Transcript
Degree: Bachelor of Science, Statistics
GPA: 3.72

Selected Coursework:
- STAT 311 Probability and Statistics
- STAT 421 Statistical Inference
- STAT 435 Introduction to Statistical Machine Learning
- CSE 142 Intro to Programming (Python)
- CSE 160 Data Programming
- INFO 340 Client-side Web Development
- ECON 424 Applied Statistics & Econometrics
- MATH 308 Linear Algebra

Honors: Dean's List (4 quarters)`;

const SAMPLE_JD = `Senior Data Analyst — Fintech Platform

About the role
We are looking for a Senior Data Analyst at Northstar Finance to partner with product, risk, and engineering teams. You will own analytics for our lending platform, build executive dashboards, and surface insights that improve approval rates and reduce defaults.

Responsibilities
- Build and maintain dashboards in Tableau or Looker for executive stakeholders.
- Design ETL pipelines using Python, SQL, and Airflow against Snowflake.
- Run A/B tests on pricing, underwriting, and onboarding funnels.
- Mentor junior analysts and codify SQL best practices.
- Partner with data engineering on data quality and modeling (dbt a plus).

Required qualifications
- Bachelor's degree in Statistics, Mathematics, Economics, Computer Science, or related field.
- At least 5 years of experience as a data or analytics professional.
- Strong SQL skills and Python (pandas, numpy).
- Experience with Snowflake, Airflow, and a BI tool (Tableau, Looker, or Power BI).
- Solid grounding in statistics, A/B testing, and experimentation.

Preferred qualifications
- Experience in fintech, banking, or lending.
- Familiarity with dbt and modern data stack practices.
- Exposure to machine learning concepts.

Northstar Finance is committed to building a diverse, equitable, and inclusive team.`;

$('#loadSampleBtn').addEventListener('click', () => {
  state.cv = { name: 'sample_cv.txt', text: SAMPLE_CV };
  state.transcript = { name: 'sample_transcript.txt', text: SAMPLE_TRANSCRIPT };
  state.jd = SAMPLE_JD;
  jdTextEl.value = SAMPLE_JD;
  $('#jdCount').textContent = `${wordCount(SAMPLE_JD)} words`;

  const cvStatus = $('#cvStatus');
  cvStatus.hidden = false; cvStatus.classList.remove('err');
  cvStatus.innerHTML = '';
  cvStatus.append(
    svgCheck(),
    el('span', {}, el('strong', {}, 'sample_cv.txt'), ` · ${formatBytes(SAMPLE_CV.length)} · ${wordCount(SAMPLE_CV)} words`),
    (() => { const b = el('button', { className: 'link-btn' }, 'Remove'); b.addEventListener('click', () => clearUpload('cv', cvStatus)); return b; })()
  );

  const trStatus = $('#trStatus');
  trStatus.hidden = false; trStatus.classList.remove('err');
  trStatus.innerHTML = '';
  trStatus.append(
    svgCheck(),
    el('span', {}, el('strong', {}, 'sample_transcript.txt'), ` · ${formatBytes(SAMPLE_TRANSCRIPT.length)} · ${wordCount(SAMPLE_TRANSCRIPT)} words`),
    (() => { const b = el('button', { className: 'link-btn' }, 'Remove'); b.addEventListener('click', () => clearUpload('transcript', trStatus)); return b; })()
  );

  updateStepper(); updateAnalyzeButton();
  toast('Sample data loaded — click "Analyze Match".');
});

// =========================================================
// 22. UTILITY
// =========================================================
function clamp(v, lo, hi) { return Math.max(lo, Math.min(hi, v)); }
function wordCount(t) { return ((t || '').trim().match(/\S+/g) || []).length; }
function escapeHtml(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  }[c]));
}
function escapeRegex(s) { return String(s || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
function capitalize(s) { return s ? s.charAt(0).toUpperCase() + s.slice(1) : ''; }
function titleCase(s) { return String(s || '').split(/\s+/).map(w => w ? w[0].toUpperCase() + w.slice(1) : '').join(' '); }
function formatBytes(n) {
  if (n < 1024) return n + ' B';
  if (n < 1024 * 1024) return (n / 1024).toFixed(1) + ' KB';
  return (n / 1024 / 1024).toFixed(1) + ' MB';
}
// Stable, small hash → small int
function quickHash(str) {
  let h = 0; const s = String(str || '');
  for (let i = 0; i < s.length; i++) {
    h = ((h << 5) - h + s.charCodeAt(i)) | 0;
  }
  return Math.abs(h);
}

// =========================================================
// 23. INITIAL UI STATE
// =========================================================
updateStepper(); updateAnalyzeButton();

// Surface a friendly error if the user is on an unsupported browser
try {
  // feature gates
  if (typeof Promise === 'undefined' || typeof Map === 'undefined' || typeof Set === 'undefined') {
    toast('This browser is missing required features. Please upgrade.', true);
  }
} catch (_) {}

})();
