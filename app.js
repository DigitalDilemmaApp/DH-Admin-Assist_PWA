/* ========================================================
   Class Visit PWA – Application Logic
   ======================================================== */
'use strict';

/* ── Constants ──────────────────────────────────────────── */
const STORAGE_KEY_SETTINGS = 'hps-cv-settings';
const STORAGE_KEY_FORMS    = 'hps-cv-forms';
const TOGGLE_FIELDS = {
  2: ['f2-lesson-obs','f2-classroom-obs','f2-learner-books','f2-educator-file','f2-lesson-file','f2-curriculum','f2-assessment-file'],
  3: ['f3-work-uptodate','f3-books-marked','f3-books-signed','f3-corrections','f3-books-neat','f3-daily-writing','f3-weekly-tests','f3-progression','f3-differentiated']
};
const TOGGLE_LABELS = {
  'f2-lesson-obs':       'Lesson Observation',
  'f2-classroom-obs':    'Classroom Observation',
  'f2-learner-books':    'Learners Book Moderation',
  'f2-educator-file':    'Educator File Moderation',
  'f2-lesson-file':      'Lesson File',
  'f2-curriculum':       'Tracking of Curriculum',
  'f2-assessment-file':  'Assessment File',
  'f3-work-uptodate':    'Is the work up to date?',
  'f3-books-marked':     'Are the books properly marked up to date?',
  'f3-books-signed':     'Are the books regularly signed and dated?',
  'f3-corrections':      'Are the corrections done and checked?',
  'f3-books-neat':       'Are the books neat and reasonably maintained?',
  'f3-daily-writing':    'Are there daily writing activities?',
  'f3-weekly-tests':     'Are there weekly spelling/mental/speed tests?',
  'f3-progression':      'Is Progression evident?',
  'f3-differentiated':   'Are there differentiated activities for intervention/extension?'
};

/* ── App State ──────────────────────────────────────────── */
const state = {
  currentStep: 1,
  totalSteps: 3,
  settings: {
    schoolName: '',
    dhName: '',
    logoDataUrl: null,
    firstLaunch: true
  },
  toggleValues: {},   /* field → 'YES' | 'NO' | null */
  signedCanvases: new Set(),
  learnerPhotos: {}   /* 'f3-l1' | 'f3-l2' | 'f3-l3' → dataUrl | null */
};

/* ── Initialisation ─────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', () => {
  registerSW();
  loadSettings();
  updateHeader();
  setupSettings();
  setupToggleSwitches();
  setupNavigation();
  setupSectionActions();
  setupImportJSON();
  setupLearnerCameras();
  restoreFormData();
  applySettingsToForms();
  setupAutosave();
  generateAppleTouchIcon();

  if (state.settings.firstLaunch) {
    openSettings();
  }
});

/* ── Service Worker ─────────────────────────────────────── */
function registerSW() {
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js').catch(() => {});
  }
}

/* ── Apple Touch Icon (dynamic) ─────────────────────────── */
function generateAppleTouchIcon() {
  try {
    const c = document.createElement('canvas');
    c.width = 180; c.height = 180;
    const ctx = c.getContext('2d');
    ctx.fillStyle = '#1a5276';
    ctx.beginPath();
    ctx.roundRect(0, 0, 180, 180, 28);
    ctx.fill();
    ctx.fillStyle = '#ffffff';
    ctx.font = 'bold 64px -apple-system, Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText('HPS', 90, 90);
    const link = document.getElementById('apple-touch-icon-link') || document.createElement('link');
    link.rel = 'apple-touch-icon';
    link.href = c.toDataURL('image/png');
    if (!document.getElementById('apple-touch-icon-link')) document.head.appendChild(link);
  } catch {}
}

/* ── Settings ───────────────────────────────────────────── */
function loadSettings() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_SETTINGS);
    if (raw) {
      const saved = JSON.parse(raw);
      state.settings = { ...state.settings, ...saved };
    }
  } catch {}
}

function persistSettings() {
  try {
    localStorage.setItem(STORAGE_KEY_SETTINGS, JSON.stringify(state.settings));
  } catch {}
}

function updateHeader() {
  const nameEl  = document.getElementById('header-school-name');
  const dhEl    = document.getElementById('header-dh-name');
  const logoEl  = document.getElementById('header-logo');
  const name    = state.settings.schoolName || '';
  const dh      = state.settings.dhName || '';

  nameEl.textContent = name;

  if (dh) {
    dhEl.textContent = dh;
    dhEl.classList.remove('hidden');
  } else {
    dhEl.classList.add('hidden');
  }

  if (state.settings.logoDataUrl) {
    logoEl.src = state.settings.logoDataUrl;
    logoEl.classList.remove('hidden');
  } else {
    logoEl.classList.add('hidden');
  }

  /* Keep data-school on form sections for print CSS ::before */
  document.querySelectorAll('.form-section').forEach(s => {
    s.setAttribute('data-school', name);
  });
}

function setupSettings() {
  const overlay     = document.getElementById('settings-overlay');
  const nameInput   = document.getElementById('setting-school-name');
  const dhInput     = document.getElementById('setting-dh-name');
  const logoInput   = document.getElementById('setting-logo-input');
  const preview     = document.getElementById('logo-preview');
  const placeholder = document.getElementById('logo-upload-placeholder');
  const removeLogo  = document.getElementById('remove-logo-btn');

  /* Open */
  document.getElementById('settings-btn').addEventListener('click', openSettings);

  /* Close buttons */
  ['cancel-settings-btn', 'cancel-settings-btn-2'].forEach(id => {
    document.getElementById(id).addEventListener('click', closeSettings);
  });

  overlay.addEventListener('click', e => { if (e.target === overlay) closeSettings(); });

  /* Logo upload area keyboard */
  document.getElementById('logo-upload-area').addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') logoInput.click();
  });

  /* Logo file selection */
  logoInput.addEventListener('change', e => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      const url = ev.target.result;
      preview.src = url;
      preview.classList.remove('hidden');
      placeholder.style.display = 'none';
      removeLogo.classList.remove('hidden');
      /* Temp store; committed on Save */
      logoInput._pendingUrl = url;
    };
    reader.readAsDataURL(file);
  });

  /* Remove logo */
  removeLogo.addEventListener('click', () => {
    preview.src = '';
    preview.classList.add('hidden');
    placeholder.style.display = '';
    removeLogo.classList.add('hidden');
    logoInput.value = '';
    logoInput._pendingUrl = null;
    state.settings.logoDataUrl = null;
  });

  /* Save */
  document.getElementById('save-settings-btn').addEventListener('click', () => {
    state.settings.schoolName  = nameInput.value.trim();
    state.settings.dhName      = dhInput.value.trim();
    state.settings.firstLaunch = false;

    if (logoInput._pendingUrl !== undefined) {
      state.settings.logoDataUrl = logoInput._pendingUrl;
    }

    persistSettings();
    updateHeader();
    applySettingsToForms();
    closeSettings();
  });
}

function openSettings() {
  const overlay     = document.getElementById('settings-overlay');
  const nameInput   = document.getElementById('setting-school-name');
  const dhInput     = document.getElementById('setting-dh-name');
  const preview     = document.getElementById('logo-preview');
  const placeholder = document.getElementById('logo-upload-placeholder');
  const removeLogo  = document.getElementById('remove-logo-btn');
  const logoInput   = document.getElementById('setting-logo-input');

  nameInput.value = state.settings.schoolName || '';
  dhInput.value   = state.settings.dhName     || '';

  if (state.settings.logoDataUrl) {
    preview.src = state.settings.logoDataUrl;
    preview.classList.remove('hidden');
    placeholder.style.display = 'none';
    removeLogo.classList.remove('hidden');
  } else {
    preview.classList.add('hidden');
    placeholder.style.display = '';
    removeLogo.classList.add('hidden');
  }
  logoInput._pendingUrl = undefined; /* reset pending */

  overlay.classList.remove('hidden');
  nameInput.focus();
}

function closeSettings() {
  document.getElementById('settings-overlay').classList.add('hidden');
}

/* Pre-fill form fields from settings when the fields are currently empty */
function applySettingsToForms() {
  const dh = state.settings.dhName;
  if (dh && !val('f1-designation')) setVal('f1-designation', dh);
}

/* ── Navigation ─────────────────────────────────────────── */
const FORM_TITLES = {
  1: 'Notice of Class Visit',
  2: 'Classroom Observation',
  3: 'Book Moderation',
  4: 'Complete'
};

function setupNavigation() {
  document.getElementById('back-btn').addEventListener('click', handleBack);
  document.getElementById('next-btn').addEventListener('click', handleNext);
  document.getElementById('download-pdf-btn').addEventListener('click', downloadPDF);
  document.getElementById('download-pdf-json-btn').addEventListener('click', downloadPDFAndJSON);
  document.getElementById('download-json-btn').addEventListener('click', downloadJSON);
  document.getElementById('print-btn').addEventListener('click', printForms);
  document.getElementById('start-over-btn').addEventListener('click', startOver);
}

function handleNext() {
  if (state.currentStep < state.totalSteps) {
    const nextStep = state.currentStep + 1;
    autofill(nextStep);
    navigateTo(nextStep);
  } else {
    /* On step 3 – submit */
    submitForms();
  }
}

function handleBack() {
  if (state.currentStep > 1) {
    navigateTo(state.currentStep - 1);
  }
}

function autofill(toStep) {
  if (toStep === 2) {
    const name = val('f1-educator-name');
    const date = val('f1-date');
    if (name && !val('f2-educator-name')) setVal('f2-educator-name', name);
    if (date && !val('f2-date'))          setVal('f2-date', date);
  }
  if (toStep === 3) {
    const name = val('f1-educator-name');
    const date = val('f1-date');
    if (name && !val('f3-educator-name')) setVal('f3-educator-name', name);
    if (date && !val('f3-date'))          setVal('f3-date', date);
  }
}

function navigateTo(step) {
  /* Hide all sections */
  document.querySelectorAll('.form-section').forEach(s => s.classList.remove('active'));

  const target = document.getElementById('form-' + step);
  if (target) target.classList.add('active');

  /* Init canvases in this section (lazy – needs visible dimensions) */
  if (target) {
    target.querySelectorAll('.sig-canvas').forEach(canvas => {
      if (!canvas._initialized) initCanvas(canvas);
    });
  }

  state.currentStep = step;
  updateNavUI();
  document.getElementById('main-content').scrollTo({ top: 0, behavior: 'smooth' });
}

function submitForms() {
  saveFormData();
  /* Show completion */
  document.querySelectorAll('.form-section').forEach(s => s.classList.remove('active'));
  document.getElementById('completion-screen').classList.add('active');
  document.getElementById('nav-footer').style.display = 'none';

  /* Mark all dots done */
  document.querySelectorAll('.step-dot').forEach(d => d.classList.add('done'));
}

function updateNavUI() {
  const backBtn  = document.getElementById('back-btn');
  const nextBtn  = document.getElementById('next-btn');
  const titleEl  = document.getElementById('form-title-label');

  backBtn.disabled = state.currentStep <= 1;
  nextBtn.textContent = state.currentStep === state.totalSteps ? 'Submit ›' : 'Next ›';
  titleEl.textContent = FORM_TITLES[state.currentStep] || '';

  /* Update step dots */
  document.querySelectorAll('.step-dot').forEach(dot => {
    const s = parseInt(dot.dataset.step);
    dot.classList.remove('active', 'done');
    if (s < state.currentStep) dot.classList.add('done');
    if (s === state.currentStep) dot.classList.add('active');
  });

  document.getElementById('nav-footer').style.display = '';
}

/* ── Toggle Switches ─────────────────────────────────────── */
function setupToggleSwitches() {
  document.querySelectorAll('.yn-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const field = btn.dataset.field;
      const value = btn.dataset.value;
      setToggle(field, value);
      saveFormData();
    });
  });
}

function setToggle(field, value) {
  /* Toggle off if already selected */
  const current = state.toggleValues[field];
  const newVal  = current === value ? null : value;
  state.toggleValues[field] = newVal;

  /* Update button visuals */
  document.querySelectorAll(`.yn-btn[data-field="${field}"]`).forEach(btn => {
    btn.classList.remove('active-yes', 'active-no');
    if (btn.dataset.value === newVal) {
      btn.classList.add(newVal === 'YES' ? 'active-yes' : 'active-no');
    }
  });

  /* Show comment field only on NO */
  const wrap = document.getElementById(`${field}-comment-wrap`);
  if (wrap) {
    if (newVal === 'NO') {
      wrap.classList.remove('hidden');
      wrap.querySelector('textarea')?.focus();
    } else {
      wrap.classList.add('hidden');
    }
  }
}

function restoreToggles(saved) {
  if (!saved) return;
  Object.entries(saved).forEach(([field, value]) => {
    if (value) setToggle(field, value);
  });
}

function restoreToggleComments(saved) {
  if (!saved) return;
  Object.entries(saved).forEach(([field, text]) => {
    const el = document.getElementById(`${field}-comment`);
    if (el && text) el.value = text;
  });
}

/* ── Signature Canvas ─────────────────────────────────────── */
function initCanvas(canvas) {
  const ctx = canvas.getContext('2d');
  const dpr = window.devicePixelRatio || 1;
  const cssW = canvas.offsetWidth  || canvas.parentElement.offsetWidth || 300;
  const cssH = parseInt(getComputedStyle(canvas).height) || 130;

  canvas.width  = cssW * dpr;
  canvas.height = cssH * dpr;
  canvas.style.width  = cssW + 'px';
  canvas.style.height = cssH + 'px';
  ctx.scale(dpr, dpr);

  drawCanvasPlaceholder(canvas, ctx, cssW, cssH);

  let isDrawing = false;
  let lastX = 0, lastY = 0;

  function getXY(e) {
    const r = canvas.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top, pressure: e.pressure || 0.5 };
  }

  function onDown(e) {
    e.preventDefault();
    if (!state.signedCanvases.has(canvas.id)) {
      /* Clear placeholder on first stroke */
      const w = canvas.width / dpr;
      const h = canvas.height / dpr;
      ctx.clearRect(0, 0, w, h);
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, w, h);
      drawSignLine(ctx, w, h);
      state.signedCanvases.add(canvas.id);
    }
    const p = getXY(e);
    lastX = p.x; lastY = p.y;
    isDrawing = true;
    canvas.setPointerCapture(e.pointerId);
  }

  function onMove(e) {
    if (!isDrawing) return;
    e.preventDefault();
    const p = getXY(e);
    const pressure = (e.pointerType === 'pen') ? Math.max(0.2, e.pressure) : 0.5;

    ctx.beginPath();
    ctx.moveTo(lastX, lastY);
    ctx.lineTo(p.x, p.y);
    ctx.strokeStyle = '#1a1a2e';
    ctx.lineWidth   = e.pointerType === 'pen' ? pressure * 3.5 : 2.5;
    ctx.lineCap     = 'round';
    ctx.lineJoin    = 'round';
    ctx.stroke();

    lastX = p.x; lastY = p.y;
  }

  function onUp(e) { isDrawing = false; saveFormData(); }

  canvas.addEventListener('pointerdown',   onDown);
  canvas.addEventListener('pointermove',   onMove);
  canvas.addEventListener('pointerup',     onUp);
  canvas.addEventListener('pointercancel', onUp);

  /* Clear button */
  const btn = document.querySelector(`.clear-sig-btn[data-target="${canvas.id}"]`);
  if (btn) {
    btn.addEventListener('click', () => clearCanvas(canvas));
  }

  canvas._initialized = true;
  canvas._ctx = ctx;
  canvas._dpr = dpr;
}

function drawCanvasPlaceholder(canvas, ctx, w, h) {
  ctx.fillStyle = '#fafafa';
  ctx.fillRect(0, 0, w, h);
  drawSignLine(ctx, w, h);
  ctx.fillStyle = '#c8cdd3';
  ctx.font = '14px -apple-system, Arial, sans-serif';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText('Sign here', w / 2, h * 0.38);
}

function drawSignLine(ctx, w, h) {
  ctx.save();
  ctx.strokeStyle = '#d5d8dc';
  ctx.lineWidth   = 1;
  ctx.setLineDash([6, 4]);
  ctx.beginPath();
  ctx.moveTo(16, h * 0.72);
  ctx.lineTo(w - 16, h * 0.72);
  ctx.stroke();
  ctx.setLineDash([]);
  ctx.restore();
}

function clearCanvas(canvas) {
  const ctx = canvas._ctx || canvas.getContext('2d');
  const dpr = canvas._dpr || window.devicePixelRatio || 1;
  const w   = canvas.width  / dpr;
  const h   = canvas.height / dpr;
  ctx.clearRect(0, 0, w, h);
  state.signedCanvases.delete(canvas.id);
  drawCanvasPlaceholder(canvas, ctx, w, h);
  saveFormData();
}

function getCanvasDataUrl(canvasId) {
  if (!state.signedCanvases.has(canvasId)) return null;
  const c = document.getElementById(canvasId);
  return c ? c.toDataURL('image/png') : null;
}

function getPhotoDataUrl(key) {
  return state.learnerPhotos[key] || null;
}

/* ── Form Data Persistence ──────────────────────────────── */
function collectToggleComments(fields) {
  const result = {};
  fields.forEach(f => {
    const el = document.getElementById(`${f}-comment`);
    if (el) result[f] = el.value;
  });
  return result;
}

function collectFormData() {
  return {
    f1: {
      noticeDate:   val('f1-notice-date'),
      educatorName: val('f1-educator-name'),
      grade:        val('f1-grade'),
      date:         val('f1-date'),
      time:         val('f1-time'),
      purposes: {
        lessonObservation:     checked('f1-lesson-obs'),
        classroomObservation:  checked('f1-classroom-obs'),
        learnerBookModeration: checked('f1-learner-books'),
        educatorFileModeration:checked('f1-educator-file'),
        curriculumTracking:    checked('f1-curriculum')
      },
      designation: val('f1-designation')
    },
    f2: {
      educatorName:    val('f2-educator-name'),
      date:            val('f2-date'),
      toggles:         { ...state.toggleValues },
      toggleComments:  collectToggleComments(TOGGLE_FIELDS[2]),
      comments:        val('f2-comments'),
      signatures: {
        moderator: { dataUrl: getCanvasDataUrl('sig-f2-moderator'), date: val('sig-f2-moderator-date') },
        deputy:    { dataUrl: getCanvasDataUrl('sig-f2-deputy'),    date: val('sig-f2-deputy-date')    },
        teacher:   { dataUrl: getCanvasDataUrl('sig-f2-teacher'),   date: val('sig-f2-teacher-date')   }
      }
    },
    f3: {
      educatorName:    val('f3-educator-name'),
      date:            val('f3-date'),
      toggles:         { ...state.toggleValues },
      toggleComments:  collectToggleComments(TOGGLE_FIELDS[3]),
      comments:        '',
      learners: [
        { name: val('f3-l1-name'), comment: val('f3-l1-comment'), photoDataUrl: getPhotoDataUrl('f3-l1') },
        { name: val('f3-l2-name'), comment: val('f3-l2-comment'), photoDataUrl: getPhotoDataUrl('f3-l2') },
        { name: val('f3-l3-name'), comment: val('f3-l3-comment'), photoDataUrl: getPhotoDataUrl('f3-l3') }
      ],
      signatures: {
        moderator: { dataUrl: getCanvasDataUrl('sig-f3-moderator'), date: val('sig-f3-moderator-date') },
        deputy:    { dataUrl: getCanvasDataUrl('sig-f3-deputy'),    date: val('sig-f3-deputy-date')    }
      }
    }
  };
}

function saveFormData() {
  try {
    const data = collectFormData();
    localStorage.setItem(STORAGE_KEY_FORMS, JSON.stringify(data));
  } catch {}
}

function applyFormData(data) {
  /* Form 1 */
  const f1 = data.f1 || {};
  setVal('f1-notice-date',   f1.noticeDate);
  setVal('f1-educator-name', f1.educatorName);
  setVal('f1-grade',         f1.grade);
  setVal('f1-date',          f1.date);
  setVal('f1-time',          f1.time);
  setVal('f1-designation',   f1.designation);
  if (f1.purposes) {
    setChecked('f1-lesson-obs',    f1.purposes.lessonObservation);
    setChecked('f1-classroom-obs', f1.purposes.classroomObservation);
    setChecked('f1-learner-books', f1.purposes.learnerBookModeration);
    setChecked('f1-educator-file', f1.purposes.educatorFileModeration);
    setChecked('f1-curriculum',    f1.purposes.curriculumTracking);
  }

  /* Form 2 */
  const f2 = data.f2 || {};
  setVal('f2-educator-name', f2.educatorName);
  setVal('f2-date',          f2.date);
  setVal('f2-comments',      f2.comments);
  if (f2.signatures) {
    setVal('sig-f2-moderator-date', f2.signatures.moderator?.date);
    setVal('sig-f2-deputy-date',    f2.signatures.deputy?.date);
    setVal('sig-f2-teacher-date',   f2.signatures.teacher?.date);
  }

  /* Form 3 */
  const f3 = data.f3 || {};
  setVal('f3-educator-name', f3.educatorName);
  setVal('f3-date',          f3.date);
  if (f3.learners) {
    f3.learners.forEach((l, i) => {
      const n = i + 1;
      setVal(`f3-l${n}-name`,    l.name);
      setVal(`f3-l${n}-comment`, l.comment);
      if (l.photoDataUrl) {
        state.learnerPhotos[`f3-l${n}`] = l.photoDataUrl;
        const thumb     = document.getElementById(`f3-l${n}-photo-thumb`);
        const thumbWrap = document.getElementById(`f3-l${n}-photo-thumb-wrap`);
        const cameraBtn = document.getElementById(`f3-l${n}-camera-btn`);
        if (thumb)     thumb.src = l.photoDataUrl;
        if (thumbWrap) thumbWrap.classList.remove('hidden');
        if (cameraBtn) cameraBtn.classList.add('hidden');
      }
    });
  }
  if (f3.signatures) {
    setVal('sig-f3-moderator-date', f3.signatures.moderator?.date);
    setVal('sig-f3-deputy-date',    f3.signatures.deputy?.date);
  }

  /* Restore toggles – merge f2 and f3 toggles */
  const allToggles = { ...(f2.toggles || {}), ...(f3.toggles || {}) };
  restoreToggles(allToggles);

  /* Restore toggle comments */
  restoreToggleComments(f2.toggleComments);
  restoreToggleComments(f3.toggleComments);

  /* Restore signature images to canvases (best-effort after layout) */
  requestAnimationFrame(() => restoreSignatureImages(data));
}

function restoreFormData() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_FORMS);
    if (!raw) return;
    applyFormData(JSON.parse(raw));
  } catch {}
}

function restoreSignatureImages(data) {
  const sigMap = {
    'sig-f2-moderator': data.f2?.signatures?.moderator?.dataUrl,
    'sig-f2-deputy':    data.f2?.signatures?.deputy?.dataUrl,
    'sig-f2-teacher':   data.f2?.signatures?.teacher?.dataUrl,
    'sig-f3-moderator': data.f3?.signatures?.moderator?.dataUrl,
    'sig-f3-deputy':    data.f3?.signatures?.deputy?.dataUrl
  };

  Object.entries(sigMap).forEach(([id, dataUrl]) => {
    if (!dataUrl) return;
    const canvas = document.getElementById(id);
    if (!canvas) return;

    /* Ensure canvas is initialized – show its parent section temporarily */
    const section = canvas.closest('.form-section');
    const wasHidden = !section.classList.contains('active');
    if (wasHidden) section.style.display = 'block';

    if (!canvas._initialized) initCanvas(canvas);

    if (wasHidden) section.style.display = '';

    const ctx = canvas._ctx || canvas.getContext('2d');
    const dpr = canvas._dpr || window.devicePixelRatio || 1;
    const w   = canvas.width / dpr;
    const h   = canvas.height / dpr;
    const img = new Image();
    img.onload = () => {
      ctx.clearRect(0, 0, w, h);
      ctx.drawImage(img, 0, 0, w, h);
      state.signedCanvases.add(id);
    };
    img.src = dataUrl;
  });
}

function setupAutosave() {
  /* Debounced save on all input/change events */
  let timer;
  document.addEventListener('input',  () => { clearTimeout(timer); timer = setTimeout(saveFormData, 400); });
  document.addEventListener('change', () => { clearTimeout(timer); timer = setTimeout(saveFormData, 400); });
}

/* ── Start Over ──────────────────────────────────────────── */
function startOver() {
  if (!confirm('Clear all forms and start a new visit?')) return;
  localStorage.removeItem(STORAGE_KEY_FORMS);
  state.toggleValues   = {};
  state.signedCanvases = new Set();

  /* Clear all inputs */
  document.querySelectorAll('.form-section input, .form-section textarea').forEach(el => {
    if (el.type === 'checkbox') el.checked = false;
    else el.value = '';
  });

  /* Reset all toggles and hide any open comment fields */
  document.querySelectorAll('.yn-btn').forEach(btn => btn.classList.remove('active-yes','active-no'));
  document.querySelectorAll('.toggle-comment-wrap').forEach(w => w.classList.add('hidden'));
  document.querySelectorAll('.toggle-comment').forEach(t => { t.value = ''; });

  /* Clear all canvases */
  document.querySelectorAll('.sig-canvas').forEach(c => {
    if (c._initialized) clearCanvas(c);
  });

  /* Clear learner photos */
  state.learnerPhotos = {};
  [1, 2, 3].forEach(n => {
    const thumb     = document.getElementById(`f3-l${n}-photo-thumb`);
    const thumbWrap = document.getElementById(`f3-l${n}-photo-thumb-wrap`);
    const cameraBtn = document.getElementById(`f3-l${n}-camera-btn`);
    if (thumb)     { thumb.src = ''; }
    if (thumbWrap) { thumbWrap.classList.add('hidden'); }
    if (cameraBtn) { cameraBtn.classList.remove('hidden'); }
  });

  document.getElementById('nav-footer').style.display = '';
  navigateTo(1);
  applySettingsToForms();
}

/* ── PDF Generation ──────────────────────────────────────── */
async function downloadPDF() {
  const jsPDF = window.jspdf?.jsPDF;
  if (!jsPDF) {
    alert('PDF library is not loaded yet. Please connect to the internet once to cache it, or use the Print button to save as PDF.');
    return;
  }

  const btn = document.getElementById('download-pdf-btn');
  btn.disabled = true;
  btn.textContent = 'Generating…';

  try {
    const data = collectFormData();
    const doc  = buildDoc(jsPDF, data, state.settings);
    const name = (data.f1.educatorName || 'class-visit').replace(/\s+/g,'-');
    const date = data.f1.date || new Date().toISOString().slice(0,10);
    doc.save(`${name}-${date}.pdf`);
  } catch (err) {
    console.error(err);
    alert('PDF generation failed. Please use the Print button instead.');
  } finally {
    btn.disabled = false;
    btn.innerHTML = '<span class="btn-icon">⬇</span> Save as PDF';
  }
}

function buildDoc(jsPDF, data, settings, sections = [1, 2, 3]) {
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  const PW = 210, PH = 297, M = 14, CW = PW - M * 2;
  const BLUE = [26, 82, 118], LBLUE = [174, 214, 241], WHITE = [255,255,255];
  const GREEN = [30, 132, 73], RED = [192, 57, 43];
  const GREY = [127, 140, 141], DARK = [44, 62, 80], LGREY = [244, 246, 248];

  /* ---- helpers ---- */
  function pageHeader(title) {
    let y = M;
    if (settings.logoDataUrl) {
      try {
        doc.addImage(settings.logoDataUrl, 'PNG', M, y, 22, 18);
        doc.setFont('helvetica','bold').setFontSize(16).setTextColor(...BLUE);
        doc.text(settings.schoolName, M + 26, y + 7);
        doc.setFont('helvetica','normal').setFontSize(11).setTextColor(...GREY);
        doc.text(title, M + 26, y + 14);
        y += 24;
      } catch {
        y = centeredHeader(title, y);
      }
    } else {
      y = centeredHeader(title, y);
    }
    doc.setDrawColor(...BLUE).setLineWidth(0.5).line(M, y, PW - M, y);
    return y + 5;
  }

  function centeredHeader(title, y) {
    doc.setFont('helvetica','bold').setFontSize(17).setTextColor(...BLUE);
    doc.text(settings.schoolName, PW/2, y + 8, { align:'center' });
    y += 13;
    doc.setFont('helvetica','normal').setFontSize(12).setTextColor(...GREY);
    doc.text(title, PW/2, y, { align:'center' });
    return y + 8;
  }

  function field(label, value, x, y, w) {
    doc.setFont('helvetica','bold').setFontSize(8).setTextColor(...GREY);
    doc.text(label.toUpperCase(), x, y);
    doc.setFont('helvetica','normal').setFontSize(10).setTextColor(...DARK);
    const disp = value || '—';
    doc.text(disp, x, y + 5.5);
    doc.setDrawColor(...LGREY).setLineWidth(0.2).line(x, y + 7, x + w, y + 7);
    return y + 12;
  }

  function checkItem(label, isChecked, x, y) {
    if (isChecked) { doc.setFillColor(...BLUE); } else { doc.setFillColor(240, 240, 240); }
    doc.roundedRect(x, y - 3, 5, 5, 1, 1, isChecked ? 'F' : 'S');
    if (isChecked) {
      doc.setTextColor(...WHITE).setFont('helvetica','bold').setFontSize(7);
      doc.text('✓', x + 1, y + 1.2);
    }
    doc.setFont('helvetica','normal').setFontSize(10).setTextColor(...DARK);
    doc.text(label, x + 7.5, y + 1);
    return y + 8;
  }

  function toggleTableHeader(x, y, w) {
    doc.setFillColor(...BLUE).rect(x, y, w, 7, 'F');
    doc.setFont('helvetica','bold').setFontSize(8).setTextColor(...WHITE);
    doc.text('ITEM', x + 2, y + 4.8);
    doc.text('RESPONSE', x + w - 22, y + 4.8);
    return y + 7;
  }

  function toggleRow(label, value, comment, x, y, w, even) {
    if (even) { doc.setFillColor(248,249,250).rect(x, y, w, 8, 'F'); }
    doc.setFont('helvetica','normal').setFontSize(9).setTextColor(...DARK);
    const lines = doc.splitTextToSize(label, w - 30);
    doc.text(lines, x + 2, y + 5);
    const rowH = Math.max(8, lines.length * 5);

    if (value) {
      if (value === 'YES') {
        doc.setFillColor(...GREEN).roundedRect(x + w - 23, y + 1, 22, rowH - 2, 2, 2, 'F');
        doc.setTextColor(...WHITE);
      } else {
        doc.setFillColor(...RED).roundedRect(x + w - 23, y + 1, 22, rowH - 2, 2, 2, 'F');
        doc.setTextColor(...WHITE);
      }
      doc.setFont('helvetica','bold').setFontSize(8);
      doc.text(value, x + w - 12, y + rowH / 2 + 1, { align:'center' });
    } else {
      doc.setDrawColor(200,200,200).setLineWidth(0.2).rect(x + w - 23, y + 1, 22, rowH - 2);
      doc.setFont('helvetica','normal').setFontSize(8).setTextColor(190,190,190);
      doc.text('—', x + w - 12, y + rowH / 2 + 1, { align:'center' });
    }

    doc.setDrawColor(220,220,220).setLineWidth(0.2).line(x, y + rowH, x + w, y + rowH);
    let totalH = rowH;

    /* Inline comment for NO answers */
    if (value === 'NO' && comment) {
      doc.setFillColor(253, 237, 236).rect(x, y + rowH, w, 0.1, 'F'); /* flush join */
      const cLines = doc.splitTextToSize('Comment: ' + comment, w - 6);
      const cH = Math.max(7, cLines.length * 4.5) + 4;
      doc.setFillColor(253, 237, 236).rect(x, y + rowH, w, cH, 'F');
      doc.setFont('helvetica','italic').setFontSize(8).setTextColor(...RED);
      doc.text(cLines, x + 3, y + rowH + 4.5);
      doc.setDrawColor(220,220,220).setLineWidth(0.2).line(x, y + rowH + cH, x + w, y + rowH + cH);
      totalH += cH;
    }

    return y + totalH;
  }

  function sigBlock(label, dataUrl, date, x, y, w, h = 22) {
    doc.setFont('helvetica','bold').setFontSize(9).setTextColor(...GREY);
    doc.text(label.toUpperCase(), x, y);
    y += 3;
    doc.setDrawColor(180,180,180).setLineWidth(0.3).rect(x, y, w, h);
    if (dataUrl) {
      try { doc.addImage(dataUrl, 'PNG', x + 1, y + 1, w - 2, h - 2); } catch {}
    }
    y += h + 2;
    doc.setFont('helvetica','normal').setFontSize(8).setTextColor(...GREY);
    doc.text('Date: ' + (date || '________________'), x, y);
    return y + 6;
  }

  function sectionTitle(title, x, y) {
    doc.setFont('helvetica','bold').setFontSize(11).setTextColor(...BLUE);
    doc.text(title, x, y);
    return y + 6;
  }

  const col1 = M, col2 = M + CW/2 + 4, colW = CW/2 - 4;
  let y = M, newPage = false, ly, ry;

  /* ============================================================
     FORM 1 – Notice of Class Visit
     ============================================================ */
  if (sections.includes(1)) {
    if (newPage) doc.addPage();
    newPage = true;
    y = pageHeader('NOTICE OF CLASS VISIT');
    y = sectionTitle('Visit Details', M, y);

  let leftY = y, rightY = y;

  leftY  = field('Notice Date', fmtDate(data.f1.noticeDate), col1, leftY, colW);
  rightY = field('Educator Name', data.f1.educatorName, col2, rightY, colW);
  leftY  = field('Grade', data.f1.grade, col1, leftY, colW);
  rightY = field('Visit Date', fmtDate(data.f1.date), col2, rightY, colW);
  leftY  = field('Visit Time', data.f1.time, col1, leftY, colW);
  rightY = field('Designation', data.f1.designation, col2, rightY, colW);

  y = Math.max(leftY, rightY) + 4;
  y = sectionTitle('Purpose of Visit', M, y);

  const p = data.f1.purposes;
  const purposes = [
    ['Lesson Observation',      p.lessonObservation],
    ['Classroom Observation',   p.classroomObservation],
    ['Learner Book Moderation', p.learnerBookModeration],
    ['Educator File Moderation',p.educatorFileModeration],
    ['Tracking of Curriculum',  p.curriculumTracking]
  ];
  purposes.forEach(([label, checked]) => {
    y = checkItem(label, checked, M, y);
  });
  } // end Form 1

  /* ============================================================
     FORM 2 – Classroom Observation
     ============================================================ */
  if (sections.includes(2)) {
    if (newPage) doc.addPage();
    newPage = true;
    y = pageHeader('CLASSROOM OBSERVATION');

  /* Fields */
  ly = y; ry = y;
  ly = field('Educator Name', data.f2.educatorName, col1, ly, colW);
  ry = field('Date', fmtDate(data.f2.date), col2, ry, colW);
  y  = Math.max(ly, ry) + 4;

  y = sectionTitle('Observation Checklist', M, y);
  y = toggleTableHeader(M, y, CW);
  TOGGLE_FIELDS[2].forEach((fld, i) => {
    const comment = (data.f2.toggleComments || {})[fld] || '';
    y = toggleRow(TOGGLE_LABELS[fld], data.f2.toggles[fld], comment, M, y, CW, i % 2 === 0);
  });
  y += 6;

  /* Comments */
  y = sectionTitle('General Comments', M, y);
  doc.setDrawColor(180,180,180).setLineWidth(0.3).rect(M, y, CW, 30);
  if (data.f2.comments) {
    doc.setFont('helvetica','normal').setFontSize(9).setTextColor(...DARK);
    const lines = doc.splitTextToSize(data.f2.comments, CW - 4);
    doc.text(lines.slice(0, 10), M + 2, y + 5);
  }
  y += 36;

  /* Check page overflow */
  if (y > PH - 80) { doc.addPage(); y = pageHeader('CLASSROOM OBSERVATION (cont.)'); }

  y = sectionTitle('Signatures', M, y);
  const sigW = (CW - 8) / 3;
  const sigs2 = [
    { label: 'Moderator',       ...data.f2.signatures.moderator },
    { label: 'Deputy Principal',...data.f2.signatures.deputy    },
    { label: 'Teacher',         ...data.f2.signatures.teacher   }
  ];
  let maxSigY = y;
  sigs2.forEach((s, i) => {
    const sx = M + i * (sigW + 4);
    const endY = sigBlock(s.label, s.dataUrl, s.date, sx, y, sigW);
    maxSigY = Math.max(maxSigY, endY);
  });
  } // end Form 2

  /* ============================================================
     FORM 3 – Book Moderation
     ============================================================ */
  if (sections.includes(3)) {
    if (newPage) doc.addPage();
    newPage = true;
    y = pageHeader('BOOK MODERATION');

  ly = y; ry = y;
  ly = field('Educator Name', data.f3.educatorName, col1, ly, colW);
  ry = field('Date', fmtDate(data.f3.date), col2, ry, colW);
  y  = Math.max(ly, ry) + 4;

  y = sectionTitle('Moderation Checklist', M, y);
  y = toggleTableHeader(M, y, CW);
  TOGGLE_FIELDS[3].forEach((fld, i) => {
    const comment = (data.f3.toggleComments || {})[fld] || '';
    y = toggleRow(TOGGLE_LABELS[fld], data.f3.toggles[fld], comment, M, y, CW, i % 2 === 0);
  });
  y += 6;

  /* Learner table */
  if (y > PH - 70) { doc.addPage(); y = pageHeader('BOOK MODERATION (cont.)'); }
  y = sectionTitle('Learner Book Review', M, y);

  /* Table header */
  doc.setFillColor(...BLUE).rect(M, y, CW, 7, 'F');
  doc.setFont('helvetica','bold').setFontSize(8).setTextColor(...WHITE);
  doc.text('LEARNER NAME', M + 2, y + 4.8);
  doc.text('COMMENT', M + CW*0.42, y + 4.8);
  y += 7;

  data.f3.learners.forEach((l, i) => {
    const rowH = 9;
    if (i % 2 === 0) { doc.setFillColor(248,249,250).rect(M, y, CW, rowH, 'F'); }
    doc.setFont('helvetica','normal').setFontSize(9).setTextColor(...DARK);
    doc.text(l.name || '', M + 2, y + 5.5);
    doc.text(l.comment || '', M + CW * 0.42, y + 5.5);
    doc.setDrawColor(220,220,220).setLineWidth(0.2).line(M, y + rowH, M + CW, y + rowH);
    y += rowH;

    if (l.photoDataUrl) {
      const photoRowH = 36;
      if (y + photoRowH > PH - M) { doc.addPage(); y = pageHeader('BOOK MODERATION (cont.)'); }
      doc.setFillColor(240, 248, 255).rect(M, y, CW, photoRowH, 'F');
      try {
        const fmt = l.photoDataUrl.startsWith('data:image/png') ? 'PNG' : 'JPEG';
        doc.addImage(l.photoDataUrl, fmt, M + 2, y + 2, 46, photoRowH - 4);
      } catch {}
      doc.setFont('helvetica','italic').setFontSize(7.5).setTextColor(...GREY);
      doc.text(`Evidence photo – ${l.name || 'Learner ' + (i + 1)}`, M + 52, y + 7);
      doc.setDrawColor(220,220,220).setLineWidth(0.2).line(M, y + photoRowH, M + CW, y + photoRowH);
      y += photoRowH;
    }
  });
  y += 6;

  /* Signatures */
  if (y > PH - 70) { doc.addPage(); y = pageHeader('BOOK MODERATION (cont.)'); }
  y = sectionTitle('Signatures', M, y);
  const sigW3 = (CW - 4) / 2;
  const sigs3 = [
    { label: 'Moderator',       ...data.f3.signatures.moderator },
    { label: 'Deputy Principal',...data.f3.signatures.deputy    }
  ];
  sigs3.forEach((s, i) => {
    const sx = M + i * (sigW3 + 4);
    sigBlock(s.label, s.dataUrl, s.date, sx, y, sigW3);
  });
  } // end Form 3

  return doc;
}

/* ── Per-section Print & PDF ─────────────────────────────── */
function setupSectionActions() {
  [1, 2, 3].forEach(n => {
    document.getElementById(`section-print-btn-${n}`).addEventListener('click', () => printSection(n));
    document.getElementById(`section-pdf-btn-${n}`).addEventListener('click', () => downloadSectionPDF(n));
  });
}

function printSection(sectionNum) {
  /* Ensure canvases in the target section are initialised */
  const section = document.getElementById('form-' + sectionNum);
  section.querySelectorAll('.sig-canvas').forEach(canvas => {
    if (!canvas._initialized) {
      const wasHidden = !section.classList.contains('active');
      if (wasHidden) section.style.display = 'block';
      initCanvas(canvas);
      if (wasHidden) section.style.display = '';
    }
  });

  document.body.dataset.printSection = String(sectionNum);
  let cleaned = false;
  const cleanup = () => {
    if (cleaned) return;
    cleaned = true;
    delete document.body.dataset.printSection;
    window.removeEventListener('afterprint', cleanup);
  };
  window.addEventListener('afterprint', cleanup);
  window.print();
  setTimeout(cleanup, 3000);
}

async function downloadSectionPDF(sectionNum) {
  const jsPDF = window.jspdf?.jsPDF;
  if (!jsPDF) {
    alert('PDF library is not loaded yet. Please connect to the internet once to cache it, or use the Print button.');
    return;
  }
  const btn = document.getElementById(`section-pdf-btn-${sectionNum}`);
  if (btn) { btn.disabled = true; btn.textContent = 'Generating…'; }
  try {
    const data   = collectFormData();
    const doc    = buildDoc(jsPDF, data, state.settings, [sectionNum]);
    const name   = (data.f1.educatorName || 'class-visit').replace(/\s+/g, '-');
    const date   = data.f1.date || new Date().toISOString().slice(0, 10);
    const labels = { 1: 'notice', 2: 'classroom-obs', 3: 'book-moderation' };
    doc.save(`${name}-${date}-${labels[sectionNum]}.pdf`);
  } catch (err) {
    console.error(err);
    alert('PDF generation failed. Please use the Print button instead.');
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = '<span class="btn-icon">⬇</span> Download PDF';
    }
  }
}

/* ── JSON Export ─────────────────────────────────────────── */
function downloadJSON() {
  const data = collectFormData();
  const name = (data.f1.educatorName || 'class-visit').replace(/\s+/g, '-');
  const date = data.f1.date || new Date().toISOString().slice(0, 10);
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = `${name}-${date}.json`;
  a.click();
  URL.revokeObjectURL(url);
}

async function downloadPDFAndJSON() {
  const jsPDF = window.jspdf?.jsPDF;
  if (!jsPDF) {
    alert('PDF library is not loaded yet. Please connect to the internet once to cache it, or use Save as JSON only.');
    return;
  }

  const btn = document.getElementById('download-pdf-json-btn');
  btn.disabled = true;
  btn.textContent = 'Generating…';

  try {
    const data = collectFormData();
    const doc  = buildDoc(jsPDF, data, state.settings);
    const name = (data.f1.educatorName || 'class-visit').replace(/\s+/g, '-');
    const date = data.f1.date || new Date().toISOString().slice(0, 10);
    doc.save(`${name}-${date}.pdf`);
    downloadJSON();
  } catch (err) {
    console.error(err);
    alert('PDF generation failed. Please use the Print button instead.');
  } finally {
    btn.disabled = false;
    btn.innerHTML = '<span class="btn-icon">⬇</span> Save as PDF + JSON';
  }
}

/* ── Learner Camera Capture ──────────────────────────────── */
function setupLearnerCameras() {
  [1, 2, 3].forEach(n => {
    const cameraBtn = document.getElementById(`f3-l${n}-camera-btn`);
    const input     = document.getElementById(`f3-l${n}-photo-input`);
    const thumbWrap = document.getElementById(`f3-l${n}-photo-thumb-wrap`);
    const thumb     = document.getElementById(`f3-l${n}-photo-thumb`);
    const removeBtn = document.getElementById(`f3-l${n}-photo-remove`);

    cameraBtn.addEventListener('click', () => input.click());

    input.addEventListener('change', e => {
      const file = e.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = ev => {
        const url = ev.target.result;
        state.learnerPhotos[`f3-l${n}`] = url;
        thumb.src = url;
        thumbWrap.classList.remove('hidden');
        cameraBtn.classList.add('hidden');
        saveFormData();
      };
      reader.readAsDataURL(file);
      input.value = '';
    });

    removeBtn.addEventListener('click', () => {
      state.learnerPhotos[`f3-l${n}`] = null;
      thumb.src = '';
      thumbWrap.classList.add('hidden');
      cameraBtn.classList.remove('hidden');
      saveFormData();
    });
  });
}

/* ── JSON Import ─────────────────────────────────────────── */
function setupImportJSON() {
  const btn   = document.getElementById('import-json-btn');
  const input = document.getElementById('import-json-input');

  btn.addEventListener('click', () => input.click());

  input.addEventListener('change', e => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const data = JSON.parse(ev.target.result);

        /* Reset all form state before re-populating */
        state.toggleValues   = {};
        state.signedCanvases = new Set();
        state.learnerPhotos  = {};
        document.querySelectorAll('.form-section input, .form-section textarea').forEach(el => {
          if (el.type === 'checkbox') el.checked = false;
          else el.value = '';
        });
        document.querySelectorAll('.yn-btn').forEach(b => b.classList.remove('active-yes', 'active-no'));
        document.querySelectorAll('.toggle-comment-wrap').forEach(w => w.classList.add('hidden'));
        document.querySelectorAll('.toggle-comment').forEach(t => { t.value = ''; });
        document.querySelectorAll('.sig-canvas').forEach(c => { if (c._initialized) clearCanvas(c); });
        [1, 2, 3].forEach(n => {
          const thumb     = document.getElementById(`f3-l${n}-photo-thumb`);
          const thumbWrap = document.getElementById(`f3-l${n}-photo-thumb-wrap`);
          const cameraBtn = document.getElementById(`f3-l${n}-camera-btn`);
          if (thumb)     { thumb.src = ''; }
          if (thumbWrap) { thumbWrap.classList.add('hidden'); }
          if (cameraBtn) { cameraBtn.classList.remove('hidden'); }
        });

        applyFormData(data);
        saveFormData();
        navigateTo(1);
      } catch {
        alert('Invalid file. Please select a JSON file exported from this app.');
      }
      input.value = '';
    };
    reader.readAsText(file);
  });
}

/* ── Print ───────────────────────────────────────────────── */
function printForms() {
  /* Ensure all canvases are rendered before printing */
  document.querySelectorAll('.sig-canvas').forEach(canvas => {
    if (!canvas._initialized) {
      const section = canvas.closest('.form-section');
      section.style.display = 'block';
      initCanvas(canvas);
      section.style.display = '';
    }
  });
  window.print();
}

/* ── Utilities ───────────────────────────────────────────── */
function val(id) {
  const el = document.getElementById(id);
  return el ? el.value : '';
}

function setVal(id, v) {
  const el = document.getElementById(id);
  if (el && v != null) el.value = v;
}

function checked(id) {
  const el = document.getElementById(id);
  return el ? el.checked : false;
}

function setChecked(id, v) {
  const el = document.getElementById(id);
  if (el) el.checked = !!v;
}

function fmtDate(iso) {
  if (!iso) return '';
  try {
    const [y, m, d] = iso.split('-');
    return `${d}/${m}/${y}`;
  } catch { return iso; }
}
