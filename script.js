/* ╔══════════════════════════════════════════════════════════╗
   ║   WINBOX7 LUCKY DRAW  —  script.js                      ║
   ║                                                          ║
   ║   Sections:                                              ║
   ║     1. State & Constants                                 ║
   ║     2. Utility helpers                                   ║
   ║     3. Audio (Web Audio API)                             ║
   ║     4. Admin / Upload screen                             ║
   ║     5. Excel processing (SheetJS)                        ║
   ║     6. Draw animation (scroll + fly + card)              ║
   ║     7. Semifinal screen                                  ║
   ║     8. Spinning Wheel (Canvas)                           ║
   ║     9. Celebration screen                                ║
   ║    10. Export (SheetJS write)                            ║
   ║    11. Misc (fullscreen, sound toggle, resize)           ║
   ║    12. Init / Event bindings                             ║
   ╚══════════════════════════════════════════════════════════╝ */

'use strict';

/* ══════════════════════════════════════════════════════════
   1.  STATE & CONSTANTS
══════════════════════════════════════════════════════════ */
const S = {
  /* Admin rows */
  rowCounter : 0,

  /* Processed data */
  files      : [],   // [{day, source, entries:[], drawCount}]
  allEntries : [],   // [{username, day, source, fileIdx}]

  /* Draw progress */
  semifinalists      : [],
  drawnCount         : 0,
  totalToDraw        : 0,
  curFileIdx         : 0,
  drawnFromCurFile   : 0,
  isDrawing          : false,
  isPaused           : false,

  /* Scroll animation */
  scrollY    : 0,
  scrollRAF  : null,
  scrollSpeed: 1.5,

  /* Draw timeout handle */
  drawTimer  : null,

  /* Wheel */
  wheelParts   : [],  // shuffled semifinalists for wheel
  wheelRot     : 0,   // current rotation (radians)
  wheelSpinning: false,
  rankNow      : 1,   // 1 → 2 → 3

  /* Spin animation helpers */
  spinStart    : null,
  spinDuration : 0,
  spinDelta    : 0,
  spinStartRot : 0,
  spinWinner   : -1,
  spinRAF      : null,

  /* Winners */
  winners  : [],   // [{username,day,source,rank}]

  /* Sound */
  soundOn  : true,
  audioCtx : null,
};

/* Wheel segment colours — cycling palette */
const WHEEL_COLORS = [
  '#7B2FBE','#5B0FA8','#9333EA','#4A1080',
  '#B8900A','#D4A017','#FFD700','#CC8800',
  '#1E0B40','#2D1060','#3B0F6B','#4C1380',
];

/* Rank decorations */
const RANKS = [
  { emoji:'🏆', label:'RANK 1 WINNER', cssClass:'rank1' },
  { emoji:'🥈', label:'RANK 2 WINNER', cssClass:'rank2' },
  { emoji:'🥉', label:'RANK 3 WINNER', cssClass:'rank3' },
];


/* ══════════════════════════════════════════════════════════
   2.  UTILITY HELPERS
══════════════════════════════════════════════════════════ */
const $ = id => document.getElementById(id);

/** Fisher-Yates shuffle (returns new array) */
function shuffle(arr) {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

/** Escape HTML to prevent XSS when inserting user data */
function esc(str) {
  const d = document.createElement('div');
  d.appendChild(document.createTextNode(String(str)));
  return d.innerHTML;
}

/** Switch the visible screen */
function showScreen(id) {
  document.querySelectorAll('.screen').forEach(s => {
    s.classList.remove('active');
    s.classList.add('hidden');
  });
  const el = $(id);
  el.classList.remove('hidden');
  el.classList.add('active');
}

/** Cubic ease-out for wheel spin deceleration */
function easeOut(t) { return 1 - Math.pow(1 - t, 3); }


/* ══════════════════════════════════════════════════════════
   3.  AUDIO  (Web Audio API — no external files needed)
══════════════════════════════════════════════════════════ */
function getCtx() {
  if (!S.audioCtx) {
    S.audioCtx = new (window.AudioContext || window.webkitAudioContext)();
  }
  return S.audioCtx;
}

/**
 * Play a synthetic beep.
 * @param {number} freq     - Hz
 * @param {number} dur      - seconds
 * @param {string} type     - OscillatorType
 * @param {number} vol      - 0–1
 */
function beep(freq = 440, dur = 0.1, type = 'sine', vol = 0.3) {
  if (!S.soundOn) return;
  try {
    const ctx  = getCtx();
    const osc  = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.connect(gain); gain.connect(ctx.destination);
    osc.type = type;
    osc.frequency.setValueAtTime(freq, ctx.currentTime);
    gain.gain.setValueAtTime(vol, ctx.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + dur);
    osc.start(ctx.currentTime);
    osc.stop(ctx.currentTime + dur);
  } catch (_) {}
}

const sfx = {
  tick       : () => beep(900, 0.04, 'square', 0.12),
  select     : () => {
    beep(523, 0.10, 'sine', 0.30);
    setTimeout(() => beep(659, 0.10, 'sine', 0.30), 110);
    setTimeout(() => beep(784, 0.20, 'sine', 0.40), 220);
  },
  winner     : () => {
    [523, 659, 784, 1047].forEach((f, i) =>
      setTimeout(() => beep(f, 0.30, 'sine', 0.45), i * 150)
    );
  },
  spinTick   : () => beep(600, 0.03, 'sawtooth', 0.08),
};


/* ══════════════════════════════════════════════════════════
   4.  ADMIN / UPLOAD SCREEN
══════════════════════════════════════════════════════════ */
function createRow() {
  S.rowCounter++;
  const id  = S.rowCounter;
  const row = document.createElement('div');
  row.className   = 'upload-row';
  row.id          = `row-${id}`;
  row.dataset.rowId = id;

  row.innerHTML = `
    <div class="upload-field">
      <label class="f-label">Day</label>
      <select class="row-day">
        <option value="">Select Day</option>
        <option>Monday</option><option>Tuesday</option><option>Wednesday</option>
        <option>Thursday</option><option>Friday</option><option>Saturday</option>
      </select>
    </div>
    <div class="upload-field">
      <label class="f-label">Source</label>
      <select class="row-src">
        <option value="">Select Source</option>
        <option>YouTube</option><option>Facebook</option><option>Instagram</option>
      </select>
    </div>
    <div class="upload-field">
      <label class="f-label">Excel File (.xlsx / .xls)</label>
      <input type="file" class="row-file" accept=".xlsx,.xls" />
      <span class="f-status">No file selected</span>
    </div>
    <button class="btn-remove" title="Remove row" data-rid="${id}">✕</button>
  `;

  /* Filename display */
  row.querySelector('.row-file').addEventListener('change', e => {
    const st = row.querySelector('.f-status');
    if (e.target.files[0]) {
      st.textContent = `✓ ${e.target.files[0].name}`;
      st.className   = 'f-status ok';
      beep(440, 0.08);
    } else {
      st.textContent = 'No file selected';
      st.className   = 'f-status';
    }
  });

  /* Remove button inside the row */
  row.querySelector('.btn-remove').addEventListener('click', () => removeRow(id));

  return row;
}

function removeRow(id) {
  const rows = document.querySelectorAll('.upload-row');
  if (rows.length <= 1) return; // keep at least one row
  const el = $(`row-${id}`);
  if (!el) return;
  el.style.opacity   = '0';
  el.style.transform = 'translateX(20px)';
  el.style.transition = 'opacity .25s ease, transform .25s ease';
  setTimeout(() => el.remove(), 280);
}

function showError(msg) {
  const el = $('upload-error');
  el.textContent = msg;
  el.classList.remove('hidden');
}


/* ══════════════════════════════════════════════════════════
   5.  EXCEL PROCESSING
══════════════════════════════════════════════════════════ */
/**
 * Read an xlsx file and return an array of non-empty Column B values (skips row 1 header).
 */
function readXlsx(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb   = XLSX.read(new Uint8Array(e.target.result), { type:'array' });
        const ws   = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header:1 });
        const vals = rows
          .slice(1)                          // skip header row
          .map(r => r[1])                    // Column B (index 1)
          .filter(v => v !== undefined && v !== null && String(v).trim() !== '')
          .map(v => String(v).trim());
        resolve(vals);
      } catch(err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

/* ══════════════════════════════════════════════════════════
   5b.  SUBMIT HANDLER — reads all files, builds state
══════════════════════════════════════════════════════════ */
async function handleSubmit() {
  $('upload-error').classList.add('hidden');

  const rows = document.querySelectorAll('.upload-row');
  if (!rows.length) { showError('Please add at least one row.'); return; }

  S.files      = [];
  S.allEntries = [];

  for (const row of rows) {
    const day  = row.querySelector('.row-day').value;
    const src  = row.querySelector('.row-src').value;
    const fi   = row.querySelector('.row-file');

    if (!day || !src) { showError('Please select Day and Source for every row.'); return; }
    if (!fi.files[0]) { showError('Please upload an Excel file for every row.'); return; }

    let entries;
    try {
      entries = await readXlsx(fi.files[0]);
    } catch(e) {
      showError(`Could not read "${fi.files[0].name}". Make sure it is a valid .xlsx file.`);
      return;
    }

    if (!entries.length) {
      showError(`"${fi.files[0].name}" has no usernames in Column B (rows 2+).`);
      return;
    }

    /* 20 % rule — min 1, max 10 */
    const drawCount = Math.max(1, Math.min(10, Math.floor(entries.length * 0.2)));

    S.files.push({ day, source:src, entries:shuffle(entries), drawCount });

    const fIdx = S.files.length - 1;
    entries.forEach(username => S.allEntries.push({ username, day, source:src, fileIdx:fIdx }));
  }

  /* Initialise draw state */
  S.semifinalists    = [];
  S.drawnCount       = 0;
  S.totalToDraw      = S.files.reduce((s, f) => s + f.drawCount, 0);
  S.curFileIdx       = 0;
  S.drawnFromCurFile = 0;

  initDrawScreen();
  showScreen('screen-draw');
}


/* ══════════════════════════════════════════════════════════
   6.  DRAW ANIMATION
══════════════════════════════════════════════════════════ */

/* ── 6a. Initialise draw screen ── */
function initDrawScreen() {
  $('total-count').textContent  = S.allEntries.length;
  $('sel-count').textContent    = '0';
  $('prog-text').textContent    = `0 / ${S.totalToDraw}`;
  $('prog-fill').style.width    = '0%';
  $('sel-container').innerHTML  = '';
  $('draw-status').textContent  = 'Ready to Start';
  buildScrollList();
}

/* ── 6b. Build the scrolling name list (duplicated for seamless loop) ── */
function buildScrollList() {
  const container = $('scroll-names');
  container.innerHTML = '';
  S.scrollY = 0;

  const shuffled = shuffle([...S.allEntries]);
  /* Duplicate the list so the scroll can loop seamlessly */
  [...shuffled, ...shuffled].forEach(entry => {
    const el = document.createElement('div');
    el.className         = 'scroll-name';
    el.textContent       = entry.username;
    el.dataset.username  = entry.username;
    container.appendChild(el);
  });
  startScroll();
}

/* ── 6c. Scroll animation (RAF loop) ── */
function startScroll() {
  stopScroll();
  const container = $('scroll-names');

  function tick() {
    if (!S.isPaused) {
      const items    = container.children;
      if (items.length) {
        const itemH    = (items[0].offsetHeight || 26) + 3; // 3px gap
        const halfH    = (items.length / 2) * itemH;
        S.scrollY     += S.scrollSpeed;
        if (S.scrollY >= halfH) S.scrollY -= halfH;
        container.style.transform = `translateY(-${S.scrollY}px)`;
      }
    }
    S.scrollRAF = requestAnimationFrame(tick);
  }
  S.scrollRAF = requestAnimationFrame(tick);
}

function stopScroll() {
  if (S.scrollRAF) { cancelAnimationFrame(S.scrollRAF); S.scrollRAF = null; }
}

/* ── 6d. Start / pause controls ── */
function startDraw() {
  if (S.isDrawing) return;
  S.isDrawing  = true;
  S.isPaused   = false;
  S.scrollSpeed = 2.5;

  $('btn-start-draw').classList.add('hidden');
  $('btn-pause-draw').classList.remove('hidden');
  scheduleNextDraw();
}

function togglePause() {
  S.isPaused = !S.isPaused;
  const btn  = $('btn-pause-draw');
  if (S.isPaused) {
    btn.querySelector('.di').textContent = '▶';
    btn.querySelector('.dl').textContent = 'Resume';
    $('draw-status').textContent = 'Paused…';
    clearTimeout(S.drawTimer);
  } else {
    btn.querySelector('.di').textContent = '⏸';
    btn.querySelector('.dl').textContent = 'Pause';
    $('draw-status').textContent = 'Drawing…';
    scheduleNextDraw();
  }
}

/* ── 6e. Scheduling logic ── */
function scheduleNextDraw() {
  /* Advance past fully-drawn files */
  while (
    S.curFileIdx < S.files.length &&
    S.drawnFromCurFile >= S.files[S.curFileIdx].drawCount
  ) {
    S.curFileIdx++;
    S.drawnFromCurFile = 0;
  }

  if (S.curFileIdx >= S.files.length) { finishDraw(); return; }

  const cf = S.files[S.curFileIdx];
  $('draw-status').textContent = `Drawing — ${cf.day} / ${cf.source}`;

  /* Dramatic delay: 3–5 seconds (feels good on live stream) */
  const delay = 3000 + Math.random() * 2000;
  S.drawTimer = setTimeout(() => { if (!S.isPaused) performDraw(); }, delay);
}

/* ── 6f. Perform one draw pick ── */
function performDraw() {
  const cf        = S.files[S.curFileIdx];
  const usedNames = new Set(
    S.semifinalists
      .filter(s => s.fileIdx === S.curFileIdx)
      .map(s => s.username)
  );
  const pool = cf.entries.filter(u => !usedNames.has(u));

  if (!pool.length) {
    /* Exhausted this file's pool (shouldn't normally happen) */
    S.drawnFromCurFile = cf.drawCount;
    scheduleNextDraw();
    return;
  }

  const username = pool[Math.floor(Math.random() * pool.length)];
  const pick = { username, day:cf.day, source:cf.source, fileIdx:S.curFileIdx };

  S.semifinalists.push(pick);
  S.drawnCount++;
  S.drawnFromCurFile++;

  updateProgress();
  animateSelection(pick);
  sfx.select();
}

/* ── 6g. Update progress bar ── */
function updateProgress() {
  const pct = (S.drawnCount / S.totalToDraw) * 100;
  $('prog-fill').style.width  = pct + '%';
  $('prog-text').textContent  = `${S.drawnCount} / ${S.totalToDraw}`;
  $('sel-count').textContent  = S.drawnCount;
}

/* ── 6h. Name fly animation  ── */
function animateSelection(pick) {
  /* Highlight matching name(s) in scroll list */
  document.querySelectorAll('.scroll-name').forEach(el => {
    if (el.dataset.username === pick.username) {
      el.classList.add('hl');
      setTimeout(() => el.classList.remove('hl'), 2200);
    }
  });

  const flyEl    = $('fly-el');
  const leftRect = document.querySelector('.panel-left')?.getBoundingClientRect();
  const rightRect= document.querySelector('.panel-right')?.getBoundingClientRect();

  if (!leftRect || !rightRect) { addSelCard(pick); afterCardAdded(); return; }

  /* Position fly element over left panel centre */
  flyEl.textContent  = pick.username;
  flyEl.style.left   = (leftRect.left  + leftRect.width  / 2) + 'px';
  flyEl.style.top    = (leftRect.top   + leftRect.height / 2) + 'px';
  flyEl.style.transition = 'none';
  flyEl.style.opacity    = '1';
  flyEl.style.transform  = 'translate(-50%,-50%) scale(1)';
  flyEl.classList.remove('hidden');

  /* Small tick sound */
  sfx.tick();

  /* Fly to right panel */
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      flyEl.style.transition = 'left 1.1s cubic-bezier(.23,1,.32,1), top 1.1s cubic-bezier(.23,1,.32,1), transform 1.1s cubic-bezier(.23,1,.32,1), opacity .35s ease .8s';
      flyEl.style.left       = (rightRect.left + rightRect.width / 2) + 'px';
      flyEl.style.top        = (rightRect.top  + 90) + 'px';
      flyEl.style.transform  = 'translate(-50%,-50%) scale(1.35)';

      /* Fade out & add card */
      setTimeout(() => {
        flyEl.style.opacity   = '0';
        flyEl.style.transform = 'translate(-50%,-50%) scale(.4)';
        setTimeout(() => {
          flyEl.classList.add('hidden');
          flyEl.style.transition = 'none';
          flyEl.style.opacity    = '1';
          addSelCard(pick);
          afterCardAdded();
        }, 380);
      }, 1050);
    });
  });
}

/** Add winner card to right panel */
function addSelCard(pick) {
  const container = $('sel-container');
  const card      = document.createElement('div');
  card.className  = 'sel-card';
  card.innerHTML  = `
    <div class="card-name">${esc(pick.username)}</div>
    <div class="card-meta">
      <span class="tag tag-day">${esc(pick.day)}</span>
      <span class="tag tag-src">${esc(pick.source)}</span>
    </div>`;
  container.insertBefore(card, container.firstChild);
  container.scrollTop = 0;
}

/** Schedule next draw 2.5 s after card appears */
function afterCardAdded() {
  S.drawTimer = setTimeout(() => {
    if (!S.isPaused) scheduleNextDraw();
  }, 2500);
}

/* ── 6i. Finish draw ── */
function finishDraw() {
  S.isDrawing = false;
  clearTimeout(S.drawTimer);
  stopScroll();

  $('draw-status').textContent = '✓ Selection Complete!';
  $('btn-pause-draw').classList.add('hidden');

  sfx.winner();
  setTimeout(showSemiFinal, 2200);
}


/* ══════════════════════════════════════════════════════════
   7.  SEMIFINAL SCREEN
══════════════════════════════════════════════════════════ */
function showSemiFinal() {
  const grid = $('semi-grid');
  grid.innerHTML = '';

  $('semi-desc').textContent = `${S.semifinalists.length} semifinalist${S.semifinalists.length !== 1 ? 's' : ''} selected`;

  S.semifinalists.forEach((s, i) => {
    const card = document.createElement('div');
    card.className = 'semi-card';
    card.style.animationDelay = `${i * 0.04}s`;
    card.innerHTML = `
      <div class="semi-num">#${i + 1}</div>
      <div class="semi-name">${esc(s.username)}</div>
      <div class="semi-tags">
        <span class="tag tag-day">${esc(s.day)}</span>
        <span class="tag tag-src">${esc(s.source)}</span>
      </div>`;
    grid.appendChild(card);
  });

  showScreen('screen-semi');
  launchConfetti(0.45);
}


/* ══════════════════════════════════════════════════════════
   8.  SPINNING WHEEL
══════════════════════════════════════════════════════════ */

/* ── 8a. Initialise wheel screen ── */
function initWheel() {
  S.wheelParts    = shuffle([...S.semifinalists]);
  S.wheelRot      = 0;
  S.rankNow       = 1;
  S.winners       = [];

  $('winners-list').innerHTML = '';
  updateRankLabel();
  drawWheel();
  showScreen('screen-wheel');
}

function updateRankLabel() {
  const labels = ['🏆 Rank 1','🥈 Rank 2','🥉 Rank 3'];
  $('rank-span').textContent = labels[S.rankNow - 1] ?? `Rank ${S.rankNow}`;
}

/* ── 8b. Draw wheel on canvas ── */
function drawWheel() {
  const canvas = $('wheel-canvas');
  const ctx    = canvas.getContext('2d');
  const W = canvas.width, H = canvas.height;
  const cx = W / 2, cy = H / 2;
  const r  = Math.min(cx, cy) - 18;

  ctx.clearRect(0, 0, W, H);

  const n = S.wheelParts.length;
  if (n === 0) return;

  const seg = (2 * Math.PI) / n;

  for (let i = 0; i < n; i++) {
    const sa  = i * seg + S.wheelRot - Math.PI / 2;
    const ea  = sa + seg;
    const mid = sa + seg / 2;
    const col = WHEEL_COLORS[i % WHEEL_COLORS.length];

    /* Segment fill */
    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.arc(cx, cy, r, sa, ea);
    ctx.closePath();
    ctx.fillStyle = col;
    ctx.fill();

    /* Segment border */
    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.arc(cx, cy, r, sa, ea);
    ctx.closePath();
    ctx.strokeStyle = 'rgba(255,215,0,.25)';
    ctx.lineWidth = 1.5;
    ctx.stroke();

    /* Segment label text */
    ctx.save();
    ctx.translate(cx, cy);
    ctx.rotate(mid);

    const fontSize   = Math.max(8, Math.min(13, 280 / n));
    ctx.font         = `700 ${fontSize}px Orbitron, monospace`;
    ctx.fillStyle    = '#ffffff';
    ctx.textAlign    = 'right';
    ctx.shadowColor  = 'rgba(0,0,0,.7)';
    ctx.shadowBlur   = 4;

    const raw        = S.wheelParts[i].username;
    const maxCh      = Math.max(4, Math.floor((r * 0.72) / (fontSize * 0.62)));
    const label      = raw.length > maxCh ? raw.slice(0, maxCh) + '…' : raw;
    ctx.fillText(label, r - 12, fontSize * 0.38);
    ctx.restore();
  }

  /* Outer ring */
  ctx.beginPath();
  ctx.arc(cx, cy, r, 0, 2 * Math.PI);
  ctx.strokeStyle = 'rgba(255,215,0,.55)';
  ctx.lineWidth   = 3;
  ctx.stroke();

  /* Centre hub hole (the emoji hub is an HTML overlay) */
  ctx.beginPath();
  ctx.arc(cx, cy, 28, 0, 2 * Math.PI);
  ctx.fillStyle   = '#09090E';
  ctx.fill();
  ctx.strokeStyle = '#FFD700';
  ctx.lineWidth   = 3;
  ctx.stroke();
}

/* ── 8c. Spin ── */
function spinWheel() {
  if (S.wheelSpinning || S.rankNow > 3 || S.wheelParts.length === 0) return;

  S.wheelSpinning = true;
  $('btn-spin').disabled = true;

  const n   = S.wheelParts.length;
  const seg = (2 * Math.PI) / n;

  /* Pre-select winner index (random) */
  const wi = Math.floor(Math.random() * n);
  S.spinWinner = wi;

  /*
   * Calculate target rotation so that segment[wi] is at the top (pointer).
   * Segment[i] midpoint angle (in wheel space) = i*seg + seg/2.
   * After rotation R, its canvas angle = i*seg + seg/2 + R − π/2.
   * We want that to be −π/2 (pointing up): R = −(i*seg + seg/2).
   */
  const targetRot = -(wi * seg + seg / 2);

  /* How far we need to travel (forward) to reach targetRot, plus extra turns */
  let delta = targetRot - S.wheelRot;
  /* Normalise to 0..2π */
  delta = ((delta % (2 * Math.PI)) + 2 * Math.PI) % (2 * Math.PI);
  /* Add 6–9 full revolutions for drama */
  delta += (6 + Math.floor(Math.random() * 4)) * 2 * Math.PI;

  S.spinStartRot  = S.wheelRot;
  S.spinDelta     = delta;
  S.spinDuration  = 7000 + Math.random() * 2000; // 7–9 s
  S.spinStart     = null;

  requestAnimationFrame(ts => animateSpin(ts));
}

/* ── 8d. Spin animation loop ── */
function animateSpin(ts) {
  if (!S.spinStart) S.spinStart = ts;

  const elapsed  = ts - S.spinStart;
  const progress = Math.min(elapsed / S.spinDuration, 1);
  const eased    = easeOut(progress);

  S.wheelRot = S.spinStartRot + S.spinDelta * eased;

  /* Play tick sounds — faster when progress is low, none near the end */
  if (progress < 0.85) {
    const speed = (1 - progress) * S.spinDelta / S.spinDuration * 1000; // rad/s
    if (speed > 2 && Math.random() < 0.12) sfx.spinTick();
  }

  drawWheel();

  if (progress < 1) {
    S.spinRAF = requestAnimationFrame(ts2 => animateSpin(ts2));
  } else {
    S.wheelSpinning = false;
    setTimeout(() => revealWinner(S.spinWinner), 400);
  }
}

/* ── 8e. Reveal winner overlay ── */
function revealWinner(wi) {
  sfx.winner();
  launchConfetti(0.4);

  const part  = S.wheelParts[wi];
  const rankD = RANKS[S.rankNow - 1];

  /* Record winner */
  S.winners.push({ ...part, rank:S.rankNow });

  /* Add to sidebar */
  const entry = document.createElement('div');
  entry.className = 'winner-entry';
  entry.innerHTML = `
    <div class="w-rank">${rankD.emoji}</div>
    <div class="w-name">${esc(part.username)}</div>
    <div class="w-meta">${esc(part.day)} &bull; ${esc(part.source)}</div>`;
  $('winners-list').appendChild(entry);

  /* Show overlay */
  const ov = document.createElement('div');
  ov.className = 'result-overlay';
  const isLast = S.rankNow >= 3;
  ov.innerHTML = `
    <div class="ov-rank">${rankD.emoji}</div>
    <div class="ov-label">${rankD.label}</div>
    <div class="ov-name">${esc(part.username)}</div>
    <div class="ov-meta">${esc(part.day)} &bull; ${esc(part.source)}</div>
    <button class="ov-btn" id="ov-continue">
      ${isLast ? '🎉 Show All Winners' : 'Continue ▶'}
    </button>`;
  document.body.appendChild(ov);

  $('ov-continue').addEventListener('click', () => {
    ov.remove();
    /* Remove this winner from the wheel */
    S.wheelParts = S.wheelParts.filter(p => p.username !== part.username);
    S.rankNow++;

    if (S.rankNow > 3) {
      setTimeout(showCelebration, 400);
    } else {
      updateRankLabel();
      drawWheel();
      $('btn-spin').disabled = false;
    }
  });
}


/* ══════════════════════════════════════════════════════════
   9.  CELEBRATION SCREEN
══════════════════════════════════════════════════════════ */
function showCelebration() {
  const display = $('final-display');
  display.innerHTML = '';

  S.winners.forEach((w, i) => {
    const rd   = RANKS[i] || RANKS[2];
    const card = document.createElement('div');
    card.className = 'final-card';
    card.style.animationDelay = `${i * 0.28}s`;
    card.innerHTML = `
      <span class="f-icon">${rd.emoji}</span>
      <div  class="f-rlbl">${rd.label}</div>
      <div  class="f-name">${esc(w.username)}</div>
      <div  class="f-day"> ${esc(w.day)}</div>
      <div  class="f-src"> ${esc(w.source)}</div>`;
    display.appendChild(card);
  });

  showScreen('screen-celebrate');

  /* Multi-burst confetti */
  [0, 800, 1800, 3000].forEach(d => setTimeout(() => launchConfetti(0.9), d));
  sfx.winner();
}


/* ══════════════════════════════════════════════════════════
   10.  EXPORT (SheetJS write)
══════════════════════════════════════════════════════════ */
function exportSheet(rows, sheetName, filename) {
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, filename);
}

function exportSemifinalists() {
  const rows = [
    ['#', 'Username', 'Day', 'Source'],
    ...S.semifinalists.map((s, i) => [i + 1, s.username, s.day, s.source]),
  ];
  exportSheet(rows, 'Semifinalists', 'WinBox7_Semifinalists.xlsx');
}

function exportWinners() {
  const rows = [
    ['Rank', 'Username', 'Day', 'Source'],
    ...S.winners.map(w => [`Rank ${w.rank}`, w.username, w.day, w.source]),
  ];
  exportSheet(rows, 'Winners', 'WinBox7_Winners.xlsx');
}


/* ══════════════════════════════════════════════════════════
   11.  MISC — fullscreen, sound, canvas resize
══════════════════════════════════════════════════════════ */
function toggleFullscreen() {
  if (!document.fullscreenElement) {
    document.documentElement.requestFullscreen().catch(() => {});
  } else {
    document.exitFullscreen().catch(() => {});
  }
}

function handleResize() {
  /* Resize wheel canvas to fit available space */
  const wheelScreen = $('screen-wheel');
  if (wheelScreen && !wheelScreen.classList.contains('hidden')) {
    const canvas = $('wheel-canvas');
    const max = Math.min(
      window.innerWidth  - 360,
      window.innerHeight - 180,
      500
    );
    if (max > 50) { canvas.width = max; canvas.height = max; drawWheel(); }
  }
}

function resetEverything() {
  /* Clear all state and return to admin screen */
  clearTimeout(S.drawTimer);
  stopScroll();
  Object.assign(S, {
    rowCounter:0, files:[], allEntries:[],
    semifinalists:[], drawnCount:0, totalToDraw:0,
    curFileIdx:0, drawnFromCurFile:0,
    isDrawing:false, isPaused:false,
    scrollY:0, scrollSpeed:1.5,
    wheelParts:[], wheelRot:0, rankNow:1, winners:[],
  });

  /* Re-add first row */
  const rc = $('rows-container');
  rc.innerHTML = '';
  rc.appendChild(createRow());

  showScreen('screen-admin');
}


/* ══════════════════════════════════════════════════════════
   CONFETTI helper
══════════════════════════════════════════════════════════ */
function launchConfetti(intensity = 0.5) {
  if (typeof confetti === 'undefined') return;
  confetti({
    particleCount : Math.floor(180 * intensity),
    spread        : 80,
    startVelocity : 38,
    origin        : { x: 0.2 + Math.random() * 0.6, y: 0.15 },
    colors        : ['#FFD700','#A855F7','#FF6B35','#FFFFFF','#FF1493','#00CFFF'],
  });
}


/* ══════════════════════════════════════════════════════════
   12.  INIT — wire up all event listeners
══════════════════════════════════════════════════════════ */
document.addEventListener('DOMContentLoaded', () => {

  /* ── Admin screen ── */
  const rc = $('rows-container');
  rc.appendChild(createRow());   // start with one row

  $('btn-add-row').addEventListener('click', () => rc.appendChild(createRow()));
  $('btn-submit').addEventListener('click', handleSubmit);

  /* ── Draw screen ── */
  $('btn-start-draw').addEventListener('click', startDraw);
  $('btn-pause-draw').addEventListener('click', togglePause);

  /* ── Semifinal screen ── */
  $('btn-start-final').addEventListener('click', initWheel);
  $('btn-export-semi').addEventListener('click', exportSemifinalists);

  /* ── Wheel screen ── */
  $('btn-spin').addEventListener('click', spinWheel);

  /* ── Celebration screen ── */
  $('btn-export-winners').addEventListener('click', exportWinners);
  $('btn-restart').addEventListener('click', resetEverything);

  /* ── Header controls ── */
  $('btn-fullscreen').addEventListener('click', toggleFullscreen);

  $('btn-sound').addEventListener('click', () => {
    S.soundOn = !S.soundOn;
    $('btn-sound').textContent = S.soundOn ? '🔊' : '🔇';
  });

  /* ── Responsive wheel resize ── */
  window.addEventListener('resize', handleResize);

});
