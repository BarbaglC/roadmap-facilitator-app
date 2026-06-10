# Roadmap Workshop App — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a single self-contained HTML file that serves as a live facilitation cockpit for roadmap strategy sessions — import cards, vote, capture decisions, prioritise into releases, export results.

**Architecture:** One HTML file with embedded CSS and JavaScript. No frameworks, no build tools, no server. State held in plain JavaScript objects. Exports use browser-native Blob downloads and SheetJS for Excel.

**Tech Stack:** HTML5, CSS3 (custom properties for theming), vanilla JavaScript (ES6+), SheetJS via CDN for Excel export.

---

## File Structure

| File | Responsibility |
|------|---------------|
| `index.html` | The entire app — HTML structure, embedded `<style>`, embedded `<script>` |

This is a single-file app. All code lives in `index.html`. The plan breaks it into logical tasks that build on each other — each task adds a working layer.

---

### Task 1: HTML Shell + CSS Design System

**Files:**
- Create: `index.html`

Build the empty page skeleton with all CSS custom properties, typography, and the three-tab navigation (Before / Vote / Prioritise). No functionality yet — just the visual foundation.

- [ ] **Step 1: Create the HTML file with head, CSS variables, and body shell**

```html
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Roadmap Workshop</title>
<style>
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

:root {
  --bg-page: #f8f9fb;
  --bg-card: #ffffff;
  --bg-sidebar: #f5f0ff;
  --border: rgba(0,0,0,0.08);
  --border-hover: rgba(0,0,0,0.14);
  --text-primary: #111827;
  --text-secondary: #4b5563;
  --text-muted: #9ca3af;
  --accent: #3b82f6;
  --accent-dim: rgba(59,130,246,0.08);
  --accent-border: rgba(59,130,246,0.25);
  --green: #16a34a;
  --green-dim: rgba(22,163,74,0.08);
  --purple: #7c3aed;
  --purple-dim: rgba(124,58,237,0.08);
  --grey-action: #6b7280;
  --radius-sm: 6px;
  --radius-md: 8px;
  --radius-lg: 12px;
  --font: 'Inter', system-ui, -apple-system, sans-serif;
  --font-mono: 'JetBrains Mono', 'Fira Code', monospace;
  /* Category accent colours */
  --cat-blue: #dbeafe;
  --cat-lavender: #ede9fe;
  --cat-mint: #d1fae5;
  --cat-yellow: #fef9c3;
  --cat-grey: #f1f5f9;
  --cat-teal: #ccfbf1;
  --cat-indigo: #e0e7ff;
  --cat-lime: #ecfccb;
}

body {
  font-family: var(--font);
  background: var(--bg-page);
  color: var(--text-primary);
  font-size: 13px;
  line-height: 1.5;
  min-height: 100vh;
  display: flex;
  flex-direction: column;
}

/* Top bar */
.topbar {
  height: 52px;
  background: var(--bg-card);
  border-bottom: 1px solid var(--border);
  display: flex;
  align-items: center;
  padding: 0 20px;
  gap: 14px;
  flex-shrink: 0;
  position: sticky;
  top: 0;
  z-index: 100;
}

.topbar-title {
  font-size: 14px;
  font-weight: 600;
  color: var(--text-primary);
  letter-spacing: -0.01em;
}

.phase-tabs {
  display: flex;
  gap: 2px;
  margin-left: auto;
}

.ptab {
  font-size: 12px;
  font-weight: 500;
  padding: 6px 14px;
  border-radius: var(--radius-sm);
  border: 1px solid transparent;
  cursor: pointer;
  transition: all 0.15s;
  color: var(--text-secondary);
  background: transparent;
  font-family: var(--font);
}

.ptab:hover {
  color: var(--text-primary);
  background: var(--bg-page);
}

.ptab.active {
  background: var(--accent-dim);
  border-color: var(--accent-border);
  color: var(--accent);
}

/* Views */
.view {
  display: none;
  flex: 1;
  overflow: hidden;
}

.view.active {
  display: flex;
}

/* Buttons */
.btn {
  font-family: var(--font);
  font-size: 12px;
  font-weight: 500;
  padding: 6px 14px;
  border-radius: var(--radius-sm);
  border: 1px solid var(--border-hover);
  background: var(--bg-card);
  color: var(--text-secondary);
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  gap: 6px;
  transition: all 0.12s;
  white-space: nowrap;
}

.btn:hover {
  background: var(--bg-page);
  color: var(--text-primary);
}

.btn-primary {
  background: var(--accent);
  color: #fff;
  border-color: var(--accent);
}

.btn-primary:hover {
  background: #2563eb;
  border-color: #2563eb;
  color: #fff;
}
</style>
</head>
<body>

<header class="topbar">
  <span class="topbar-title">Roadmap Workshop</span>
  <nav class="phase-tabs">
    <button class="ptab active" data-view="before">Before</button>
    <button class="ptab" data-view="vote">Vote topics</button>
    <button class="ptab" data-view="prioritise">Prioritise</button>
  </nav>
</header>

<div class="view active" id="view-before"></div>
<div class="view" id="view-vote"></div>
<div class="view" id="view-prioritise"></div>

<script>
// View switching
document.querySelectorAll('.ptab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.ptab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
    tab.classList.add('active');
    document.getElementById('view-' + tab.dataset.view).classList.add('active');
  });
});
</script>
</body>
</html>
```

- [ ] **Step 2: Open the file in a browser and verify**

Run: Open `index.html` in a browser (double-click the file or `open index.html` on Mac).

Expected: A white top bar with "Roadmap Workshop" title and three tabs. Clicking tabs switches the active state (blue highlight). Background is light grey. Three empty view panels.

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "feat: HTML shell with CSS design system and tab navigation"
```

---

### Task 2: App State + Category Colour Assignment

**Files:**
- Modify: `index.html` (add to `<script>` section)

Define the core data model and a utility to assign pastel colours to categories consistently.

- [ ] **Step 1: Add state variables and helper functions to the script section**

Add this at the top of the `<script>` block (replace the existing content, keeping the view switching code):

```javascript
// ─── STATE ───────────────────────────────────────────────────────
const state = {
  sessionName: 'Roadmap Workshop',
  sessionDate: new Date().toISOString().slice(0, 10),
  topics: [],       // {id, title, desc, category, votes, questions:[], decisions:[], parked:boolean}
  releases: [],     // {id, name, cardIds:[]}
  hasUnsavedChanges: false,
  lastExportTime: null,
};

let idCounter = 0;
function newId() { return 't' + (++idCounter) + '_' + Date.now(); }

// ─── CATEGORY COLOURS ────────────────────────────────────────────
const CATEGORY_COLOURS = [
  { bg: '#dbeafe', border: '#93c5fd' },  // soft blue
  { bg: '#ede9fe', border: '#c4b5fd' },  // lavender
  { bg: '#d1fae5', border: '#6ee7b7' },  // mint
  { bg: '#fef9c3', border: '#fde047' },  // soft yellow
  { bg: '#f1f5f9', border: '#cbd5e1' },  // light grey
  { bg: '#ccfbf1', border: '#5eead4' },  // teal
  { bg: '#e0e7ff', border: '#a5b4fc' },  // indigo
  { bg: '#ecfccb', border: '#bef264' },  // lime
];

const categoryColourMap = {};
let nextColourIndex = 0;

function getCategoryColour(category) {
  if (!category) category = 'Other';
  if (!categoryColourMap[category]) {
    categoryColourMap[category] = CATEGORY_COLOURS[nextColourIndex % CATEGORY_COLOURS.length];
    nextColourIndex++;
  }
  return categoryColourMap[category];
}

// ─── UNSAVED CHANGES WARNING ─────────────────────────────────────
function markDirty() {
  state.hasUnsavedChanges = true;
}

window.addEventListener('beforeunload', (e) => {
  if (state.hasUnsavedChanges) {
    e.preventDefault();
    e.returnValue = '';
  }
});

// ─── VIEW SWITCHING ──────────────────────────────────────────────
function showView(viewName) {
  document.querySelectorAll('.ptab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.querySelector(`[data-view="${viewName}"]`).classList.add('active');
  document.getElementById('view-' + viewName).classList.add('active');
  if (viewName === 'vote') renderVoteView();
  if (viewName === 'prioritise') renderPrioView();
}

document.querySelectorAll('.ptab').forEach(tab => {
  tab.addEventListener('click', () => showView(tab.dataset.view));
});
```

- [ ] **Step 2: Open in browser, verify no console errors**

Open browser dev tools (F12 → Console tab). Reload `index.html`.

Expected: No errors. Tab switching still works.

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "feat: add app state model, category colours, and unsaved-changes warning"
```

---

### Task 3: Phase 1 — Before (Import & Setup)

**Files:**
- Modify: `index.html` (add CSS + HTML for Before view + import logic in script)

Build the import UI: session name/date fields, file upload zone, card preview with category editing/deletion, and the "Start voting" button.

- [ ] **Step 1: Add Before view CSS**

Add these styles before the closing `</style>` tag:

```css
/* ═══ BEFORE VIEW ═══ */
.before-view {
  flex: 1;
  display: flex;
  flex-direction: column;
  align-items: center;
  padding: 40px 20px;
  gap: 24px;
  overflow-y: auto;
}

.session-fields {
  display: flex;
  gap: 16px;
  align-items: center;
  width: 100%;
  max-width: 560px;
}

.session-field {
  flex: 1;
  display: flex;
  flex-direction: column;
  gap: 4px;
}

.session-field label {
  font-size: 11px;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 0.05em;
  color: var(--text-muted);
}

.session-field input {
  font-family: var(--font);
  font-size: 14px;
  font-weight: 500;
  padding: 8px 12px;
  border: 1px solid var(--border);
  border-radius: var(--radius-md);
  background: var(--bg-card);
  color: var(--text-primary);
  outline: none;
  transition: border-color 0.15s;
}

.session-field input:focus {
  border-color: var(--accent);
}

.upload-zone {
  width: 100%;
  max-width: 560px;
  border: 1.5px dashed var(--accent-border);
  border-radius: var(--radius-lg);
  padding: 48px 40px;
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 10px;
  cursor: pointer;
  transition: all 0.2s;
  background: var(--accent-dim);
}

.upload-zone:hover,
.upload-zone.drag-over {
  border-color: var(--accent);
  background: rgba(59,130,246,0.12);
}

.upload-zone.has-file {
  border-color: var(--green);
  border-style: solid;
  background: var(--green-dim);
  cursor: default;
}

.upload-title {
  font-size: 14px;
  font-weight: 600;
  color: var(--text-primary);
}

.upload-sub {
  font-size: 12px;
  color: var(--text-muted);
  text-align: center;
}

/* Card preview */
.preview-panel {
  width: 100%;
  max-width: 560px;
  background: var(--bg-card);
  border: 1px solid var(--border);
  border-radius: var(--radius-lg);
  overflow: hidden;
  display: none;
}

.preview-panel.visible {
  display: block;
}

.preview-header {
  padding: 12px 18px;
  border-bottom: 1px solid var(--border);
  font-size: 12px;
  font-weight: 600;
  color: var(--text-secondary);
  display: flex;
  align-items: center;
  justify-content: space-between;
}

.preview-categories {
  padding: 8px 18px;
  font-size: 11px;
  color: var(--text-muted);
  border-bottom: 1px solid var(--border);
}

.preview-body {
  max-height: 320px;
  overflow-y: auto;
  padding: 12px 18px;
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.preview-card {
  display: flex;
  align-items: flex-start;
  gap: 10px;
  padding: 10px 12px;
  background: var(--bg-page);
  border: 1px solid var(--border);
  border-radius: var(--radius-md);
  border-left: 3px solid var(--cat-grey);
}

.preview-card-content {
  flex: 1;
  min-width: 0;
}

.preview-card-title {
  font-size: 13px;
  font-weight: 500;
  color: var(--text-primary);
}

.preview-card-desc {
  font-size: 11px;
  color: var(--text-muted);
  margin-top: 2px;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

.preview-card-cat {
  font-size: 10px;
  font-weight: 600;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: var(--text-muted);
  cursor: pointer;
  padding: 2px 6px;
  border-radius: 4px;
  border: 1px solid var(--border);
  background: var(--bg-card);
  font-family: var(--font);
}

.preview-card-cat:hover {
  border-color: var(--accent-border);
  color: var(--accent);
}

.preview-card-del {
  font-size: 16px;
  color: var(--text-muted);
  cursor: pointer;
  padding: 2px 4px;
  border-radius: 4px;
  border: none;
  background: none;
  font-family: var(--font);
  line-height: 1;
}

.preview-card-del:hover {
  color: var(--grey-action);
  background: var(--bg-page);
}

.preview-footer {
  padding: 14px 18px;
  border-top: 1px solid var(--border);
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.manual-add-row {
  width: 100%;
  max-width: 560px;
  display: flex;
  gap: 8px;
}

.manual-add-row input {
  flex: 1;
  font-family: var(--font);
  font-size: 12px;
  padding: 8px 12px;
  border: 1px solid var(--border);
  border-radius: var(--radius-md);
  background: var(--bg-card);
  color: var(--text-primary);
  outline: none;
}

.manual-add-row input:focus {
  border-color: var(--accent);
}
```

- [ ] **Step 2: Add Before view HTML**

Replace the empty `<div class="view active" id="view-before"></div>` with:

```html
<div class="view active" id="view-before">
  <div class="before-view">
    <div style="text-align:center">
      <div style="font-size:20px;font-weight:600;color:var(--text-primary)">Set up your workshop</div>
      <div style="font-size:12px;color:var(--text-muted);margin-top:6px">Import a prepared JSON, load a previous session, or add cards manually.</div>
    </div>

    <div class="session-fields">
      <div class="session-field">
        <label>Session name</label>
        <input type="text" id="session-name" value="Roadmap Workshop" onchange="state.sessionName=this.value">
      </div>
      <div class="session-field" style="max-width:160px">
        <label>Date</label>
        <input type="date" id="session-date" onchange="state.sessionDate=this.value">
      </div>
    </div>

    <input type="file" id="file-input" accept=".json,.html,.htm" style="display:none">
    <div class="upload-zone" id="upload-zone">
      <div class="upload-title">Drop file here or click to upload</div>
      <div class="upload-sub">Prepared JSON file or previous session HTML</div>
    </div>

    <div class="manual-add-row">
      <input type="text" id="manual-title" placeholder="Add a card manually — type title and press Enter" onkeydown="if(event.key==='Enter')addManualCard()">
      <button class="btn" onclick="addManualCard()">Add card</button>
    </div>

    <div class="preview-panel" id="preview-panel">
      <div class="preview-header">
        <span id="preview-count">0 topics loaded</span>
        <button class="btn" onclick="resetImport()">Clear all</button>
      </div>
      <div class="preview-categories" id="preview-categories"></div>
      <div class="preview-body" id="preview-body"></div>
      <div class="preview-footer">
        <span style="font-size:11px;color:var(--text-muted)">Click a category to rename it</span>
        <button class="btn btn-primary" onclick="showView('vote')">Start voting →</button>
      </div>
    </div>
  </div>
</div>
```

- [ ] **Step 3: Add Before view JavaScript**

Add these functions to the `<script>` block:

```javascript
// ─── BEFORE VIEW ─────────────────────────────────────────────────
function initBeforeView() {
  document.getElementById('session-date').value = state.sessionDate;
  document.getElementById('session-name').value = state.sessionName;
}

// File upload handling
const uploadZone = document.getElementById('upload-zone');
const fileInput = document.getElementById('file-input');

uploadZone.addEventListener('click', () => {
  if (!uploadZone.classList.contains('has-file')) fileInput.click();
});

uploadZone.addEventListener('dragover', (e) => {
  e.preventDefault();
  uploadZone.classList.add('drag-over');
});

uploadZone.addEventListener('dragleave', () => {
  uploadZone.classList.remove('drag-over');
});

uploadZone.addEventListener('drop', (e) => {
  e.preventDefault();
  uploadZone.classList.remove('drag-over');
  if (e.dataTransfer.files.length) handleFileImport(e.dataTransfer.files[0]);
});

fileInput.addEventListener('change', (e) => {
  if (e.target.files.length) handleFileImport(e.target.files[0]);
});

async function handleFileImport(file) {
  const text = await file.text();

  if (file.name.endsWith('.json')) {
    try {
      const data = JSON.parse(text);
      if (data.sessionName) {
        state.sessionName = data.sessionName;
        document.getElementById('session-name').value = data.sessionName;
      }
      if (data.date) {
        state.sessionDate = data.date;
        document.getElementById('session-date').value = data.date;
      }
      if (Array.isArray(data.topics)) {
        state.topics = data.topics.map(t => ({
          id: t.id || newId(),
          title: t.title || '',
          desc: t.desc || '',
          category: t.category || 'Other',
          votes: t.votes || 0,
          questions: t.questions || [],
          decisions: t.decisions || [],
          parked: false,
        }));
      }
    } catch (err) {
      alert('Could not parse JSON file. Please check the format.');
      return;
    }
  } else {
    // Try parsing as session HTML export
    const parser = new DOMParser();
    const doc = parser.parseFromString(text, 'text/html');
    const isArchive = doc.querySelector('meta[name="sr-archive"]');
    if (isArchive) {
      const rows = doc.querySelectorAll('table:last-of-type tbody tr');
      state.topics = Array.from(rows).map(row => {
        const cells = row.querySelectorAll('td');
        return {
          id: newId(),
          title: (cells[0] && cells[0].textContent.trim()) || '',
          desc: (cells[1] && cells[1].textContent.trim()) || '',
          category: 'Imported',
          votes: parseInt(cells[2] && cells[2].textContent.trim()) || 0,
          questions: [],
          decisions: [],
          parked: false,
        };
      }).filter(t => t.title);
      const meetingMeta = doc.querySelector('meta[name="sr-meeting"]');
      if (meetingMeta) {
        state.sessionName = meetingMeta.content;
        document.getElementById('session-name').value = meetingMeta.content;
      }
    } else {
      alert('HTML file does not appear to be a session export. Look for files with "sr-archive" metadata.');
      return;
    }
  }

  uploadZone.classList.add('has-file');
  uploadZone.querySelector('.upload-title').textContent = file.name;
  uploadZone.querySelector('.upload-sub').textContent = state.topics.length + ' topics loaded';
  renderPreview();
}

function addManualCard() {
  const input = document.getElementById('manual-title');
  const title = input.value.trim();
  if (!title) return;
  state.topics.push({
    id: newId(),
    title: title,
    desc: '',
    category: 'Other',
    votes: 0,
    questions: [],
    decisions: [],
    parked: false,
  });
  input.value = '';
  renderPreview();
}

function resetImport() {
  state.topics = [];
  uploadZone.classList.remove('has-file');
  uploadZone.querySelector('.upload-title').textContent = 'Drop file here or click to upload';
  uploadZone.querySelector('.upload-sub').textContent = 'Prepared JSON file or previous session HTML';
  fileInput.value = '';
  document.getElementById('preview-panel').classList.remove('visible');
}

function renderPreview() {
  const panel = document.getElementById('preview-panel');
  if (!state.topics.length) {
    panel.classList.remove('visible');
    return;
  }
  panel.classList.add('visible');

  // Count by category
  const catCounts = {};
  state.topics.forEach(t => {
    const cat = t.category || 'Other';
    catCounts[cat] = (catCounts[cat] || 0) + 1;
  });
  document.getElementById('preview-count').textContent = state.topics.length + ' topics loaded';
  document.getElementById('preview-categories').textContent =
    Object.entries(catCounts).map(([cat, n]) => `${cat} (${n})`).join(' · ');

  // Render cards
  const body = document.getElementById('preview-body');
  body.innerHTML = state.topics.map(t => {
    const colour = getCategoryColour(t.category);
    return `<div class="preview-card" style="border-left-color:${colour.border}">
      <div class="preview-card-content">
        <div class="preview-card-title">${escapeHtml(t.title)}</div>
        ${t.desc ? `<div class="preview-card-desc">${escapeHtml(t.desc)}</div>` : ''}
      </div>
      <button class="preview-card-cat" onclick="renameCategory('${escapeAttr(t.category)}')">${escapeHtml(t.category)}</button>
      <button class="preview-card-del" onclick="deletePreviewCard('${t.id}')" title="Remove">×</button>
    </div>`;
  }).join('');
}

function deletePreviewCard(id) {
  state.topics = state.topics.filter(t => t.id !== id);
  renderPreview();
}

function renameCategory(oldName) {
  const newName = prompt('Rename category:', oldName);
  if (newName && newName.trim() && newName.trim() !== oldName) {
    state.topics.forEach(t => {
      if (t.category === oldName) t.category = newName.trim();
    });
    // Update colour map
    if (categoryColourMap[oldName]) {
      categoryColourMap[newName.trim()] = categoryColourMap[oldName];
      delete categoryColourMap[oldName];
    }
    renderPreview();
  }
}

// HTML escaping
function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

function escapeAttr(str) {
  return str.replace(/'/g, "\\'").replace(/"/g, '&quot;');
}

// Init
initBeforeView();
```

- [ ] **Step 4: Open in browser and test the Before phase**

Test:
1. Page loads with session name "Roadmap Workshop" and today's date
2. Create a test JSON file and drag it onto the upload zone — cards appear in preview
3. Click a category label — rename prompt works
4. Click × on a card — it's removed
5. Type in the manual input and press Enter — card appears in preview

- [ ] **Step 5: Commit**

```bash
git add index.html
git commit -m "feat: Phase 1 Before view — file import, manual entry, card preview"
```

---

### Task 4: Phase 2 — Vote Topics View

**Files:**
- Modify: `index.html` (add CSS + HTML + JS for voting)

Build the voting interface: category-grouped cards with vote buttons, auto-sort, inline edit/delete, expand for questions/decisions, parking lot sidebar.

- [ ] **Step 1: Add Vote view CSS**

Add before the closing `</style>`:

```css
/* ═══ VOTE VIEW ═══ */
.vote-layout {
  flex: 1;
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

.vote-toolbar {
  background: var(--bg-card);
  border-bottom: 1px solid var(--border);
  padding: 10px 20px;
  display: flex;
  gap: 10px;
  align-items: center;
  flex-shrink: 0;
}

.vote-toolbar-info {
  font-size: 12px;
  color: var(--text-muted);
}

.vote-toolbar-info strong {
  color: var(--green);
  font-family: var(--font-mono);
}

.vote-body {
  flex: 1;
  display: flex;
  overflow: hidden;
}

.vote-main {
  flex: 1;
  overflow-y: auto;
  padding: 20px 24px;
}

.vote-main::-webkit-scrollbar { width: 4px; }
.vote-main::-webkit-scrollbar-thumb { background: var(--border-hover); border-radius: 2px; }

/* Category groups */
.cat-group { margin-bottom: 24px; }

.cat-header {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 8px 14px;
  border-radius: var(--radius-md);
  margin-bottom: 12px;
  font-size: 12px;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 0.05em;
  color: var(--text-secondary);
}

.cat-header-count {
  font-size: 11px;
  font-weight: 400;
  color: var(--text-muted);
  margin-left: 4px;
  font-family: var(--font-mono);
}

/* Vote cards */
.vote-cards {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
  gap: 12px;
}

.vote-card {
  background: var(--bg-card);
  border: 1px solid var(--border);
  border-radius: var(--radius-md);
  border-left: 3px solid var(--cat-grey);
  padding: 14px;
  display: flex;
  flex-direction: column;
  gap: 8px;
  transition: box-shadow 0.15s;
}

.vote-card:hover {
  box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}

.vote-card-top {
  display: flex;
  align-items: flex-start;
  gap: 12px;
}

.vote-count {
  font-size: 20px;
  font-weight: 700;
  font-family: var(--font-mono);
  color: var(--text-muted);
  min-width: 32px;
  text-align: center;
  line-height: 1.2;
}

.vote-count.has-votes { color: var(--green); }

.vote-card-body {
  flex: 1;
  min-width: 0;
}

.vote-card-title {
  font-size: 13px;
  font-weight: 600;
  color: var(--text-primary);
  line-height: 1.4;
}

.vote-card-desc {
  font-size: 12px;
  color: var(--text-secondary);
  margin-top: 3px;
  display: -webkit-box;
  -webkit-line-clamp: 1;
  -webkit-box-orient: vertical;
  overflow: hidden;
}

.vote-card-actions {
  display: flex;
  align-items: center;
  gap: 6px;
  margin-top: 4px;
}

.vote-btn {
  width: 28px;
  height: 28px;
  border-radius: 50%;
  border: 1px solid var(--border-hover);
  background: transparent;
  cursor: pointer;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 16px;
  font-weight: 500;
  color: var(--text-secondary);
  transition: all 0.12s;
  font-family: var(--font);
}

.vote-btn.up:hover {
  background: var(--accent-dim);
  color: var(--accent);
  border-color: var(--accent-border);
}

.vote-btn.down:hover {
  background: var(--bg-page);
  color: var(--grey-action);
}

.vote-btn:disabled { opacity: 0.3; cursor: not-allowed; }

.card-icon-btn {
  width: 24px;
  height: 24px;
  border-radius: var(--radius-sm);
  border: none;
  background: transparent;
  cursor: pointer;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 12px;
  color: var(--text-muted);
  transition: all 0.12s;
}

.card-icon-btn:hover { background: var(--bg-page); color: var(--text-primary); }
.card-icon-btn.park:hover { background: var(--cat-lavender); color: var(--purple); }

/* Expand section for questions/decisions */
.card-expand {
  display: none;
  border-top: 1px solid var(--border);
  padding-top: 10px;
  margin-top: 4px;
}

.card-expand.open { display: block; }

.note-section { margin-bottom: 8px; }

.note-label {
  font-size: 10px;
  font-weight: 600;
  letter-spacing: 0.05em;
  text-transform: uppercase;
  color: var(--text-muted);
  margin-bottom: 4px;
}

.note-label.questions { color: var(--purple); }
.note-label.decisions { color: var(--green); }

.note-list { display: flex; flex-direction: column; gap: 3px; margin-bottom: 6px; }

.note-item {
  font-size: 12px;
  color: var(--text-primary);
  display: flex;
  align-items: flex-start;
  gap: 6px;
}

.note-dot {
  width: 5px;
  height: 5px;
  border-radius: 50%;
  flex-shrink: 0;
  margin-top: 5px;
}

.note-dot.q { background: var(--purple); }
.note-dot.d { background: var(--green); }

.note-input-row {
  display: flex;
  gap: 6px;
}

.note-input {
  flex: 1;
  font-family: var(--font);
  font-size: 11px;
  padding: 5px 8px;
  border: 1px solid var(--border);
  border-radius: var(--radius-sm);
  background: var(--bg-page);
  color: var(--text-primary);
  outline: none;
}

.note-input:focus { border-color: var(--accent); background: var(--bg-card); }

.note-add-btn {
  width: 22px;
  height: 22px;
  border-radius: 50%;
  border: 1px solid var(--border-hover);
  background: transparent;
  cursor: pointer;
  font-size: 14px;
  color: var(--text-muted);
  display: flex;
  align-items: center;
  justify-content: center;
  font-family: var(--font);
}

.note-add-btn:hover { background: var(--accent-dim); color: var(--accent); }

/* Edit mode */
.card-edit-input {
  width: 100%;
  font-family: var(--font);
  font-size: 13px;
  font-weight: 600;
  padding: 4px 8px;
  border: 1px solid var(--accent);
  border-radius: var(--radius-sm);
  outline: none;
  color: var(--text-primary);
}

.card-edit-textarea {
  width: 100%;
  font-family: var(--font);
  font-size: 12px;
  padding: 4px 8px;
  border: 1px solid var(--accent);
  border-radius: var(--radius-sm);
  outline: none;
  color: var(--text-primary);
  resize: vertical;
  min-height: 40px;
}

/* Parking lot sidebar */
.park-sidebar {
  width: 220px;
  flex-shrink: 0;
  background: var(--bg-sidebar);
  border-left: 1px solid var(--border);
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

.park-header {
  padding: 12px 14px;
  border-bottom: 1px solid var(--border);
  display: flex;
  align-items: center;
  gap: 8px;
}

.park-title {
  font-size: 11px;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 0.06em;
  color: var(--text-muted);
}

.park-count {
  font-size: 11px;
  color: var(--text-muted);
  font-family: var(--font-mono);
  margin-left: auto;
}

.park-list {
  flex: 1;
  overflow-y: auto;
  padding: 10px;
}

.park-list::-webkit-scrollbar { width: 3px; }
.park-list::-webkit-scrollbar-thumb { background: var(--border-hover); border-radius: 2px; }

.park-chip {
  padding: 8px 10px;
  background: var(--bg-card);
  border: 1px solid var(--border);
  border-radius: var(--radius-sm);
  margin-bottom: 6px;
  font-size: 12px;
  font-weight: 500;
  color: var(--text-primary);
  display: flex;
  align-items: center;
  gap: 8px;
}

.park-chip-unpark {
  font-size: 10px;
  color: var(--text-muted);
  cursor: pointer;
  margin-left: auto;
  padding: 2px 6px;
  border: 1px solid var(--border);
  border-radius: 4px;
  background: transparent;
  font-family: var(--font);
}

.park-chip-unpark:hover {
  color: var(--accent);
  border-color: var(--accent-border);
}

.park-empty {
  padding: 16px 8px;
  text-align: center;
  font-size: 11px;
  color: var(--text-muted);
}
```

- [ ] **Step 2: Add Vote view HTML**

Replace the empty `<div class="view" id="view-vote"></div>` with:

```html
<div class="view" id="view-vote">
  <div class="vote-layout">
    <div class="vote-toolbar">
      <span class="vote-toolbar-info">Total votes: <strong id="total-votes">0</strong></span>
      <div style="margin-left:auto;display:flex;gap:8px">
        <button class="btn" onclick="downloadExcel()">Export Excel</button>
        <button class="btn" onclick="downloadSessionHTML()">Export HTML</button>
        <button class="btn btn-primary" onclick="showView('prioritise')">Prioritise →</button>
      </div>
    </div>
    <div class="vote-body">
      <div class="vote-main" id="vote-main"></div>
      <div class="park-sidebar">
        <div class="park-header">
          <span class="park-title">Parking lot</span>
          <span class="park-count" id="park-count">0</span>
        </div>
        <div class="park-list" id="park-list">
          <div class="park-empty">Park topics to revisit later</div>
        </div>
      </div>
    </div>
  </div>
</div>
```

- [ ] **Step 3: Add Vote view JavaScript**

Add to the `<script>` block:

```javascript
// ─── VOTE VIEW ───────────────────────────────────────────────────
function renderVoteView() {
  const main = document.getElementById('vote-main');
  const activeTopics = state.topics.filter(t => !t.parked);

  if (!activeTopics.length) {
    main.innerHTML = '<div style="padding:60px;text-align:center;color:var(--text-muted)"><div style="font-size:14px;font-weight:500;color:var(--text-secondary)">No topics yet</div><div style="margin-top:4px;font-size:12px">Go to Before and import cards first.</div></div>';
    renderParkList();
    updateTotalVotes();
    return;
  }

  // Group by category, sort cards within each by votes desc
  const groups = {};
  activeTopics.forEach(t => {
    const cat = t.category || 'Other';
    if (!groups[cat]) groups[cat] = [];
    groups[cat].push(t);
  });
  Object.values(groups).forEach(cards => cards.sort((a, b) => b.votes - a.votes));

  main.innerHTML = Object.entries(groups).map(([cat, cards]) => {
    const colour = getCategoryColour(cat);
    return `<div class="cat-group">
      <div class="cat-header" style="background:${colour.bg}">
        ${escapeHtml(cat)} <span class="cat-header-count">${cards.length}</span>
      </div>
      <div class="vote-cards">
        ${cards.map(t => voteCardHTML(t)).join('')}
      </div>
    </div>`;
  }).join('');

  renderParkList();
  updateTotalVotes();
}

function voteCardHTML(t) {
  const colour = getCategoryColour(t.category);
  const qList = (t.questions || []).map(q => `<div class="note-item"><span class="note-dot q"></span><span>${escapeHtml(q)}</span></div>`).join('');
  const dList = (t.decisions || []).map(d => `<div class="note-item"><span class="note-dot d"></span><span>${escapeHtml(d)}</span></div>`).join('');

  return `<div class="vote-card" style="border-left-color:${colour.border}" id="vc-${t.id}">
    <div class="vote-card-top">
      <div class="vote-count ${t.votes > 0 ? 'has-votes' : ''}">${t.votes}</div>
      <div class="vote-card-body">
        <div class="vote-card-title">${escapeHtml(t.title)}</div>
        ${t.desc ? `<div class="vote-card-desc">${escapeHtml(t.desc)}</div>` : ''}
      </div>
    </div>
    <div class="vote-card-actions">
      <button class="vote-btn up" onclick="addVote('${t.id}')" title="Upvote">+</button>
      <button class="vote-btn down" onclick="removeVote('${t.id}')" ${t.votes === 0 ? 'disabled' : ''} title="Remove vote">−</button>
      <div style="flex:1"></div>
      <button class="card-icon-btn" onclick="toggleExpand('${t.id}')" title="Questions & Decisions">▾</button>
      <button class="card-icon-btn" onclick="startCardEdit('${t.id}')" title="Edit">✎</button>
      <button class="card-icon-btn park" onclick="parkTopic('${t.id}')" title="Park">P</button>
      <button class="card-icon-btn" onclick="confirmDeleteTopic('${t.id}')" title="Delete">×</button>
    </div>
    <div class="card-expand" id="expand-${t.id}">
      <div class="note-section">
        <div class="note-label questions">Open questions</div>
        <div class="note-list">${qList}</div>
        <div class="note-input-row">
          <input class="note-input" id="qinput-${t.id}" placeholder="Add a question…" onkeydown="if(event.key==='Enter')addNote('${t.id}','q')">
          <button class="note-add-btn" onclick="addNote('${t.id}','q')">+</button>
        </div>
      </div>
      <div class="note-section">
        <div class="note-label decisions">Decisions made</div>
        <div class="note-list">${dList}</div>
        <div class="note-input-row">
          <input class="note-input" id="dinput-${t.id}" placeholder="Add a decision…" onkeydown="if(event.key==='Enter')addNote('${t.id}','d')">
          <button class="note-add-btn" onclick="addNote('${t.id}','d')">+</button>
        </div>
      </div>
    </div>
  </div>`;
}

function addVote(id) {
  const t = state.topics.find(x => x.id === id);
  if (!t) return;
  t.votes++;
  markDirty();
  renderVoteView();
}

function removeVote(id) {
  const t = state.topics.find(x => x.id === id);
  if (!t || t.votes === 0) return;
  t.votes--;
  markDirty();
  renderVoteView();
}

function toggleExpand(id) {
  const el = document.getElementById('expand-' + id);
  if (el) el.classList.toggle('open');
}

function addNote(id, type) {
  const t = state.topics.find(x => x.id === id);
  if (!t) return;
  const inputId = (type === 'q' ? 'qinput-' : 'dinput-') + id;
  const input = document.getElementById(inputId);
  const val = input.value.trim();
  if (!val) return;
  if (type === 'q') t.questions.push(val);
  else t.decisions.push(val);
  input.value = '';
  markDirty();
  renderVoteView();
  // Re-open the expand after re-render
  const el = document.getElementById('expand-' + id);
  if (el) el.classList.add('open');
}

function parkTopic(id) {
  const t = state.topics.find(x => x.id === id);
  if (t) { t.parked = true; markDirty(); renderVoteView(); }
}

function unparkTopic(id) {
  const t = state.topics.find(x => x.id === id);
  if (t) { t.parked = false; markDirty(); renderVoteView(); }
}

function confirmDeleteTopic(id) {
  const t = state.topics.find(x => x.id === id);
  if (!t) return;
  if (confirm('Delete "' + t.title + '"?')) {
    state.topics = state.topics.filter(x => x.id !== id);
    state.releases.forEach(r => { r.cardIds = r.cardIds.filter(cid => cid !== id); });
    markDirty();
    renderVoteView();
  }
}

function startCardEdit(id) {
  const t = state.topics.find(x => x.id === id);
  if (!t) return;
  const card = document.getElementById('vc-' + id);
  const titleEl = card.querySelector('.vote-card-title');
  const descEl = card.querySelector('.vote-card-desc');

  titleEl.innerHTML = `<input class="card-edit-input" id="edit-title-${id}" value="${escapeHtml(t.title)}" onkeydown="if(event.key==='Enter')saveCardEdit('${id}');if(event.key==='Escape')renderVoteView()">`;
  if (descEl) {
    descEl.innerHTML = `<textarea class="card-edit-textarea" id="edit-desc-${id}" onkeydown="if(event.key==='Escape')renderVoteView()">${escapeHtml(t.desc)}</textarea>`;
  }
  document.getElementById('edit-title-' + id).focus();
}

function saveCardEdit(id) {
  const t = state.topics.find(x => x.id === id);
  if (!t) return;
  const titleInput = document.getElementById('edit-title-' + id);
  const descInput = document.getElementById('edit-desc-' + id);
  if (titleInput && titleInput.value.trim()) t.title = titleInput.value.trim();
  if (descInput) t.desc = descInput.value.trim();
  markDirty();
  renderVoteView();
}

function renderParkList() {
  const parked = state.topics.filter(t => t.parked);
  document.getElementById('park-count').textContent = parked.length;
  const list = document.getElementById('park-list');
  if (!parked.length) {
    list.innerHTML = '<div class="park-empty">Park topics to revisit later</div>';
    return;
  }
  list.innerHTML = parked.map(t =>
    `<div class="park-chip">
      <span style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escapeHtml(t.title)}</span>
      <button class="park-chip-unpark" onclick="unparkTopic('${t.id}')">Unpark</button>
    </div>`
  ).join('');
}

function updateTotalVotes() {
  const total = state.topics.reduce((sum, t) => sum + t.votes, 0);
  document.getElementById('total-votes').textContent = total;
}
```

- [ ] **Step 4: Test the Vote view**

1. Import a JSON with several topics in the Before phase
2. Click "Start voting" — cards appear grouped by category
3. Click + to add votes — count updates, card re-sorts to top of its group
4. Click ▾ to expand questions/decisions — add entries
5. Park a topic — it moves to sidebar
6. Edit a card title — saves on Enter
7. Delete a card — confirmation appears

- [ ] **Step 5: Commit**

```bash
git add index.html
git commit -m "feat: Phase 2 Vote view — voting, parking, questions, decisions, edit/delete"
```

---

### Task 5: Phase 3 — Prioritise View

**Files:**
- Modify: `index.html` (add CSS + HTML + JS for prioritisation board)

Build the prioritise view with collapsible category accordions, Unprioritised + release columns, drag-and-drop.

- [ ] **Step 1: Add Prioritise view CSS**

Add before the closing `</style>`:

```css
/* ═══ PRIORITISE VIEW ═══ */
.prio-layout {
  flex: 1;
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

.prio-toolbar {
  background: var(--bg-card);
  border-bottom: 1px solid var(--border);
  padding: 10px 20px;
  display: flex;
  gap: 10px;
  align-items: center;
  flex-shrink: 0;
}

.prio-board {
  flex: 1;
  overflow: auto;
  padding: 20px 24px;
}

.prio-board::-webkit-scrollbar { width: 4px; height: 6px; }
.prio-board::-webkit-scrollbar-thumb { background: var(--border-hover); border-radius: 3px; }

/* Column headers */
.prio-columns-header {
  display: flex;
  gap: 8px;
  margin-bottom: 12px;
  padding-left: 0;
  position: sticky;
  top: 0;
  z-index: 10;
  background: var(--bg-page);
  padding-bottom: 8px;
}

.prio-col-head {
  min-width: 180px;
  flex: 1;
  max-width: 240px;
  background: var(--bg-card);
  border: 1px solid var(--border);
  border-radius: var(--radius-md);
  padding: 8px 12px;
  display: flex;
  align-items: center;
  gap: 6px;
}

.prio-col-head.unprioritised {
  background: var(--cat-yellow);
  border-color: rgba(0,0,0,0.06);
}

.release-name-input {
  flex: 1;
  font-size: 12px;
  font-weight: 600;
  color: var(--text-primary);
  border: none;
  background: transparent;
  font-family: var(--font);
  outline: none;
  min-width: 0;
}

.release-name-input:focus { color: var(--accent); }

.col-count {
  font-size: 11px;
  color: var(--text-muted);
  font-family: var(--font-mono);
}

.col-del-btn {
  font-size: 14px;
  color: var(--text-muted);
  cursor: pointer;
  border: none;
  background: none;
  padding: 0 2px;
  font-family: var(--font);
}

.col-del-btn:hover { color: var(--grey-action); }

.add-col-btn {
  min-width: 140px;
  padding: 8px 14px;
  border: 1.5px dashed var(--border-hover);
  border-radius: var(--radius-md);
  background: transparent;
  font-size: 12px;
  color: var(--text-muted);
  cursor: pointer;
  font-family: var(--font);
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 4px;
  transition: all 0.15s;
}

.add-col-btn:hover {
  border-color: var(--accent);
  color: var(--accent);
  background: var(--accent-dim);
}

/* Category accordions */
.prio-category {
  margin-bottom: 12px;
  border: 1px solid var(--border);
  border-radius: var(--radius-md);
  overflow: hidden;
  background: var(--bg-card);
}

.prio-cat-header {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 10px 14px;
  cursor: pointer;
  user-select: none;
  transition: background 0.12s;
}

.prio-cat-header:hover { background: var(--bg-page); }

.prio-cat-chevron {
  font-size: 12px;
  color: var(--text-muted);
  transition: transform 0.2s;
}

.prio-category.collapsed .prio-cat-chevron { transform: rotate(-90deg); }

.prio-cat-name {
  font-size: 12px;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 0.04em;
  color: var(--text-secondary);
}

.prio-cat-count {
  font-size: 11px;
  color: var(--text-muted);
  font-family: var(--font-mono);
}

.prio-cat-body {
  display: flex;
  gap: 8px;
  padding: 12px;
  border-top: 1px solid var(--border);
  overflow-x: auto;
}

.prio-category.collapsed .prio-cat-body { display: none; }

/* Drop columns within accordion */
.prio-drop-col {
  min-width: 180px;
  flex: 1;
  max-width: 240px;
  min-height: 60px;
  border: 1px solid var(--border);
  border-radius: var(--radius-sm);
  padding: 8px;
  display: flex;
  flex-direction: column;
  gap: 6px;
  background: var(--bg-page);
  transition: background 0.12s, border-color 0.12s;
}

.prio-drop-col.drag-over {
  background: var(--accent-dim);
  border-color: var(--accent-border);
  border-style: dashed;
}

/* Prio cards */
.prio-card {
  background: var(--bg-card);
  border: 1px solid var(--border);
  border-left: 3px solid var(--cat-grey);
  border-radius: var(--radius-sm);
  padding: 8px 10px;
  cursor: grab;
  transition: box-shadow 0.12s, opacity 0.15s;
  user-select: none;
}

.prio-card:hover { box-shadow: 0 1px 6px rgba(0,0,0,0.06); }
.prio-card:active { cursor: grabbing; }
.prio-card.dragging { opacity: 0.35; }

.prio-card-title {
  font-size: 12px;
  font-weight: 600;
  color: var(--text-primary);
  line-height: 1.3;
}

.prio-card-votes {
  font-size: 11px;
  color: var(--text-muted);
  margin-top: 3px;
  font-family: var(--font-mono);
}
```

- [ ] **Step 2: Add Prioritise view HTML**

Replace the empty `<div class="view" id="view-prioritise"></div>` with:

```html
<div class="view" id="view-prioritise">
  <div class="prio-layout">
    <div class="prio-toolbar">
      <span style="font-size:12px;color:var(--text-muted)">Drag cards into release columns · collapse categories when done</span>
      <div style="margin-left:auto;display:flex;gap:8px">
        <button class="btn" onclick="addRelease()">+ Add release</button>
        <button class="btn" onclick="downloadExcel()">Export Excel</button>
        <button class="btn" onclick="downloadSessionHTML()">Export HTML</button>
      </div>
    </div>
    <div class="prio-board" id="prio-board"></div>
  </div>
</div>
```

- [ ] **Step 3: Add Prioritise view JavaScript**

Add to the `<script>` block:

```javascript
// ─── PRIORITISE VIEW ─────────────────────────────────────────────
let dragState = { cardId: null, fromRelease: null };

function initReleases() {
  if (state.releases.length === 0) {
    const now = new Date();
    for (let i = 0; i < 3; i++) {
      const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
      state.releases.push({
        id: 'r' + (i + 1),
        name: d.toLocaleString('default', { month: 'long', year: 'numeric' }),
        cardIds: [],
      });
    }
  }
}

function addRelease() {
  const nextMonth = new Date();
  nextMonth.setMonth(nextMonth.getMonth() + state.releases.length);
  const name = nextMonth.toLocaleString('default', { month: 'long', year: 'numeric' });
  state.releases.push({ id: 'r' + Date.now(), name, cardIds: [] });
  renderPrioView();
}

function deleteRelease(id) {
  const r = state.releases.find(x => x.id === id);
  if (r && confirm('Delete "' + r.name + '"? Cards will return to Unprioritised.')) {
    state.releases = state.releases.filter(x => x.id !== id);
    markDirty();
    renderPrioView();
  }
}

function renameRelease(id, name) {
  const r = state.releases.find(x => x.id === id);
  if (r) r.name = name;
}

function votedTopics() {
  return state.topics.filter(t => t.votes > 0 && !t.parked);
}

function placedCardIds() {
  const ids = new Set();
  state.releases.forEach(r => r.cardIds.forEach(id => ids.add(id)));
  return ids;
}

function renderPrioView() {
  initReleases();
  const board = document.getElementById('prio-board');
  const voted = votedTopics();
  const votedIds = new Set(voted.map(t => t.id));

  // Clean stale IDs from releases
  state.releases.forEach(r => { r.cardIds = r.cardIds.filter(id => votedIds.has(id)); });

  const placed = placedCardIds();
  const unplacedTopics = voted.filter(t => !placed.has(t.id));

  // Group voted topics by category
  const categories = {};
  voted.forEach(t => {
    const cat = t.category || 'Other';
    if (!categories[cat]) categories[cat] = [];
    categories[cat].push(t);
  });

  if (!voted.length) {
    board.innerHTML = '<div style="padding:60px;text-align:center;color:var(--text-muted)"><div style="font-size:14px;font-weight:500;color:var(--text-secondary)">No voted topics yet</div><div style="margin-top:4px;font-size:12px">Go to Vote and give topics at least 1 vote.</div></div>';
    return;
  }

  // Column headers
  let html = `<div class="prio-columns-header">
    <div class="prio-col-head unprioritised">
      <span class="release-name-input" style="cursor:default;font-weight:600">Unprioritised</span>
      <span class="col-count">${unplacedTopics.length}</span>
    </div>
    ${state.releases.map(r => {
      const count = r.cardIds.filter(id => votedIds.has(id)).length;
      return `<div class="prio-col-head">
        <input class="release-name-input" value="${escapeHtml(r.name)}" onchange="renameRelease('${r.id}',this.value)">
        <span class="col-count">${count}</span>
        <button class="col-del-btn" onclick="deleteRelease('${r.id}')" title="Delete column">×</button>
      </div>`;
    }).join('')}
    <button class="add-col-btn" onclick="addRelease()">+ Add release</button>
  </div>`;

  // Category accordions
  html += Object.entries(categories).map(([cat, catTopics]) => {
    const colour = getCategoryColour(cat);
    const catUnplaced = catTopics.filter(t => !placed.has(t.id));

    return `<div class="prio-category" data-cat="${escapeHtml(cat)}">
      <div class="prio-cat-header" onclick="toggleCategory(this)" style="border-left:3px solid ${colour.border}">
        <span class="prio-cat-chevron">▾</span>
        <span class="prio-cat-name">${escapeHtml(cat)}</span>
        <span class="prio-cat-count">${catTopics.length} cards</span>
      </div>
      <div class="prio-cat-body">
        <div class="prio-drop-col" data-release="unplaced" data-cat="${escapeHtml(cat)}"
          ondragover="onPrioDragOver(event)" ondragleave="onPrioDragLeave(event)" ondrop="onPrioDrop(event,'unplaced')">
          ${catUnplaced.map(t => prioCardHTML(t, 'unplaced', colour)).join('')}
        </div>
        ${state.releases.map(r => {
          const cardsHere = r.cardIds
            .map(id => catTopics.find(t => t.id === id))
            .filter(Boolean);
          return `<div class="prio-drop-col" data-release="${r.id}" data-cat="${escapeHtml(cat)}"
            ondragover="onPrioDragOver(event)" ondragleave="onPrioDragLeave(event)" ondrop="onPrioDrop(event,'${r.id}')">
            ${cardsHere.map(t => prioCardHTML(t, r.id, colour)).join('')}
          </div>`;
        }).join('')}
      </div>
    </div>`;
  }).join('');

  board.innerHTML = html;

  // Attach drag listeners to cards
  board.querySelectorAll('.prio-card').forEach(card => {
    card.addEventListener('dragstart', (e) => {
      dragState.cardId = card.dataset.id;
      dragState.fromRelease = card.dataset.release;
      setTimeout(() => card.classList.add('dragging'), 0);
      e.dataTransfer.effectAllowed = 'move';
    });
    card.addEventListener('dragend', () => card.classList.remove('dragging'));
  });
}

function prioCardHTML(t, releaseId, colour) {
  return `<div class="prio-card" draggable="true" data-id="${t.id}" data-release="${releaseId}" style="border-left-color:${colour.border}">
    <div class="prio-card-title">${escapeHtml(t.title)}</div>
    <div class="prio-card-votes">${t.votes} vote${t.votes !== 1 ? 's' : ''}</div>
  </div>`;
}

function toggleCategory(headerEl) {
  headerEl.closest('.prio-category').classList.toggle('collapsed');
}

function onPrioDragOver(e) {
  e.preventDefault();
  e.currentTarget.classList.add('drag-over');
  e.dataTransfer.dropEffect = 'move';
}

function onPrioDragLeave(e) {
  e.currentTarget.classList.remove('drag-over');
}

function onPrioDrop(e, targetRelease) {
  e.preventDefault();
  e.currentTarget.classList.remove('drag-over');
  if (!dragState.cardId) return;

  // Remove from previous release
  if (dragState.fromRelease && dragState.fromRelease !== 'unplaced') {
    const r = state.releases.find(x => x.id === dragState.fromRelease);
    if (r) r.cardIds = r.cardIds.filter(id => id !== dragState.cardId);
  }

  // Add to target release
  if (targetRelease !== 'unplaced') {
    const r = state.releases.find(x => x.id === targetRelease);
    if (r && !r.cardIds.includes(dragState.cardId)) r.cardIds.push(dragState.cardId);
  }

  dragState.cardId = null;
  dragState.fromRelease = null;
  markDirty();
  renderPrioView();
}
```

- [ ] **Step 4: Test the Prioritise view**

1. Have topics with votes from the Vote phase
2. Switch to Prioritise — see categories expanded with cards in Unprioritised column
3. Drag a card from Unprioritised to a release column — card moves
4. Click a category header — it collapses/expands
5. Click "Add release" — new column appears
6. Edit a column name — saves on change
7. Delete a column — cards return to Unprioritised

- [ ] **Step 5: Commit**

```bash
git add index.html
git commit -m "feat: Phase 3 Prioritise view — collapsible categories, drag-and-drop, release columns"
```

---

### Task 6: Export Functions (HTML + Excel)

**Files:**
- Modify: `index.html` (add export logic to script)

Build both export functions — HTML produces a self-contained archive (re-importable), Excel uses SheetJS.

- [ ] **Step 1: Add export JavaScript**

Add to the `<script>` block:

```javascript
// ─── EXPORTS ─────────────────────────────────────────────────────
function downloadSessionHTML() {
  const stamp = state.sessionDate;
  const name = state.sessionName;
  const voted = votedTopics().sort((a, b) => b.votes - a.votes);
  const placed = placedCardIds();
  const unplaced = voted.filter(t => !placed.has(t.id));

  const releaseRows = state.releases.map(r => {
    const cards = r.cardIds.map(id => state.topics.find(t => t.id === id)).filter(Boolean);
    return `<h3 style="font-size:14px;margin:20px 0 8px">${escapeHtml(r.name)}</h3>
    <table style="width:100%;border-collapse:collapse;border:1px solid #e5e7eb;border-radius:8px;overflow:hidden;margin-bottom:12px">
    <thead><tr><th style="padding:6px 10px;background:#f9fafb;font-size:11px;text-align:left;color:#6b7280;border-bottom:1px solid #e5e7eb">Topic</th><th style="padding:6px 10px;background:#f9fafb;font-size:11px;text-align:left;color:#6b7280;border-bottom:1px solid #e5e7eb">Category</th><th style="padding:6px 10px;background:#f9fafb;font-size:11px;text-align:left;color:#6b7280;border-bottom:1px solid #e5e7eb">Votes</th></tr></thead>
    <tbody>${cards.length ? cards.map(c => `<tr><td style="padding:6px 10px;font-size:13px;border-bottom:1px solid #f3f4f6">${escapeHtml(c.title)}</td><td style="padding:6px 10px;font-size:12px;color:#6b7280;border-bottom:1px solid #f3f4f6">${escapeHtml(c.category)}</td><td style="padding:6px 10px;font-size:13px;font-family:monospace;border-bottom:1px solid #f3f4f6">${c.votes}</td></tr>`).join('') : '<tr><td colspan="3" style="padding:10px;color:#9ca3af;font-size:12px">No items assigned</td></tr>'}</tbody>
    </table>`;
  }).join('');

  const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>${escapeHtml(name)} — Session Archive (${stamp})</title>
<meta name="sr-archive" content="true">
<meta name="sr-date" content="${stamp}">
<meta name="sr-meeting" content="${escapeHtml(name)}">
<style>
body{font-family:system-ui,sans-serif;max-width:860px;margin:0 auto;padding:40px 32px;color:#111827;line-height:1.6}
h1{font-size:20px;font-weight:600;margin:0 0 4px}
h2{font-size:12px;font-weight:600;margin:28px 0 10px;padding-bottom:6px;border-bottom:2px solid #f3f4f6;text-transform:uppercase;letter-spacing:.05em;color:#6b7280}
table{width:100%;border-collapse:collapse;border:1px solid #e5e7eb;border-radius:8px;overflow:hidden}
th{padding:8px 12px;background:#f9fafb;font-size:11px;font-weight:600;text-align:left;color:#6b7280;text-transform:uppercase;letter-spacing:.05em;border-bottom:1px solid #e5e7eb}
td{padding:8px 12px;border-bottom:1px solid #f3f4f6;font-size:13px}
tr:last-child td{border-bottom:none}
</style>
</head>
<body>
<div style="background:#f9fafb;border:1px solid #e5e7eb;border-radius:12px;padding:20px 24px;margin-bottom:28px">
  <div style="font-size:10px;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:#9ca3af;margin-bottom:4px">Roadmap Workshop — Session Archive</div>
  <h1>${escapeHtml(name)}</h1>
  <div style="font-size:13px;color:#6b7280;margin-top:4px">Date: ${stamp} · ${voted.length} topics voted · ${state.releases.length} releases</div>
</div>

<h2>Vote Results</h2>
<table>
<thead><tr><th>#</th><th>Topic</th><th>Description</th><th>Category</th><th>Votes</th><th>Open Questions</th><th>Decisions</th></tr></thead>
<tbody>${voted.map((t, i) => `<tr><td style="color:#9ca3af;font-family:monospace">${i + 1}</td><td style="font-weight:500">${escapeHtml(t.title)}</td><td style="color:#6b7280">${escapeHtml(t.desc)}</td><td style="color:#6b7280">${escapeHtml(t.category)}</td><td style="font-family:monospace;font-weight:600;color:#16a34a">${t.votes}</td><td style="color:#6b7280;font-size:12px">${(t.questions || []).map(q => escapeHtml(q)).join('<br>')}</td><td style="color:#6b7280;font-size:12px">${(t.decisions || []).map(d => escapeHtml(d)).join('<br>')}</td></tr>`).join('')}
</tbody></table>

<h2>Release Plan</h2>
${releaseRows}
${unplaced.length ? `<h3 style="font-size:14px;margin:20px 0 8px;color:#6b7280">Unplaced</h3><ul style="padding-left:20px">${unplaced.map(t => `<li style="padding:3px 0;font-size:13px">${escapeHtml(t.title)} — ${t.votes} vote${t.votes !== 1 ? 's' : ''}</li>`).join('')}</ul>` : ''}

<h2>All Topics</h2>
<table>
<thead><tr><th>Topic</th><th>Description</th><th>Category</th><th>Votes</th></tr></thead>
<tbody>${state.topics.map(t => `<tr><td style="font-weight:500">${escapeHtml(t.title)}</td><td style="color:#6b7280">${escapeHtml(t.desc)}</td><td style="color:#6b7280">${escapeHtml(t.category)}</td><td style="font-family:monospace">${t.votes}</td></tr>`).join('')}
</tbody></table>

<div style="margin-top:40px;padding-top:20px;border-top:1px solid #f3f4f6;font-size:11px;color:#d1d5db">
  Generated by Roadmap Workshop · Upload this file to reuse topics in your next session.
</div>
</body>
</html>`;

  const blob = new Blob([html], { type: 'text/html' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `roadmap-session-${stamp}.html`;
  a.click();
  URL.revokeObjectURL(a.href);

  state.hasUnsavedChanges = false;
  state.lastExportTime = Date.now();
}

function downloadExcel() {
  const stamp = state.sessionDate;
  const voted = votedTopics().sort((a, b) => b.votes - a.votes);
  const placed = placedCardIds();
  const unplaced = voted.filter(t => !placed.has(t.id));

  function buildAndSave() {
    const XLSX = window.XLSX;

    const voteRows = [['#', 'Topic', 'Description', 'Category', 'Votes', 'Open Questions', 'Decisions']];
    voted.forEach((t, i) => voteRows.push([
      i + 1, t.title, t.desc, t.category, t.votes,
      (t.questions || []).join(' | '),
      (t.decisions || []).join(' | ')
    ]));

    const planRows = [['Release', 'Topic', 'Category', 'Votes', 'Open Questions', 'Decisions']];
    state.releases.forEach(r => {
      const cards = r.cardIds.map(id => state.topics.find(t => t.id === id)).filter(Boolean);
      if (cards.length) {
        cards.forEach(c => planRows.push([r.name, c.title, c.category, c.votes, (c.questions || []).join(' | '), (c.decisions || []).join(' | ')]));
      } else {
        planRows.push([r.name, '(empty)', '', '', '', '']);
      }
    });
    if (unplaced.length) {
      unplaced.forEach(t => planRows.push(['Unplaced', t.title, t.category, t.votes, (t.questions || []).join(' | '), (t.decisions || []).join(' | ')]));
    }

    const wb = XLSX.utils.book_new();
    const ws1 = XLSX.utils.aoa_to_sheet(voteRows);
    const ws2 = XLSX.utils.aoa_to_sheet(planRows);
    ws1['!cols'] = [{ wch: 4 }, { wch: 32 }, { wch: 48 }, { wch: 16 }, { wch: 8 }, { wch: 40 }, { wch: 40 }];
    ws2['!cols'] = [{ wch: 18 }, { wch: 32 }, { wch: 16 }, { wch: 8 }, { wch: 40 }, { wch: 40 }];
    XLSX.utils.book_append_sheet(wb, ws1, 'Vote Results');
    XLSX.utils.book_append_sheet(wb, ws2, 'Release Plan');
    XLSX.writeFile(wb, `roadmap-session-${stamp}.xlsx`);

    state.hasUnsavedChanges = false;
    state.lastExportTime = Date.now();
  }

  if (window.XLSX) {
    buildAndSave();
  } else {
    const s = document.createElement('script');
    s.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
    s.onload = buildAndSave;
    s.onerror = () => alert('Could not load Excel library. Check your internet connection.');
    document.head.appendChild(s);
  }
}
```

- [ ] **Step 2: Test exports**

1. Vote on some topics, add questions/decisions, place some in releases
2. Click "Export HTML" — file downloads, open it, verify it contains all data and has the `sr-archive` meta tag
3. Click "Export Excel" — file downloads, open it, verify two sheets with correct data
4. After exporting, try closing the tab — no warning (unsaved changes cleared)
5. Add another vote, try closing tab — warning appears

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "feat: HTML and Excel export with unsaved-changes tracking"
```

---

### Task 7: Polish + Final Integration Testing

**Files:**
- Modify: `index.html` (minor fixes and integration)

Final pass: ensure all three phases link together cleanly, add the topbar title binding to session name, test the full flow end-to-end.

- [ ] **Step 1: Bind topbar title to session name**

In the `initBeforeView` function, after setting the input values, add:

```javascript
document.getElementById('session-name').addEventListener('input', (e) => {
  state.sessionName = e.target.value;
  document.querySelector('.topbar-title').textContent = e.target.value || 'Roadmap Workshop';
});
```

- [ ] **Step 2: Ensure vote view re-renders when entering from Before**

In the `showView` function, the `renderVoteView()` and `renderPrioView()` calls already handle this. Verify they work after importing cards.

- [ ] **Step 3: Full end-to-end test**

Create a test JSON file:

```json
{
  "sessionName": "Q3 Strategy Planning",
  "date": "2026-06-10",
  "topics": [
    {"id": "t1", "title": "Self-service reporting", "desc": "Let stakeholders build their own dashboards", "category": "Platform"},
    {"id": "t2", "title": "API rate limiting", "desc": "Protect services from abuse", "category": "Platform"},
    {"id": "t3", "title": "Onboarding wizard redesign", "desc": "Reduce time-to-first-value for new users", "category": "UX"},
    {"id": "t4", "title": "Accessibility audit", "desc": "WCAG 2.1 AA compliance", "category": "UX"},
    {"id": "t5", "title": "Slack integration", "desc": "Push notifications to team channels", "category": "Integrations"},
    {"id": "t6", "title": "Usage analytics pipeline", "desc": "Track feature adoption across cohorts", "category": "Data"},
    {"id": "t7", "title": "Mobile responsive views", "desc": "Support tablet and phone viewports", "category": "UX"},
    {"id": "t8", "title": "Bulk import tool", "desc": "CSV upload for batch operations", "category": "Platform"},
    {"id": "t9", "title": "SSO support", "desc": "SAML and OIDC for enterprise clients", "category": "Platform"},
    {"id": "t10", "title": "Real-time dashboards", "desc": "WebSocket-powered live data updates", "category": "Data"}
  ]
}
```

Test this flow:
1. Open the app fresh — title shows "Roadmap Workshop"
2. Import the JSON — title updates to "Q3 Strategy Planning", cards show in preview with colours
3. Rename a category in preview — works
4. Delete a card in preview — works
5. Click "Start voting" — cards appear in Vote view, grouped and coloured
6. Vote on 6+ topics, add questions and decisions to a few
7. Park one topic — appears in sidebar
8. Switch to Prioritise — only voted topics appear, in collapsible categories
9. Drag cards between columns — works
10. Add a release column, rename it, delete one — works
11. Export HTML — file contains all data, re-import it in a new session works
12. Export Excel — two sheets, data correct
13. Close tab without exporting after changes — warning appears

- [ ] **Step 4: Commit**

```bash
git add index.html
git commit -m "feat: topbar session name binding and final integration"
```

---

## Summary

| Task | What it builds | Depends on |
|------|---------------|------------|
| 1 | HTML shell + CSS design system + tab nav | — |
| 2 | App state model + category colours + beforeunload | Task 1 |
| 3 | Before view (import, manual entry, preview) | Task 2 |
| 4 | Vote view (voting, parking, questions, edit/delete) | Task 3 |
| 5 | Prioritise view (accordions, drag-and-drop, columns) | Task 4 |
| 6 | Export functions (HTML + Excel) | Task 5 |
| 7 | Polish + full integration test | Task 6 |
