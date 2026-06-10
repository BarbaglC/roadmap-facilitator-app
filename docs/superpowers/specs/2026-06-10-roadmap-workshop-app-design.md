# Roadmap Workshop Facilitation App — Design Spec

## Overview

A single self-contained HTML file that serves as a live facilitation cockpit for roadmap strategy sessions. The facilitator drives all interactions (no multi-user/real-time). Designed for sessions with 30+ topic cards across many categories.

## Target Users

- **Facilitator (primary):** Drives the session, controls votes, captures decisions
- **Participants:** Management team, viewing a projected/shared screen — they call out votes verbally

## Core Principles

- Decision-oriented: every section produces captured decisions or open questions
- Output-first: structured exports (Excel + HTML) are always available
- No server, no API keys required in-app
- No dark mode
- Professional pastel palette — no reds or oranges

---

## Visual Design System

### Palette

| Role | Colour | Hex |
|------|--------|-----|
| Page background | Warm light grey | #f8f9fb |
| Card background | White | #ffffff |
| Card border | Subtle grey | 1px solid rgba(0,0,0,0.08) |
| Interactive accent | Medium blue | #3b82f6 |
| Vote count highlight | Green | #16a34a |
| Delete/destructive | Muted grey (with confirmation) | #6b7280 |

### Category Accent Colours

Used as 3px left-border on cards and tinted background on category headers:

| Colour | Hex | Example use |
|--------|-----|-------------|
| Soft blue | #dbeafe | Platform |
| Lavender | #ede9fe | UX |
| Mint green | #d1fae5 | Integrations |
| Soft yellow | #fef9c3 | Data & Analytics |
| Light grey | #f1f5f9 | Other / uncategorised |

Additional pastels available if more than 5 categories: light teal (#ccfbf1), light indigo (#e0e7ff), light lime (#ecfccb).

### Typography

- Font: Inter or system-ui sans-serif
- Size hierarchy: section headers (14px bold) > card titles (13px semibold) > descriptions (12px regular) > metadata (11px)

### Card Design

- Compact: title + 1-line truncated description + vote count
- 3px left border in category colour
- Rounded corners (8px)
- Subtle shadow on hover only
- Questions/decisions hidden behind expand toggle

---

## Phase 1: Before (Import & Setup)

### Purpose

Load topic cards into the app to prepare for voting.

### Input Methods (in priority order)

1. **Prepared JSON file** — generated in terminal session beforehand (primary flow)
2. **Previous session HTML** — app reads its own export format (identified by `<meta name="sr-archive">` tag)
3. **Manual entry** — add/edit/delete cards directly in-app

### JSON File Format

```json
{
  "sessionName": "Q3 Roadmap Workshop",
  "date": "2026-06-10",
  "topics": [
    {
      "id": "t1",
      "title": "Self-service analytics dashboard",
      "desc": "Enable stakeholders to build their own reports without eng support",
      "category": "Platform"
    }
  ]
}
```

### UI

- Clean upload zone (soft dashed border in light blue)
- After import: preview cards grouped by category with pastel colour coding
- Card count per category displayed (e.g. "Platform (7) · UX (4) · Data (3)")
- Can rename categories in preview
- Can delete individual cards from preview
- "Start voting" button to proceed

---

## Phase 2: Vote Topics

### Purpose

Run the live voting session — facilitator adds votes on behalf of participants, captures questions and decisions.

### Layout

- **Main area:** Cards by category (category header with pastel band + card count)
- **Right sidebar:** Parking lot for deferred topics

### Card Behaviour

- Compact display: title + 1-line description + large green vote count on left
- Upvote (+) and downvote (−) buttons — facilitator clicks once per person
- Auto-sort within category by vote count (highest first)
- Edit card text inline (title + description)
- Delete card (with confirmation — muted grey, no red)
- Expand toggle reveals open questions and decisions sections

### Per-Card Capture

- **Open questions:** Add via input field, displayed as purple-dotted list
- **Decisions made:** Add via input field, displayed as green-dotted list

### Parking Lot Sidebar

- Lavender-tinted background
- Parked items shown as minimal chips (title only)
- Unpark button to restore a card to the main grid

### Toolbar

- Total votes cast indicator
- Export Excel button
- Export HTML button
- "Go to Prioritise" button

---

## Phase 3: Prioritise Topics

### Purpose

Assign voted topics to release columns for sequencing and release planning.

### Entry Criteria

Only cards with at least 1 vote appear in this phase.

### Layout: Collapsible Category Accordions × Release Columns

- **Rows:** One collapsible accordion per category
  - Header shows: category name (pastel colour), card count, expand/collapse chevron
  - All categories start expanded on first load
  - Collapse categories you've finished to reduce clutter
- **Columns (horizontal kanban within each expanded category):**
  - First column: "Unprioritised" — all voted cards land here initially
  - Subsequent columns: release columns with editable names
  - Add/remove release columns freely

### Card Design (Prioritise View)

- Very compact: category colour left border, title, vote count badge
- Draggable between columns (drag from Unprioritised into a release)
- Draggable between releases to resequence
- Order within a column determined by drag position (manual sequencing)

### Column Behaviour

- Editable column names (click to rename)
- Delete column (cards return to Unprioritised)
- Add column button (generates next month name as default, editable)
- Columns sized to fit cards without excessive whitespace

### Toolbar

- "Add release" button
- Export Excel button
- Export HTML button

---

## Exports

### HTML Export

Produces a clean, self-contained summary document:
- Session metadata (name, date, totals)
- Vote results table (rank, topic, description, votes, questions, decisions)
- Release plan (grouped by release, showing assigned cards)
- Unplaced items section
- Contains `<meta name="sr-archive">` tag so it can be re-imported in future sessions

### Excel Export

Two sheets:
1. **Vote results:** Rank, topic, description, votes, source, open questions, decisions
2. **Release plan:** Release name, topic, category, votes, open questions, decisions

Uses SheetJS (loaded from CDN on first export).

---

## Data & State

- All state lives in memory during the session
- No local storage, no server, no persistence
- **Close-tab warning:** Browser `beforeunload` prompt if votes or decisions exist and no export has been done since last change
- Exports are the permanent record

---

## Constraints & Non-Goals

- No multi-user / real-time sync
- No dark mode
- No AI/API features in-app (extraction happens in terminal beforehand)
- No reds or oranges in the UI
- No server-side components
- Must handle 30+ cards across 6+ categories without feeling cramped
