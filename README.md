# Product Roadmap Workshop Facilitator

A single-file web app for running roadmap prioritisation workshops with your team. No installation, no server, no dependencies.

## Getting Started

1. Open `index.html` in any modern browser (Chrome, Edge, Safari, Firefox)
2. Upload your JSON file with the card data (or add cards manually)
3. Walk through the three tabs with your team

That's it. Everything runs locally in the browser. The only feature that needs internet is the Excel export (it loads a library from a CDN).

## The Three Tabs

### 1. Before (Setup)

Import your cards from a JSON file or add them manually. Each card has a title, description, category, color tag, and optional committed release label. Once you're happy with the list, click "Start Voting" to move to the next phase.

You can also export your card list as JSON at any point so you can re-import it later.

### 2. Vote Topics (Facilitated Discussion)

Cards are displayed in a grid layout for the group to discuss. The facilitator can:

- Vote cards up (click the vote button to increase priority)
- Park cards that the group decides to skip for now
- Capture open questions and decisions against each card during discussion
- Expand a card to see its full description

### 3. Prioritise (Release Planning)

A drag-and-drop board where you assign cards into release columns. Release columns are auto-created from any `committedRelease` tags in the imported data, or you can add new ones manually.

- Drag cards between release columns to re-prioritise
- Cards are grouped by category in collapsible accordion rows
- Click a card to expand and see its description
- Export the final plan as HTML or Excel when done

## Card Colors

Each card has a color tag that communicates its status at a glance:

| Color | Meaning |
|-------|---------|
| Blue | Prioritised — actively planned work |
| Grey | Backlog — acknowledged but not yet scheduled |
| Yellow | Untagged — no status assigned yet |
| Green | Done/Shipped — completed in a past release, shown for progress visibility |
| Orange | Late/Pushed — delayed or pushed to a later quarter due to unresolved dependencies |

## JSON Format

```json
{
  "sessionName": "Q3 Planning",
  "date": "2026-06-10",
  "topics": [
    {
      "id": "t1",
      "title": "Feature name",
      "desc": "Full description of the feature or user story",
      "category": "Category Name",
      "color": "blue",
      "committedRelease": "Jul-26"
    }
  ]
}
```

**Fields per card:**
- `id` — unique identifier (optional, auto-generated if missing)
- `title` — card title displayed on the board
- `desc` — full description, visible when you expand the card
- `category` — used for grouping in the prioritise view
- `color` — one of: `blue`, `grey`, `yellow`, `green`, `orange` (optional)
- `committedRelease` — release label in `MMM-YY` format, e.g. `Jul-26` (optional)

## Exports

- **JSON** — re-importable card list with all current data
- **HTML** — standalone report with the release plan, questions, decisions, and parked items
- **Excel** — multi-sheet workbook with release plan, questions, decisions, and parked items (requires internet)
