# broadcast-ppt

Map live worship PowerPoint slides onto the existing broadcast (chroma-key) PowerPoint templates.

Design: Python owns every file operation. Rules classify what is certain; an AI model is only a fallback classifier returning structured JSON; it never decides coordinates. Output is validated automatically and ambiguous slides go to a person, whose choices are stored for regression tests.

## Status

Phase A — structure inventory of the sample decks. Template IDs, region IDs and the rule set are decided after the inventory.

## Setup

```bash
uv sync
uv run pytest
```

## Commands

```bash
# mirror the Drive folder of decks into local/input (needs the gws CLI, signed in)
uv run python scripts/pull_drive.py <driveFolderId>

# inventory every deck: per-file JSON, summary.csv, clusters.md
uv run broadcast-ppt analyze local/input --out local/analysis

# rebuild summary.csv and clusters.md from existing JSON
uv run broadcast-ppt summarize --out local/analysis
```

## Data policy

Slide decks contain names, roles and church information. Decks, analysis JSON and requirement notes live under `local/`, which is git-ignored. Only code and generic documentation go into this repository. `clusters.md` carries counts, colours, keywords and file names, never slide text.
