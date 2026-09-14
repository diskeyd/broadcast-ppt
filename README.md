# broadcast-ppt

Convert a live worship PowerPoint deck into the broadcast (chroma-key) deck by re-using the broadcast template's slides.

Design: Python owns every file operation. Rules classify each original slide; nothing is sent to an AI model in this version. Template slides are duplicated and only the text inside their fixed text boxes is replaced — fonts, colours and box sizes stay as designed. Long text is split across several template slides whose height variant fits the estimated line count. Scripture and sermon sentences are produced twice: the original slide as a full-screen page, followed by lower-third captions. Pictures, videos and hymn score slides are copied through unchanged and flagged for review.

## Setup

```bash
uv sync
uv run pytest
```

Keep the canonical template at `local/templates/24.상반기.pptx` (see `config/scope.yaml`).

## Weekly use

```bash
# one step: plan -> build -> validate
uv run broadcast-ppt run "local/input/미편집/주일오전예배ppt(YYMMDD).pptx"

# or step by step, with a review in between
uv run broadcast-ppt plan  "<original.pptx>"          # prints a table: source slide -> kind -> template slides
uv run broadcast-ppt build "local/work/<name>/plan.json"
uv run broadcast-ppt validate "local/work/<name>/broadcast.pptx" --plan "local/work/<name>/plan.json"
```

Output: `local/work/<name>/broadcast.pptx`, `plan.json`, `validation.json`. A deck with any `critical` finding must not go on air before it is fixed.

### Review and corrections

Write `local/work/<name>/overrides.yaml` and re-run `plan --overrides` then `build`:

```yaml
slides:
  17: {skip: true}      # source slide 17 is not needed on air
  22: {op: copy}        # copy source slide 22 unchanged instead of converting it
```

Review flags printed by `plan` and `validate`:

| flag | meaning |
|---|---|
| `review:inherited_reference` | verse-numbered text without its own reference; the previous reference was used |
| `review:no_page_number` | the scripture card has no page number; the page line is left blank |
| `review:sentence_over_picture` | a short sentence on a picture was turned into a caption |
| `review:graphic` | a diagram or question sheet was copied unchanged |
| `review:unclassified` | no rule matched; the slide was copied unchanged |
| `deferred:hymn` | hymn score/lyrics images are copied unchanged (chroma handling comes later) |

## Maintenance commands

```bash
uv run broadcast-ppt analyze local/input --out local/analysis          # structure inventory of sample decks
uv run broadcast-ppt catalog local/templates/24.상반기.pptx              # rebuild config/templates.yaml after template edits
uv run python scripts/pull_drive.py <driveFolderId>                     # mirror the Drive folder into local/input
```

`config/roles.yaml` assigns template slides to roles and text slots; `config/templates.yaml` is generated from it. `config/widths.yaml` holds character widths calibrated from PowerPoint's own auto-fit boxes.

## Data policy

Slide decks contain names, roles and church information. Decks, the template, analysis output, plans, overrides, validation reports and built decks live under `local/`, which is git-ignored. Only code, generic configuration and documentation are committed. `config/templates.yaml` stores geometry, formatting, text lengths and SHA-1 digests of the template's sample text — never the text itself. Command output and reports never print slide text beyond short fit statistics.
