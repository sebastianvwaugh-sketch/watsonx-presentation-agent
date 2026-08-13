# Sources

Drop your source materials here before asking the agent to build a presentation.
The agent will automatically detect and use whatever you put here.

---

## What to drop in

| File type | Examples | How the agent uses it |
|---|---|---|
| **Client logo** | `client_logo.png`, `hsbc_logo.png` | Injected onto the title slide bottom-right |
| **IBM logo** | `ibm_logo.png` | Can be placed alongside client logo |
| **Brief / talking points** | `brief.md`, `talking_points.txt` | Agent reads and structures into slides |
| **Metrics / data** | `metrics.csv`, `kpis.txt`, `q4_data.csv` | Numbers pulled into stats slides |
| **Data files** | `pipeline.xlsx`, `results.csv` | Structured data for table slides |
| **Reference decks** | `example_deck.pptx` | Design / content reference |

---

## How to use

**Option A — Let the agent discover automatically:**
> "I've dropped files into sources/ — please check what's there and use them to build a deck on [topic]"

**Option B — Name specific files:**
> "Use the brief in sources/brief.md and the logo in sources/client_logo.png"

**Option C — Collaborative mode:**
> "Let's plan this together — I've put my metrics and brief in sources/"

---

## Logo naming tips

Any file with `logo` in the name (e.g. `client_logo.png`, `ibm_logo.svg`) is auto-detected
and injected into the title slide bottom-right. Keep files under 2 MB.

---

## After generation

Finished presentations are saved to `deliverables/` with readable names.
