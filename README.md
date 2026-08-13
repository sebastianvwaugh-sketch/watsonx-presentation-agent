# IBM Presentation Generation System
### watsonx Orchestrate — Deployable Package

A complete, self-contained package for building professional IBM-branded PowerPoint
presentations using natural language. Built with the watsonx Orchestrate ADK.

---

## Package contents

```
wxo-pptx-package/
│
├── agents/
│   ├── ibm_presentation_generator.yaml     Primary dual-engine agent
│   └── file_library_agent.yaml             Sources & deliverables manager
│
├── tools/
│   ├── plan_presentation.py                Tool: plan_presentation     (Engine A)
│   ├── create_presentation.py              Tool: create_presentation   (Engine A)
│   ├── create_premium_presentation.py      Tool: create_premium_presentation (Engine B)
│   ├── sources_tools.py                    Tools: 7× file management tools
│   ├── pptx_workflow.py                    Core PPTX engine (bundled, not a tool)
│   ├── ibm_presentation_generator.py       Standalone Python class (local dev / testing)
│   ├── AI_Accelerated_PT_Method_
│   │   Presentation_Template_v1.0.13.pptx  IBM official template (bundled, not a tool)
│   ├── requirements.txt                    Tool runtime dependencies
│   └── __init__.py                         Package init
│
├── sources/
│   └── README.md                           Drop logos, briefs, data files here
│
├── deliverables/
│   └── README.md                           Generated .pptx files land here
│
├── import_all.ps1                          One-shot import script (Windows / PowerShell)
├── import_all.sh                           One-shot import script (Mac / Linux / bash)
├── requirements.txt                        Top-level dev dependencies
├── .env.example                            Environment variable template
└── README.md                               This file
```

---

## Prerequisites

1. **Python 3.9+** installed
2. **watsonx Orchestrate ADK CLI** installed:
   ```bash
   pip install ibm-watsonx-orchestrate
   ```
3. Access to a **watsonx Orchestrate environment** (instance URL + API key)

---

## Setup & import

### Step 1 — Configure your environment (first time only)

```bash
orchestrate env add --name env_NAME --url https://your-instance.watsonx.ibm.com
```

### Step 2 — Activate your environment

```bash
orchestrate env activate env_NAME
```

### Step 3 — Run the import script

Run from **inside** the `wxo-pptx-package/` folder:

**Windows (PowerShell):**
```powershell
.\import_all.ps1
```

**Mac / Linux (bash):**
```bash
bash import_all.sh
```

The script will:
1. Remove any existing registrations with the same names (safe to re-run)
2. Import all 4 tool source files → 10 registered tools
3. Import both agent YAML files
4. Print a verification listing of tools and agents

---

## Manual import (step-by-step)

```bash
# Tools — -p tools bundles the whole tools/ directory (pptx_workflow + template)
orchestrate tools import -k python -p tools -f tools/plan_presentation.py           -r tools/requirements.txt
orchestrate tools import -k python -p tools -f tools/create_presentation.py         -r tools/requirements.txt
orchestrate tools import -k python -p tools -f tools/create_premium_presentation.py -r tools/requirements.txt
orchestrate tools import -k python -p tools -f tools/sources_tools.py               -r tools/requirements.txt

# Agents — import file_library_agent first (ibm_presentation_generator lists it as a collaborator)
orchestrate agents import -f agents/file_library_agent.yaml
orchestrate agents import -f agents/ibm_presentation_generator.yaml
```

---

## Using the system

### Starting a presentation

Talk to **`ibm_presentation_generator`** in your watsonx Orchestrate chat:

> "Create a 10-slide consulting proposal about AI transformation for a financial services client"

> "Build a premium deck in the consulting theme covering our Q4 results"

> "I've dropped a brief and logo into sources/ — read them and build a deck"

### Managing files

Talk to **`file_library_agent`**:

> "What's in sources/?"

> "Rename that UUID file to Q4 Results Review"

> "Delete the test files from deliverables/"

---

## Two engines

| Engine | When to use | Tools called |
|--------|-------------|--------------|
| **A — IBM Template** | Default. Official 32-layout IBM .pptx template. Timelines, pictograms, stats grids, process steps, tables. | `plan_presentation` → `create_presentation` |
| **B — Premium Carbon** | "premium", "Carbon", named theme, "modern design". Programmatic. Full-bleed header bands. 7 layout types, unlimited slides. | `create_premium_presentation` |

### Engine B — 6 colour themes

| Theme | Palette | Best for |
|-------|---------|----------|
| `consulting` | Navy + cyan | Formal client work, QBRs, exec proposals |
| `carbon_dark` | Dark charcoal + electric blue | Tech, AI, innovation |
| `professional` | Purple + magenta | AI strategy, transformation |
| `carbon_light` | White + IBM blue | Clean classic IBM |
| `teal` | Teal + green | Sustainability, growth |
| `executive` | Warm charcoal + orange | C-suite briefings |

### Engine B — 7 layout types

| Type | Use for |
|------|---------|
| `title` | Cover slide |
| `content` | Heading + bullets (most common) |
| `section` | Full-bleed section divider |
| `two_column` | Comparison, before/after |
| `stats` | 2–4 hero metrics |
| `quote` | Pull quote / testimonial |
| `thank_you` | Closing slide |

---

## Engine A — 32 slide layouts

| Slides | Layout |
|--------|--------|
| 1 | Title / Cover |
| 2–3 | Agenda (dark / light) |
| 4–5 | Section header — title only |
| 6 | Section header + portrait image |
| 7 | Section header + body + images |
| 8–9 | 3- or 4-column grid with pictograms |
| 10 | Content + bullet list |
| 11–12 | Two- / three-column content |
| 13 | Stats + bullet list |
| 14 | Stats + pictograms |
| 15 | Stats + landscape image |
| 16 | 5-step process with pictograms |
| 17 | 4-item grid + square image |
| 18 | Hero stat + portrait image |
| 19 | 4-item list + portrait image |
| 20 | Stat + 2-item + portrait image |
| 21 | Bullets + image |
| 22–23 | Bold statement slide |
| 24–26 | Bullet list (3 variants) |
| 27 | Table + narrative text |
| 28 | Table + annotations |
| 29 | 7-step timeline |
| 30 | 5-item roadmap |
| 31–32 | Simple data table |
| 33–34 | Asset libraries — images & pictograms (source only, never in final deck) |

---

## 10 registered tools

| Tool | Source file | Purpose |
|------|-------------|---------|
| `plan_presentation` | `plan_presentation.py` | Returns slide structure + exact allowed shape names per slide |
| `create_presentation` | `create_presentation.py` | Generates IBM template PPTX, returns bytes |
| `create_premium_presentation` | `create_premium_presentation.py` | Generates Carbon-styled PPTX, returns bytes |
| `list_sources` | `sources_tools.py` | Lists all files in `sources/` |
| `read_source_text` | `sources_tools.py` | Reads .md / .txt / .csv / .json from `sources/` |
| `read_source_data` | `sources_tools.py` | Reads .xlsx / .xls / .csv, returns structured data + slide suggestions |
| `save_deliverable` | `sources_tools.py` | Saves PPTX to `deliverables/`, optionally injects logo |
| `list_deliverables` | `sources_tools.py` | Lists all .pptx in `deliverables/`, newest first |
| `delete_file` | `sources_tools.py` | Deletes a file from `sources/` or `deliverables/` |
| `rename_deliverable` | `sources_tools.py` | Renames a UUID file to a human-readable title |

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| `orchestrate: command not found` | `pip install ibm-watsonx-orchestrate` and check your PATH |
| Authentication error | Re-run `orchestrate env activate env_NAME` |
| `illegal shape` error on generation | Agent will auto-retry; shape names must come from `plan_presentation` output |
| Template not found at runtime | Make sure you used `-p tools` (not `-p .`) — this bundles the whole `tools/` directory |
| `file_library_agent` collaborator error | Import `file_library_agent` before `ibm_presentation_generator` |

---

## Local development / testing

`tools/ibm_presentation_generator.py` is a standalone Python class you can use
outside of watsonx Orchestrate — no ADK dependency required. Useful for testing
presentation output locally before deploying:

```python
from tools.ibm_presentation_generator import create_ibm_presentation

create_ibm_presentation(
    title="AI Strategy 2026",
    slides=[
        {"type": "title",   "title": "AI Strategy 2026", "subtitle": "IBM Consulting"},
        {"type": "content", "title": "Key Priorities",   "content": ["Point one", "Point two"]},
        {"type": "thank_you"},
    ],
    theme="consulting",
    output_path="deliverables/test.pptx"
)
```

Run the built-in smoke test (generates all 6 themes):
```bash
cd wxo-pptx-package
pip install -r requirements.txt
python tools/ibm_presentation_generator.py
```

---

*Made with IBM Bob*
