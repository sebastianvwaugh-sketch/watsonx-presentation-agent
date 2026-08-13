#!/usr/bin/env pwsh
# import_all.ps1
# Imports the full IBM Presentation Generation system into watsonx Orchestrate.
# Run from the repo root after authenticating:
#   orchestrate env activate env_NAME

Set-StrictMode -Version Latest
$ErrorActionPreference = "SilentlyContinue"   # ignore remove errors for missing items

Write-Host ""
Write-Host "== IBM Presentation System — remove old, import clean ==" -ForegroundColor Cyan
Write-Host ""

# ── 1. Remove existing tools and agents ───────────────────────────────────────
Write-Host "[1/4] Removing existing tools and agents..." -ForegroundColor Yellow
orchestrate tools remove -n "plan_presentation"          2>$null
orchestrate tools remove -n "create_presentation"        2>$null
orchestrate tools remove -n "create_premium_presentation" 2>$null
orchestrate tools remove -n "list_sources"               2>$null
orchestrate tools remove -n "read_source_text"           2>$null
orchestrate tools remove -n "read_source_data"           2>$null
orchestrate tools remove -n "save_deliverable"           2>$null
orchestrate tools remove -n "list_deliverables"          2>$null
orchestrate tools remove -n "delete_file"                2>$null
orchestrate tools remove -n "rename_deliverable"         2>$null
orchestrate agents remove -n "ibm_presentation_generator" 2>$null
orchestrate agents remove -n "file_library_agent"         2>$null
Write-Host "  Done." -ForegroundColor Green

# ── 2. Import tools ────────────────────────────────────────────────────────────
$ErrorActionPreference = "Stop"
Write-Host ""
Write-Host "[2/4] Importing tools..." -ForegroundColor Yellow

# Presentation planning tool
orchestrate tools import -k python `
  -p tools `
  -f tools/plan_presentation.py `
  -r tools/requirements.txt

# Presentation generation tool (bundles pptx_workflow + template)
orchestrate tools import -k python `
  -p tools `
  -f tools/create_presentation.py `
  -r tools/requirements.txt

# Premium Carbon Design System presentation generator
orchestrate tools import -k python `
  -p tools `
  -f tools/create_premium_presentation.py `
  -r tools/requirements.txt

# File library tools (list_sources, read_source_text, read_source_data,
#                     save_deliverable, list_deliverables, delete_file, rename_deliverable)
orchestrate tools import -k python `
  -p tools `
  -f tools/sources_tools.py `
  -r tools/requirements.txt

Write-Host "  Tools imported." -ForegroundColor Green

# ── 3. Import agents ───────────────────────────────────────────────────────────
Write-Host ""
Write-Host "[3/4] Importing agents..." -ForegroundColor Yellow

# File library agent — depository for sources/ and deliverables/
orchestrate agents import -f agents/file_library_agent.yaml

# Presentation generator — uses file_library_agent as collaborator
orchestrate agents import -f agents/ibm_presentation_generator.yaml

Write-Host "  Agents imported." -ForegroundColor Green

# ── 4. Verify ──────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "[4/4] Verifying..." -ForegroundColor Yellow
orchestrate tools list
orchestrate agents list
Write-Host ""
Write-Host "Done." -ForegroundColor Cyan
Write-Host ""
Write-Host "Agents available:" -ForegroundColor Cyan
Write-Host "  file_library_agent          — browse sources/ and deliverables/"
Write-Host "  ibm_presentation_generator  — build presentations (template + premium Carbon)"
Write-Host ""
Write-Host "Tools registered (10):" -ForegroundColor Cyan
Write-Host "  plan_presentation, create_presentation, create_premium_presentation"
Write-Host "  list_sources, read_source_text, read_source_data, save_deliverable"
Write-Host "  list_deliverables, delete_file, rename_deliverable"
