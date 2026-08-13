#!/usr/bin/env bash
# import_all.sh
# Imports the full IBM Presentation Generation system into watsonx Orchestrate.
#
# PREREQUISITES
# ─────────────
#  1. Install the watsonx Orchestrate ADK CLI:
#       pip install ibm-watsonx-orchestrate
#  2. Configure your environment (first time only):
#       orchestrate env add --name env_NAME --url https://your-instance.watsonx.ibm.com
#  3. Authenticate:
#       orchestrate env activate env_NAME
#
# USAGE (run from inside wxo-pptx-package/)
# ──────────────────────────────────────────
#   orchestrate env activate env_NAME
#   bash import_all.sh

set -e

echo ""
echo "== IBM Presentation System — remove old, import clean =="
echo ""

# ── 1. Remove existing tools and agents ──────────────────────────────────────
echo "[1/4] Removing existing tools and agents..."
orchestrate tools remove -n "plan_presentation"           2>/dev/null || true
orchestrate tools remove -n "create_presentation"         2>/dev/null || true
orchestrate tools remove -n "create_premium_presentation" 2>/dev/null || true
orchestrate tools remove -n "list_sources"                2>/dev/null || true
orchestrate tools remove -n "read_source_text"            2>/dev/null || true
orchestrate tools remove -n "read_source_data"            2>/dev/null || true
orchestrate tools remove -n "save_deliverable"            2>/dev/null || true
orchestrate tools remove -n "list_deliverables"           2>/dev/null || true
orchestrate tools remove -n "delete_file"                 2>/dev/null || true
orchestrate tools remove -n "rename_deliverable"          2>/dev/null || true
orchestrate agents remove -n "ibm_presentation_generator" 2>/dev/null || true
orchestrate agents remove -n "file_library_agent"         2>/dev/null || true
echo "  Done."

# ── 2. Import tools ──────────────────────────────────────────────────────────
echo ""
echo "[2/4] Importing tools..."

orchestrate tools import -k python \
  -p tools \
  -f tools/plan_presentation.py \
  -r tools/requirements.txt

orchestrate tools import -k python \
  -p tools \
  -f tools/create_presentation.py \
  -r tools/requirements.txt

orchestrate tools import -k python \
  -p tools \
  -f tools/create_premium_presentation.py \
  -r tools/requirements.txt

orchestrate tools import -k python \
  -p tools \
  -f tools/sources_tools.py \
  -r tools/requirements.txt

echo "  Tools imported."

# ── 3. Import agents ─────────────────────────────────────────────────────────
echo ""
echo "[3/4] Importing agents..."

orchestrate agents import -f agents/file_library_agent.yaml
orchestrate agents import -f agents/ibm_presentation_generator.yaml

echo "  Agents imported."

# ── 4. Verify ────────────────────────────────────────────────────────────────
echo ""
echo "[4/4] Verifying..."
orchestrate tools list
orchestrate agents list
echo ""
echo "Done."
echo ""
echo "Agents available:"
echo "  file_library_agent          — browse sources/ and deliverables/"
echo "  ibm_presentation_generator  — build presentations (template + premium Carbon)"
echo ""
echo "Tools registered (10):"
echo "  plan_presentation, create_presentation, create_premium_presentation"
echo "  list_sources, read_source_text, read_source_data, save_deliverable"
echo "  list_deliverables, delete_file, rename_deliverable"
