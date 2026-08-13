"""
IBM Presentation Generator Tool

This tool generates an IBM-themed PowerPoint presentation from a structured request.
Returns a PPTX file that can be downloaded.
"""

import json
import os
import sys
import tempfile
from typing import List, Optional
from pydantic import BaseModel, Field
from ibm_watsonx_orchestrate.agent_builder.tools import tool

# Support both package import (local dev) and direct file import (server runtime)
_dir = os.path.dirname(os.path.abspath(__file__))
if _dir not in sys.path:
    sys.path.insert(0, _dir)
import pptx_workflow


# Template path - use __file__ to locate files in the package at runtime
TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "AI_Accelerated_PT_Method_Presentation_Template_v1.0.13.pptx")


class DeckSlide(BaseModel):
    id: str = Field(description="Unique identifier for this slide, e.g. 'slide_1'")
    slide_no: int = Field(description="Template slide number to use (1-32)")


class PresentationRequest(BaseModel):
    deck: List[DeckSlide] = Field(
        description="Ordered list of slides. Each entry has 'id' and 'slide_no'."
    )
    content: str = Field(
        description=(
            "JSON string mapping slide IDs to shape content. "
            "Keys must be the exact shape names returned by plan_presentation for that slide_no. "
            "Example: '{\"slide_1\": {\"title\": \"My Title\", \"subtitle\": \"Author\"}}'"
        )
    )


@tool(
    name="create_presentation",
    display_name="Create IBM Presentation",
    description="Generate an IBM-themed PowerPoint presentation from a complete structured request. Returns a PPTX file for download. This should be the final step after planning with plan_presentation tool."
)
def create_presentation(request: PresentationRequest) -> bytes:
    """Generate an IBM-themed PowerPoint presentation from a structured request.

    Args:
        request: Complete presentation request with deck (list of slides) and
                 content (JSON string mapping slide IDs to shape texts).

    Returns:
        Binary PPTX file content that can be downloaded.

    Raises:
        ValueError: If request format is invalid or generation fails.
        FileNotFoundError: If template is not found.
    """
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"Template file not found: {TEMPLATE_PATH}")

    # Parse content JSON string into dict
    try:
        content_dict = json.loads(request.content) if isinstance(request.content, str) else request.content
    except json.JSONDecodeError as exc:
        raise ValueError(f"'content' must be valid JSON: {exc}")

    # Build the raw dict that pptx_workflow expects
    request_dict = {
        "deck": [{"id": s.id, "slide_no": s.slide_no} for s in request.deck],
        "content": content_dict,
    }

    if len(request_dict["deck"]) == 0:
        raise ValueError("'deck' must contain at least one slide")

    # Add template path
    request_dict["template"] = TEMPLATE_PATH
    
    # Create temporary files for processing
    with tempfile.TemporaryDirectory() as temp_dir:
        # Write request to temp JSON file
        request_path = os.path.join(temp_dir, "request.json")
        with open(request_path, 'w', encoding='utf-8') as f:
            json.dump(request_dict, f, indent=2, ensure_ascii=False)
        
        # Generate output path
        output_path = os.path.join(temp_dir, "presentation.pptx")
        report_path = os.path.join(temp_dir, "report.json")
        
        try:
            # Generate the presentation
            pptx_workflow.generate_from_request(
                cli_template_path=TEMPLATE_PATH,
                request_path=request_path,
                out_pptx=output_path,
                out_report=report_path
            )
            
            # Read the generated PPTX file
            with open(output_path, 'rb') as f:
                pptx_bytes = f.read()
            
            # Check for warnings in the report
            if os.path.exists(report_path):
                with open(report_path, 'r', encoding='utf-8') as f:
                    warnings = json.load(f)
                    if warnings:
                        # Log warnings but still return the file
                        print(f"Presentation generated with {len(warnings)} warnings")
            
            # Return the file bytes - WXO will handle this as a downloadable file
            return pptx_bytes
            
        except SystemExit as e:
            # Handle validation errors from pptx_workflow
            error_msg = "Presentation generation failed"
            if os.path.exists(report_path):
                with open(report_path, 'r', encoding='utf-8') as f:
                    error_report = json.load(f)
                    if 'illegal' in error_report:
                        # Format illegal shapes error
                        illegal = error_report['illegal']
                        allowed = error_report.get('allowed_by_slide', {})
                        error_msg = f"Invalid shape names in request:\n"
                        for item in illegal:
                            slide_id = item.get('id')
                            shape = item.get('illegal_shape')
                            error_msg += f"  - Slide '{slide_id}': '{shape}' is not allowed\n"
                            if slide_id in allowed:
                                allowed_shapes = [s['shape'] for s in allowed[slide_id].get('allowed_shapes', [])]
                                error_msg += f"    Allowed shapes: {', '.join(allowed_shapes)}\n"
                    else:
                        error_msg = f"Presentation generation failed: {json.dumps(error_report, indent=2)}"
            raise ValueError(error_msg)
        except Exception as e:
            raise ValueError(f"Presentation generation failed: {str(e)}")

# Made with Bob
