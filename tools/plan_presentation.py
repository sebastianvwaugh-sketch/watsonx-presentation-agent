"""
IBM Presentation Planner Tool

This tool creates a structured presentation plan based on user requirements.
"""

from typing import List, Optional, Dict, Any
from ibm_watsonx_orchestrate.agent_builder.tools import tool

# Exact allowed shape names per slide number, extracted directly from the template.
# These are the ONLY valid keys allowed in the content dict for each slide_no.
# Using any other key will cause an "illegal shape" error in create_presentation.
SLIDE_SHAPES: Dict[int, Dict[str, str]] = {
    1:  {"title": "set_text", "subtitle": "set_text", "date": "set_text",
         "image_background_pt_placeholder": "swap_picture"},
    2:  {"title": "set_text", "agenda_list": "set_paragraphs"},
    3:  {"title": "set_text", "agenda_list": "set_paragraphs"},
    4:  {"title": "set_text"},
    5:  {"title": "set_text"},
    6:  {"title": "set_text", "section_number": "set_text",
         "image_background_pt_placeholder": "swap_picture"},
    7:  {"body_text": "set_text", "section_number": "set_text",
         "image_background_ls_placeholder": "swap_picture",
         "image_background_pt_placeholder": "swap_picture"},
    8:  {"title": "set_text",
         "item_01_title": "set_text", "item_01_body": "set_paragraphs", "item_01_pictogram": "swap_picture",
         "item_02_title": "set_text", "item_02_body": "set_paragraphs", "item_02_pictogram": "swap_picture",
         "item_03_title": "set_text", "item_03_body": "set_paragraphs", "item_03_pictogram": "swap_picture"},
    9:  {"title": "set_text",
         "item_01_title": "set_text", "item_01_body": "set_paragraphs", "item_01_pictogram": "swap_picture",
         "item_02_title": "set_text", "item_02_body": "set_paragraphs", "item_02_pictogram": "swap_picture",
         "item_03_title": "set_text", "item_03_body": "set_paragraphs", "item_03_pictogram": "swap_picture",
         "item_04_title": "set_text", "item_04_body": "set_paragraphs", "item_04_pictogram": "swap_picture"},
    10: {"heading": "set_text", "heading_eyebrow": "set_text", "body_text": "set_bullets"},
    11: {"title": "set_text", "body_text": "set_text",
         "item_01_heading": "set_text", "item_01_body": "set_text",
         "item_02_heading": "set_text", "item_02_body": "set_paragraphs"},
    12: {"title": "set_text", "body_text": "set_text",
         "item_01_heading": "set_text", "item_01_body": "set_text",
         "item_02_heading": "set_text", "item_02_body": "set_text",
         "item_03_heading": "set_text", "item_03_body": "set_text"},
    13: {"title": "set_text", "body": "set_bullets",
         "stat_01": "set_text", "stat_01_body": "set_text",
         "stat_02": "set_text", "stat_02_body": "set_text",
         "stat_03": "set_text", "stat_03_body": "set_text",
         "stat_04": "set_text", "stat_04_body": "set_text"},
    14: {"title": "set_text", "body_text": "set_text",
         "stat_01": "set_text", "stat_01_body": "set_text", "stat_01_eyebrow": "set_text", "stat_01_pictogram": "swap_picture",
         "stat_02": "set_text", "stat_02_body": "set_text", "stat_02_eyebrow": "set_text", "stat_02_pictogram": "swap_picture",
         "stat_03": "set_text", "stat_03_body": "set_text", "stat_03_eyebrow": "set_text", "stat_03_pictogram": "swap_picture"},
    15: {"title": "set_text", "title_eyebrow": "set_text", "body": "set_paragraphs",
         "stat_01": "set_text", "stat_01_body": "set_text",
         "stat_02": "set_text", "stat_02_body": "set_text",
         "stat_03": "set_text", "stat_03_body": "set_text",
         "stat_04": "set_text", "stat_04_body": "set_text",
         "image_background_ls_placeholder": "swap_picture"},
    16: {"title": "set_text",
         "step_01_title": "set_text", "step_01_pictogram": "set_text", "step_01_body": "set_text",
         "step_02_title": "set_text", "step_02_pictogram": "set_text", "step_02_body": "set_text",
         "step_03_title": "set_text", "step_03_pictogram": "set_text", "step_03_body": "set_text",
         "step_04_title": "set_text", "step_04_pictogram": "set_text", "step_04_body": "set_paragraphs",
         "step_05_title": "set_text", "step_05_pictogram": "set_text", "step_05_body": "set_text"},
    17: {"title": "set_text", "body": "set_paragraphs", "image_background_sq_placeholder": "set_text",
         "item_01_title": "set_text", "item_01_pictogram": "set_text", "item_01_body": "set_text",
         "item_02_title": "set_text", "item_02_pictogram": "set_text", "item_02_body": "set_text",
         "item_03_title": "set_text", "item_03_pictogram": "set_text", "item_03_body": "set_text",
         "item_04_title": "set_text", "item_04_pictogram": "set_text", "item_04_body": "set_text"},
    18: {"stat": "set_text", "body": "set_text",
         "image_background_pt_placeholder": "swap_picture"},
    19: {"title": "set_text", "image_background_pt_placeholder": "swap_picture",
         "item_01_heading": "set_text", "item_01_body": "set_text",
         "item_02_heading": "set_text", "item_02_body": "set_text",
         "item_03_heading": "set_text", "item_03_body": "set_text",
         "item_04_heading": "set_text", "item_04_body": "set_text"},
    20: {"stat": "set_text", "stat_body": "set_text", "image_background_pt_placeholder": "swap_picture",
         "item_01_heading": "set_text", "item_01_body": "set_text",
         "item_02_heading": "set_text", "item_02_body": "set_text"},
    21: {"title": "set_text", "body": "set_bullets", "image": "swap_picture"},
    22: {"body": "set_text", "body_eyebrow": "set_text"},
    23: {"body": "set_text", "body_eyebrow": "set_text"},
    24: {"title": "set_text", "body": "set_bullets"},
    25: {"title": "set_text", "body": "set_bullets"},
    26: {"title": "set_text", "body": "set_bullets"},
    27: {"title": "set_text", "section_header": "set_text",
         "body_01": "set_text", "body_02": "set_text", "table": "set_table"},
    28: {"title": "set_text", "section_header": "set_text",
         "body_01": "set_text", "body_02": "set_text",
         "table_01": "set_table",
         "table_annotation_body_01": "set_text", "table_annotation_body_02": "set_text"},
    29: {"title": "set_text",
         "step_01_body": "set_paragraphs", "step_01_pictogram": "swap_picture",
         "step_02_body": "set_paragraphs", "step_02_pictogram": "swap_picture",
         "step_03_body": "set_paragraphs", "step_03_pictogram": "swap_picture",
         "step_04_body": "set_paragraphs", "step_04_pictogram": "swap_picture",
         "step_05_body": "set_paragraphs", "step_05_pictogram": "swap_picture",
         "step_06_body": "set_paragraphs", "step_06_pictogram": "swap_picture",
         "step_07_body": "set_paragraphs", "step_07_pictogram": "swap_picture"},
    30: {"title": "set_text",
         "item_01_title": "set_text", "item_01_body": "set_text",
         "item_02_title": "set_text", "item_02_body": "set_text",
         "item_03_title": "set_text", "item_03_body": "set_text",
         "item_04_title": "set_text", "item_04_body": "set_text",
         "item_05_title": "set_text", "item_05_body": "set_text"},
    31: {"title": "set_text", "table": "set_table"},
    32: {"title": "set_text", "table": "set_table"},
}

# Asset shape names available on slides 33–34 (for use as swap_picture values)
IMAGE_ASSETS = [
    "image_background_pt_plaza", "image_background_pt_skyscraper", "image_background_pt_freeway",
    "image_background_pt_freeway_intersection", "image_background_pt_factory", "image_background_pt_boat",
    "image_background_pt_person", "image_background_pt_person_office",
    "image_background_pt_people_planning",
    "image_background_pt_informal_meeting_room_open_door",
    "image_background_pt_informal_meeting_room_through_window",
    "image_background_ls_turbine", "image_background_ls_sand",
]

PICTOGRAM_ASSETS = [
    "pictogram_user_person", "pictogram_user_person_circled", "pictogram_users_group",
    "pictogram_users_group_with_one_starred", "pictogram_users_high_five",
    "pictogram_goal_rocket", "pictogram_goal_mountain",
    "pictogram_ibm_bee", "pictogram_ibm_eye",
    "pictogram_transform_loop", "pictogram_transform_block",
    "pictogram_optimize_loop", "pictogram_optimize_blocks",
    "pictogram_automate_loop", "pictogram_automate_block",
    "pictogram_simplify_loop", "pictogram_simplify_many_to_one_block",
    "pictogram_adoption", "pictogram_discovery", "pictogram_evaluation",
    "pictogram_eliminate", "pictogram_prioritization",
    "pictogram_process",
]


@tool(
    name="plan_presentation",
    display_name="Plan IBM Presentation",
    description="Create a structured presentation plan based on user requirements. Analyzes the topic, purpose, audience, and key points to recommend an appropriate slide structure with IBM template slides."
)
def plan_presentation(
    topic: str,
    purpose: str,
    audience: str,
    key_points: List[str],
    sections: Optional[List[str]] = None,
    approximate_slides: Optional[int] = None
) -> Dict[str, Any]:
    """Create a structured presentation plan based on user requirements.
    
    This tool analyzes the requirements and recommends a slide structure
    with appropriate layouts and content organization using the IBM template.
    
    Args:
        topic: Main topic or title of the presentation
        purpose: Purpose of the presentation (e.g., "pitch", "update", "training", "proposal")
        audience: Target audience (e.g., "executives", "technical team", "clients", "stakeholders")
        key_points: List of key messages or points to cover in the presentation
        sections: Optional list of section names to organize the presentation
        approximate_slides: Optional target number of slides
        
    Returns:
        Dictionary containing recommended deck structure and content outline with:
        - presentation_overview: Summary of inputs
        - recommended_structure: List of slides with types, content outlines, and EXACT allowed shape names
        - slide_shape_reference: Complete allowed shapes for every slide number
        - asset_library: All valid image and pictogram asset names
        - content_guidelines: Character limits and style guidelines
        - next_steps: Instructions for using create_presentation tool
        - estimated_slides: Total number of recommended slides
    """
    
    plan = {
        'presentation_overview': {
            'topic': topic,
            'purpose': purpose,
            'audience': audience,
            'key_points': key_points,
            'sections': sections or []
        },
        'recommended_structure': [],
        'slide_shape_reference': SLIDE_SHAPES,
        'asset_library': {
            'images': IMAGE_ASSETS,
            'pictograms': PICTOGRAM_ASSETS,
            'usage_note': (
                'For swap_picture shapes, provide {"asset_slide": <33 for images, 34 for pictograms>, '
                '"asset_shape": "<exact asset name from lists above>"}. '
                'Use asset_slide 33 for image_background_* shapes, asset_slide 34 for pictogram_* shapes.'
            )
        },
        'content_guidelines': {},
        'next_steps': []
    }
    
    slide_id_counter = 1

    # Always start with title slide (Slide 1)
    plan['recommended_structure'].append({
        'id': f'slide_{slide_id_counter}',
        'slide_no': 1,
        'slide_type': 'Title Slide',
        'purpose': 'Cover slide',
        'allowed_shapes': list(SLIDE_SHAPES[1].keys()),
        'content_outline': {
            'title': topic,
            'subtitle': f'Prepared for {audience}',
            'date': 'To be filled with actual date',
            'image_background_pt_placeholder': {'asset_slide': 33, 'asset_shape': 'image_background_pt_plaza'}
        },
        'note': 'Cover slide. Use only shapes listed in allowed_shapes. Update the date field with the actual presentation date.'
    })
    slide_id_counter += 1

    # Add agenda if we have sections (Slide 3 - light variant)
    if sections and len(sections) > 1:
        agenda_items = [f'{i+1:02d}. {section}' for i, section in enumerate(sections)]
        plan['recommended_structure'].append({
            'id': f'slide_{slide_id_counter}',
            'slide_no': 3,
            'slide_type': 'Agenda',
            'purpose': 'Outline presentation sections',
            'allowed_shapes': list(SLIDE_SHAPES[3].keys()),
            'content_outline': {
                'title': 'Agenda',
                'agenda_list': agenda_items
            },
            'note': 'Agenda slide. Only "title" (set_text) and "agenda_list" (list of strings) are allowed.'
        })
        slide_id_counter += 1

    # Add section headers and content slides
    if sections:
        for i, section in enumerate(sections):
            # Section header (Slide 5 - title only)
            plan['recommended_structure'].append({
                'id': f'slide_{slide_id_counter}',
                'slide_no': 5,
                'slide_type': 'Section Header',
                'purpose': f'Introduce section: {section}',
                'allowed_shapes': list(SLIDE_SHAPES[5].keys()),
                'content_outline': {
                    'title': section[:25]
                },
                'note': 'Section header. Only "title" is allowed on slide 5.'
            })
            slide_id_counter += 1

            # Add 2 content slides per section (slide 10: heading + bullets)
            for j in range(2):
                plan['recommended_structure'].append({
                    'id': f'slide_{slide_id_counter}',
                    'slide_no': 10,
                    'slide_type': 'Content with Bullets',
                    'purpose': f'Content for {section}',
                    'allowed_shapes': list(SLIDE_SHAPES[10].keys()),
                    'content_outline': {
                        'heading': f'{section} - Point {j+1}',
                        'heading_eyebrow': section,
                        'body_text': ['Add bullet point 1 here', 'Add bullet point 2 here', 'Add bullet point 3 here']
                    },
                    'note': (
                        'Slide 10 allows only: heading (set_text), heading_eyebrow (set_text), body_text (list of bullet strings). '
                        'Do NOT use: title, caption, stat_01, pictogram, or any other shape name.'
                    )
                })
                slide_id_counter += 1
    else:
        # No sections - add content slides for key points
        for i, point in enumerate(key_points[:5]):
            plan['recommended_structure'].append({
                'id': f'slide_{slide_id_counter}',
                'slide_no': 10,
                'slide_type': 'Content with Bullets',
                'purpose': f'Cover key point: {point}',
                'allowed_shapes': list(SLIDE_SHAPES[10].keys()),
                'content_outline': {
                    'heading': point[:55],
                    'body_text': ['Add supporting detail here']
                },
                'note': 'Slide 10 allows only: heading, heading_eyebrow, body_text.'
            })
            slide_id_counter += 1

    plan['content_guidelines'] = {
        'CRITICAL_RULE': (
            'You MUST only use shape names listed in the "allowed_shapes" field for each slide. '
            'The complete reference for every slide number is in slide_shape_reference. '
            'Using any shape name not in that list will cause an error.'
        ),
        'character_limits': {
            'title': '50-70 characters',
            'heading': '10-32 characters',
            'body_paragraph': '300-650 characters',
            'bullet_point': '1-2 lines each, 3-6 bullets per slide',
            'stat_value': '2-10 characters (use K/M/B suffixes)',
            'section_title': 'max 25 characters'
        },
        'value_formats': {
            'set_text': 'plain string',
            'set_bullets': 'list of strings',
            'set_paragraphs': 'list of strings',
            'set_table': '{"headers": ["Col1", "Col2"], "rows": [["A", "B"]]}',
            'swap_picture': '{"asset_slide": 33, "asset_shape": "image_background_pt_plaza"}'
        },
        'style_guidelines': {
            'stats': 'Use compact suffixes (K/M/B), e.g., 2.5M, 50%, £1.2M',
            'bullets': 'Parallel structure, lead with outcome/impact',
            'tables': 'Keep cells to 1-3 lines, consistent formatting',
            'dates': '"Month Day, Year" format (e.g., "January 15, 2026")'
        }
    }

    plan['next_steps'] = [
        '1. Review the recommended_structure - each entry has an "allowed_shapes" list',
        '2. Build content using ONLY the shape names in allowed_shapes for each slide_no',
        '3. Cross-check every shape key in your content against slide_shape_reference[slide_no]',
        '4. For swap_picture shapes use: {"asset_slide": 33, "asset_shape": "<name from asset_library.images>"}',
        '5. For pictogram swap_picture shapes use: {"asset_slide": 34, "asset_shape": "<name from asset_library.pictograms>"}',
        '6. Call create_presentation with the complete deck + content request'
    ]

    plan['estimated_slides'] = len(plan['recommended_structure'])

    plan['example_request_format'] = {
        'deck': [
            {'id': 'slide_1', 'slide_no': 1},
            {'id': 'slide_2', 'slide_no': 3},
            {'id': 'slide_3', 'slide_no': 10}
        ],
        'content': {
            'slide_1': {
                'title': 'Your Presentation Title',
                'subtitle': 'Author Name',
                'date': 'January 15, 2026',
                'image_background_pt_placeholder': {'asset_slide': 33, 'asset_shape': 'image_background_pt_plaza'}
            },
            'slide_2': {
                'title': 'Agenda',
                'agenda_list': ['01. Introduction', '02. Main Content', '03. Conclusion']
            },
            'slide_3': {
                'heading': 'Key Point Heading',
                'heading_eyebrow': 'Section Name',
                'body_text': ['Bullet one', 'Bullet two', 'Bullet three']
            }
        }
    }

    return plan

# Made with Bob
