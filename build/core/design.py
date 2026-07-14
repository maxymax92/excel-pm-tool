"""Workbook design-system tokens.

This module is the single visual authority for the generated workbook.  Raw
values live in ``PALETTE``; writers consume semantic aliases from ``COLORS``
and the shared typography/size ramps below.  Keeping these values out of the
sheet writers prevents the workbook from drifting back into a collection of
one-off fills, fonts and row heights.

The system is deliberately Office-native and restrained: Aptos, quiet neutral
surfaces, one blue accent, and semantic colour only where the data has semantic
meaning.  Normal-size text/background pairs are chosen to meet WCAG 2.2 AA's
4.5:1 contrast target.
"""

# Context-agnostic global values. Writers consume semantic aliases from COLORS.
PALETTE = {
    "white": "#FFFFFF",
    "canvas": "#F5F7FA",
    "surface_subtle": "#EEF2F6",
    "surface_editable": "#F8FBFE",
    "surface_derived": "#F1F4F7",
    "ink": "#172B4D",
    "text": "#24364B",
    "text_secondary": "#4A5E73",
    "text_muted": "#5E7184",
    "border": "#C8D2DE",
    "border_strong": "#71869D",
    "brand": "#0F6CBD",
    "brand_dark": "#0B4A6F",
    "brand_tint": "#E8F3FC",
    "teal": "#0E6655",
    "slate": "#66788A",
    "success_bg": "#E6F4EA",
    "success_fg": "#137333",
    "warning_bg": "#FFF4CE",
    "warning_fg": "#7A4A00",
    "danger_bg": "#FDE7E9",
    "danger_fg": "#A4262C",
    "danger_strong": "#C42B1C",
    "info_bg": "#E8F3FC",
    "info_fg": "#0F548C",
    "info_soft_bg": "#F3F9FD",
    "info_soft_fg": "#356A92",
    # Examples use the informational blue family; yellow denotes attention.
    "sample_bg": "#E8F3FC",
    "sample_fg": "#0B4A6F",
    "today": "#5B6573",
}


# Semantic aliases used by worksheet writers and presentation QA.
COLORS = {
    # Surfaces and text.
    "canvas": PALETTE["canvas"],
    "surface": PALETTE["white"],
    "surface_subtle": PALETTE["surface_subtle"],
    "input_bg": PALETTE["surface_editable"],
    "formula_bg": PALETTE["surface_derived"],
    "formula_fg": PALETTE["text_secondary"],
    "text": PALETTE["text"],
    "text_secondary": PALETTE["text_secondary"],
    "text_muted": PALETTE["text_muted"],
    "border": PALETTE["border"],
    "border_strong": PALETTE["border_strong"],
    # Primary hierarchy and interaction.
    "header_bg": PALETTE["ink"],
    "header_fg": PALETTE["white"],
    "brand": PALETTE["brand"],
    "brand_dark": PALETTE["brand_dark"],
    "brand_tint": PALETTE["brand_tint"],
    "info_bg": PALETTE["info_bg"],
    "info_fg": PALETTE["info_fg"],
    "info_soft_bg": PALETTE["info_soft_bg"],
    "info_soft_fg": PALETTE["info_soft_fg"],
    # Semantic status pairs.
    "rag_g_bg": PALETTE["success_bg"],
    "rag_g_fg": PALETTE["success_fg"],
    "rag_a_bg": PALETTE["warning_bg"],
    "rag_a_fg": PALETTE["warning_fg"],
    "rag_r_bg": PALETTE["danger_bg"],
    "rag_r_fg": PALETTE["danger_fg"],
    "danger_strong": PALETTE["danger_strong"],
    # Example rows use a compact informational treatment.
    "example_bg": PALETTE["sample_bg"],
    "example_fg": PALETTE["sample_fg"],
    # Plan states pair each hue with a distinct glyph.
    "bar_done_bg": PALETTE["teal"],
    "bar_done_fg": PALETTE["white"],
    "bar_active_bg": PALETTE["brand"],
    "bar_active_fg": PALETTE["white"],
    "bar_plan_bg": PALETTE["slate"],
    "bar_plan_fg": PALETTE["white"],
    "bar_over_bg": PALETTE["danger_strong"],
    "bar_over_fg": PALETTE["white"],
    "bar_cancel_bg": PALETTE["surface_subtle"],
    "bar_cancel_fg": PALETTE["text_secondary"],
    "pt_next": PALETTE["ink"],
    "pt_done": PALETTE["success_fg"],
    "pt_over": PALETTE["danger_fg"],
    "today": PALETTE["today"],
    "bar_grey": PALETTE["border"],
    # Workbook navigation groups.
    "tab_view": PALETTE["brand"],
    "tab_data": PALETTE["teal"],
    "tab_sys": PALETTE["slate"],
}


TYPOGRAPHY = {
    "body_font": "Aptos",
    "display_font": "Aptos Display",
    "body": 10,
    "caption": 9,
    "table_header": 9,
    "section": 11,
    "page_title": 18,
}


# Alignment follows semantic content roles. Wrapped narrative starts at the
# top; compact fields use vertical centering. Numbers and dates align right for
# comparison, controls and timeline markers center, and text aligns left.
ALIGNMENT = {
    "text": {"align": "left", "valign": "vcenter"},
    "narrative": {"align": "left", "valign": "top"},
    "number": {"align": "right", "valign": "vcenter"},
    "date": {"align": "right", "valign": "vcenter"},
    "panel_text": {"align": "left", "valign": "top"},
    "panel_date": {"align": "right", "valign": "top"},
    "control": {"align": "center", "valign": "vcenter"},
    "axis": {"align": "center", "valign": "vcenter"},
    "metadata": {"align": "right", "valign": "vcenter"},
}


HIERARCHY = {
    1: {"font_size": 12, "bold": True, "indent": 0, "row_height": 30},
    2: {"font_size": 11, "bold": True, "indent": 1, "row_height": 28},
    3: {"font_size": 10.5, "bold": True, "indent": 2, "row_height": 26},
    4: {"font_size": 10, "bold": False, "indent": 3, "row_height": 24},
    5: {"font_size": 10, "bold": False, "indent": 4, "row_height": 24},
    6: {"font_size": 10, "bold": False, "indent": 5, "row_height": 24},
}


# Stored opening geometry for the non-macro QA twin. The release workbook also
# maximises itself in Workbook_Open so it adapts to the active screen.
WORKBOOK_WINDOW = {"x": 0, "y": 0, "width": 1920, "height": 1080}


# Points.  The 4/8/12/16/24/32 Fluent rhythm is translated into a compact set
# of Excel row roles rather than applied as literal web pixels.
ROWS = {
    "page_title": 32,
    "toolbar": 30,
    "table_header": 34,
    "data_compact": 24,
    "data_parent_1": 30,
    "data_parent_2": 28,
    "data_parent_3": 26,
    "data_wrapped": 44,
    "panel_title": 26,
    "panel_header": 30,
    "panel_body": 48,
    "axis_month": 20,
    "axis_week": 22,
}


# DrawingML action shapes. Their descriptions are also the deterministic key
# used by package_style.py to attach the VBA macro after XlsxWriter writes the
# drawing. DrawingML provides consistent flat styling on Mac and Windows.
MACRO_ACTIONS = {
    "export": {
        "cell": "A32",
        "macro": "ExportMarkdown",
        "caption": "Export to Markdown",
        "description": "Export Overview, Items and RAID as a UTF-8 Markdown brief",
        "width": 160,
        "height": 28,
        "x_offset": 0,
        "y_offset": 6,
    },
    "organise": {
        "cell": "J1",
        "macro": "OrganiseItems",
        "caption": "Organise rows",
        "description": "Sort Items into hierarchy order and rebuild row groups",
        "width": 120,
        "height": 28,
        "x_offset": 0,
        "y_offset": 6,
    },
}


SPACING = {"xs": 4, "sm": 8, "md": 12, "lg": 16, "xl": 24, "xxl": 32}


# Microsoft Office theme signature used by package and presentation QA.
OFFICE_THEME_SIGNATURE = {
    "major_font": "Aptos Display",
    "minor_font": "Aptos",
    "accent1": "156082",
    "dk2": "0E2841",
}
