#!/usr/bin/env python3
"""
Label 4 Generator - Outer Cover Label
Generates SVG/PDF label 4 files from template and Excel dataframe input.
"""

import base64
import os
import re
import textwrap
import xml.etree.ElementTree as ET
from copy import deepcopy
from functools import lru_cache
from html import escape
from typing import List, Optional, Tuple

import pandas as pd

from generate_label2 import _configure_fontconfig, HAS_CAIROSVG, sanitize_filename
from term_config import (
    find_label4_title_match,
    get_label4_distributor_lines,
    get_label4_made_for_company,
)

# Ensure local fonts are discoverable in server environments.
_configure_fontconfig()

if HAS_CAIROSVG:
    import cairosvg

try:
    from PIL import ImageFont
except Exception:  # pragma: no cover - optional dependency in some runtimes
    ImageFont = None


ORIGIN_COUNTRY_MAP = {
    "CN": "CHINA",
    "VN": "VIETNAM",
    "KHM": "CAMBODIA",
}

# Label 4 layout constants (SVG coordinates)
LABEL_CENTER_X = 122.0
LABEL_TOP_Y = 19.71
LABEL_INNER_WIDTH = 151.03
MATERIAL_START_Y = 78.78
MATERIAL_LINE_HEIGHT = 16.08
SECTION_SPACING = 16.08
WASH_LINE_HEIGHT = 11.58
ICON_HEIGHT = 16.0
DISTRIBUTOR_LINE_HEIGHT = 8.4
DISTRIBUTOR_WRAP_MAX_CHARS = 42
MATERIAL_WRAP_MAX_CHARS = 34
MATERIAL_TEXT_MAX_WIDTH = LABEL_INNER_WIDTH - 2.0
WASH_TEXT_MAX_WIDTH = LABEL_INNER_WIDTH - 2.0

MATERIAL_CLASS = "cls-25"
PART_TITLE_BOLD_CLASS = "cls-20"
PART_TITLE_FRENCH_BOLD_CLASS = "cls-19"
SUB_PART_TITLE_CLASS = "cls-25"
DEFAULT_MATERIAL_CLASS = "cls-25"
WASH_TEXT_CLASS = "cls-18"

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
WASHING_ICON_DIR = os.path.join(SCRIPT_DIR, "template", "icons")
OUTER_COVER_TITLE_BASE_FONT_SIZE = 13.4
OUTER_COVER_TITLE_MIN_FONT_SIZE = 9.8
OUTER_COVER_TITLE_MAX_TEXT_WIDTH = LABEL_INNER_WIDTH - 2.0
OUTER_COVER_TITLE_FALLBACK_CHAR_WIDTH_FACTOR = 0.34
OUTER_COVER_MEASURE_FONT_SIZE = 240
OUTER_COVER_MEASURE_FONT_PATH = os.path.join(SCRIPT_DIR, "font", "AvenirNextCondensed-DemiBold.ttf")
MATERIAL_MEASURE_FONT_SIZE = 240
MATERIAL_MEASURE_FONT_PATHS = {
    MATERIAL_CLASS: os.path.join(SCRIPT_DIR, "font", "AvenirNextCondensed-Medium.ttf"),
    SUB_PART_TITLE_CLASS: os.path.join(SCRIPT_DIR, "font", "AvenirNextCondensed-Medium.ttf"),
    DEFAULT_MATERIAL_CLASS: os.path.join(SCRIPT_DIR, "font", "AvenirNextCondensed-Medium.ttf"),
    WASH_TEXT_CLASS: os.path.join(SCRIPT_DIR, "font", "AvenirNextCondensed-Medium.ttf"),
    PART_TITLE_BOLD_CLASS: os.path.join(SCRIPT_DIR, "font", "AvenirNextCondensed-DemiBold.ttf"),
    PART_TITLE_FRENCH_BOLD_CLASS: os.path.join(SCRIPT_DIR, "font", "AvenirNextCondensed-DemiBold.ttf"),
}
MATERIAL_FONT_SIZE_BY_CLASS = {
    MATERIAL_CLASS: 13.4,
    SUB_PART_TITLE_CLASS: 13.4,
    DEFAULT_MATERIAL_CLASS: 13.4,
    WASH_TEXT_CLASS: 8.04,
    PART_TITLE_BOLD_CLASS: 13.4,
    PART_TITLE_FRENCH_BOLD_CLASS: 13.4,
}
MATERIAL_FALLBACK_CHAR_WIDTH_BY_CLASS = {
    MATERIAL_CLASS: 0.34,
    SUB_PART_TITLE_CLASS: 0.34,
    DEFAULT_MATERIAL_CLASS: 0.34,
    WASH_TEXT_CLASS: 0.34,
    PART_TITLE_BOLD_CLASS: 0.34,
    PART_TITLE_FRENCH_BOLD_CLASS: 0.34,
}
WASH_GUIDES_WITHOUT_ICONS = {
    "air_dry",
    "do_not_place_cover_or_foam_in_washer_or_dryer",
    "do_not_remove_cover",
    "do_not_use_fabric_softeners",
    "lay_flat_to_dry",
    "spot_clean",
    "spot_clean_only",
    "note_washing_reduces_spill_resistant_finish_over_time_spot_clean_recommended",
    "professional_cleaning_only",
    "wash_separately",
}
SVG_NS = "http://www.w3.org/2000/svg"
XLINK_NS = "http://www.w3.org/1999/xlink"

ET.register_namespace("", SVG_NS)
ET.register_namespace("xlink", XLINK_NS)


def _convert_svg_to_pdf(svg_content: str) -> Optional[bytes]:
    if not HAS_CAIROSVG:
        return None
    try:
        return cairosvg.svg2pdf(bytestring=svg_content.encode("utf-8"))
    except Exception as exc:
        print(f"PDF conversion failed: {exc}")
        return None


def _build_embedded_font_faces() -> str:
    """Embed project fonts directly into SVG so server font availability is irrelevant."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    font_dir = os.path.join(script_dir, "font")

    font_specs = [
        ("AvenirNextCondensed-Medium", "AvenirNextCondensed-Medium.ttf", 500),
        ("AvenirNextCondensed-DemiBold", "AvenirNextCondensed-DemiBold.ttf", 300),
        ("AvenirNextCondensed-Bold", "AvenirNextCondensed-Bold.ttf", 700),
        ("AvenirNextCondensed-UltraLight", "AvenirNextCondensed-UltraLight.ttf", 100),
    ]

    rules = []
    for family, filename, weight in font_specs:
        font_path = os.path.join(font_dir, filename)
        if not os.path.exists(font_path):
            continue

        with open(font_path, "rb") as font_file:
            encoded = base64.b64encode(font_file.read()).decode("ascii")

        rules.append(
            "\n".join(
                [
                    "@font-face {",
                    f"    font-family: '{family}';",
                    f"    src: url('data:font/truetype;base64,{encoded}') format('truetype');",
                    f"    font-weight: {weight};",
                    "    font-style: normal;",
                    "}",
                ]
            )
        )

    return "\n\n".join(rules)


def _build_material_lines(material_text: str) -> Tuple[List[Tuple[str, str]], List[str]]:
    def _ensure_colon_suffix(text: str) -> str:
        return text if text.endswith(":") else f"{text}:"

    def _parse_segment(segment: str) -> None:
        pct_match = re.match(r"^(\d+%)\s*(.*)$", segment)
        if pct_match:
            percentage = pct_match.group(1)
            pct_has_space = bool(re.match(r"^\d+%\s+", segment))
            material_part = pct_match.group(2).strip()
        else:
            percentage = ""
            pct_has_space = False
            material_part = segment

        match_key, french, line_type = find_label4_title_match(material_part)
        if not match_key:
            inline_title_match = re.match(r"^(.*?):\s*(.+)$", material_part)
            if inline_title_match:
                inline_title = inline_title_match.group(1).strip()
                inline_remainder = inline_title_match.group(2).strip()
                inline_match_key, inline_french, inline_line_type = find_label4_title_match(inline_title)
                if inline_match_key and inline_line_type != "Material":
                    translated_title = _ensure_colon_suffix(f"{inline_match_key}({inline_french})")
                    parsed_lines.append((translated_title, SUB_PART_TITLE_CLASS if inline_line_type == "Sub-part Title" else PART_TITLE_BOLD_CLASS))
                    if inline_remainder:
                        _parse_segment(inline_remainder)
                    return

            fallback_class = MATERIAL_CLASS if percentage else DEFAULT_MATERIAL_CLASS
            parsed_lines.append((segment, fallback_class))
            return

        if line_type == "Material":
            joiner = " " if pct_has_space else ""
            bilingual_material = f"{match_key}({french})"
            text_value = f"{percentage}{joiner}{bilingual_material}" if percentage else bilingual_material
            parsed_lines.append((text_value, MATERIAL_CLASS))
            return

        if line_type == "Sub-part Title":
            combined_text = f"{match_key}({french})"
            combined_text = _ensure_colon_suffix(combined_text)
            parsed_lines.append((combined_text, SUB_PART_TITLE_CLASS))
            return

        combined_text = f"{match_key}({french})"
        combined_text = _ensure_colon_suffix(combined_text)

        # If translated part title is too long for the label width, split French to the next line.
        if len(combined_text) > MATERIAL_WRAP_MAX_CHARS:
            parsed_lines.append((match_key, PART_TITLE_BOLD_CLASS))
            french_line = _ensure_colon_suffix(f"({french})")
            parsed_lines.append((french_line, PART_TITLE_FRENCH_BOLD_CLASS))
            return

        if percentage:
            joiner = " " if pct_has_space else ""
            parsed_lines.append((f"{percentage}{joiner}{combined_text}", PART_TITLE_BOLD_CLASS))
        else:
            parsed_lines.append((combined_text, PART_TITLE_BOLD_CLASS))

    alerts: List[str] = []
    parsed_lines: List[Tuple[str, str]] = []

    normalized = material_text.replace("\\n", "\n")
    raw_segments: List[str] = []
    for raw_line in normalized.split("\n"):
        for chunk in raw_line.split(","):
            stripped = chunk.strip()
            if stripped:
                raw_segments.append(stripped)

    for segment in raw_segments:
        _parse_segment(segment)

    return parsed_lines, alerts


def _clean_instruction_display(instruction: str) -> str:
    cleaned = instruction.replace("（", "(").replace("）", ")")
    cleaned = re.sub(r"\([^)]*[\u4e00-\u9fff][^)]*\)", "", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned.upper()


def _wash_guide_to_icon_name(instruction: str) -> str:
    """
    Convert wash guide text to icon file base name.
    Rules:
    - lowercase
    - spaces and separators -> underscore
    - patterns like 30°C / 30℃ -> 30_degree
    """
    normalized = instruction.replace("（", "(").replace("）", ")")
    normalized = re.sub(r"\([^)]*\)", "", normalized)
    normalized = normalized.lower().strip()
    normalized = re.sub(r"(\d+)\s*[°º]\s*c", r"\1_degree", normalized, flags=re.IGNORECASE)
    normalized = re.sub(r"(\d+)\s*℃", r"\1_degree", normalized, flags=re.IGNORECASE)
    normalized = re.sub(r"[^a-z0-9]+", "_", normalized)
    normalized = re.sub(r"_+", "_", normalized).strip("_")
    return normalized


def _parse_washing_guides(washing_guide_text: str) -> Tuple[List[str], List[str], List[str]]:
    alerts: List[str] = []
    icon_names: List[str] = []
    text_lines: List[str] = []

    normalized = washing_guide_text.replace("\\n", "\n")
    instructions = [line.strip() for line in normalized.split("\n") if line.strip()]

    for raw_instruction in instructions:
        icon_name = _wash_guide_to_icon_name(raw_instruction)
        display_line = _clean_instruction_display(raw_instruction)
        if not icon_name:
            alerts.append(f"washing guide can not convert to icon name: '{raw_instruction}'")
            continue

        if icon_name in WASH_GUIDES_WITHOUT_ICONS:
            if display_line:
                text_lines.append(display_line)
            continue

        icon_path = os.path.join(WASHING_ICON_DIR, f"{icon_name}.svg")
        if not os.path.exists(icon_path):
            alerts.append(
                f"washing guide icon file not found: '{raw_instruction}' -> '{icon_name}.svg'"
            )
            continue

        if icon_name not in icon_names:
            icon_names.append(icon_name)

        if display_line:
            text_lines.append(display_line)

    return icon_names, text_lines, alerts


@lru_cache(maxsize=None)
def _load_icon_svg_source(icon_name: str) -> Optional[bytes]:
    icon_path = os.path.join(WASHING_ICON_DIR, f"{icon_name}.svg")
    if not os.path.exists(icon_path):
        return None

    with open(icon_path, "rb") as f:
        return f.read()


def _svg_local_name(tag: str) -> str:
    if "}" in tag:
        return tag.rsplit("}", 1)[1]
    return tag


def _svg_namespace(tag: str) -> str:
    if tag.startswith("{") and "}" in tag:
        return tag[1:].split("}", 1)[0]
    return ""


def _replace_svg_id_refs(text: str, id_map: dict[str, str]) -> str:
    updated = text
    for original, namespaced in id_map.items():
        escaped = re.escape(original)
        updated = re.sub(rf"url\(#({escaped})\)", f"url(#{namespaced})", updated)
        updated = re.sub(rf'="#({escaped})"', f'="#{namespaced}"', updated)
        updated = re.sub(rf"='#({escaped})'", f"='#{namespaced}'", updated)
    return updated


def _rewrite_svg_style_text(text: str, id_map: dict[str, str], class_map: dict[str, str]) -> str:
    updated = _replace_svg_id_refs(text, id_map)
    for original, namespaced in id_map.items():
        escaped = re.escape(original)
        updated = re.sub(
            rf"(?<![A-Za-z0-9_-])#{escaped}(?=[\s,{{:.>#]|$)",
            f"#{namespaced}",
            updated,
        )
    for original, namespaced in class_map.items():
        escaped = re.escape(original)
        updated = re.sub(
            rf"(?<![A-Za-z0-9_-])\.{escaped}(?=[\s,{{:.>#]|$)",
            f".{namespaced}",
            updated,
        )
    return updated


def _collect_svg_prefix_maps(root: ET.Element, prefix: str) -> Tuple[dict[str, str], dict[str, str]]:
    id_map: dict[str, str] = {}
    class_map: dict[str, str] = {}

    for element in root.iter():
        element_id = element.attrib.get("id")
        if element_id:
            id_map[element_id] = f"{prefix}_{element_id}"

        class_attr = element.attrib.get("class", "")
        for class_name in class_attr.split():
            if class_name and class_name not in class_map:
                class_map[class_name] = f"{prefix}_{class_name}"

    return id_map, class_map


def _apply_svg_prefixes(element: ET.Element, id_map: dict[str, str], class_map: dict[str, str]) -> None:
    element_id = element.attrib.get("id")
    if element_id in id_map:
        element.set("id", id_map[element_id])

    class_attr = element.attrib.get("class", "")
    if class_attr:
        class_names = [class_map.get(name, name) for name in class_attr.split()]
        element.set("class", " ".join(class_names))

    for attr_name, attr_value in list(element.attrib.items()):
        attr_namespace = _svg_namespace(attr_name)
        if attr_namespace and attr_namespace not in {SVG_NS, XLINK_NS}:
            del element.attrib[attr_name]
            continue
        if attr_name in {"id", "class"}:
            continue
        element.set(attr_name, _replace_svg_id_refs(attr_value, id_map))

    if _svg_local_name(element.tag) == "style" and element.text:
        element.text = _rewrite_svg_style_text(element.text, id_map, class_map)

    for child in list(element):
        _apply_svg_prefixes(child, id_map, class_map)


def _derive_svg_view_box(root: ET.Element) -> str:
    view_box = root.attrib.get("viewBox")
    if view_box:
        return view_box

    width = root.attrib.get("width", "").strip()
    height = root.attrib.get("height", "").strip()
    width_match = re.match(r"^([0-9]+(?:\.[0-9]+)?)", width)
    height_match = re.match(r"^([0-9]+(?:\.[0-9]+)?)", height)
    if width_match and height_match:
        return f"0 0 {width_match.group(1)} {height_match.group(1)}"

    return "0 0 16 16"


def _build_inline_icon_svg(icon_name: str, x: float, y: float, size: float, index: int) -> Optional[str]:
    svg_bytes = _load_icon_svg_source(icon_name)
    if not svg_bytes:
        return None

    root = ET.fromstring(svg_bytes)
    view_box = _derive_svg_view_box(root)
    prefix_base = re.sub(r"[^a-zA-Z0-9_]+", "_", icon_name).strip("_") or "icon"
    prefix = f"icon_{index + 1}_{prefix_base}"
    id_map, class_map = _collect_svg_prefix_maps(root, prefix)

    nested_svg = ET.Element(
        f"{{{SVG_NS}}}svg",
        {
            "x": f"{x:.2f}",
            "y": f"{y:.2f}",
            "width": f"{size:.2f}",
            "height": f"{size:.2f}",
            "viewBox": view_box,
            "preserveAspectRatio": "xMidYMid meet",
        },
    )

    for child in list(root):
        if _svg_namespace(child.tag) not in {"", SVG_NS}:
            continue
        cloned_child = deepcopy(child)
        _apply_svg_prefixes(cloned_child, id_map, class_map)
        nested_svg.append(cloned_child)

    return ET.tostring(nested_svg, encoding="unicode")


def _build_washing_icons(icon_names: List[str], icon_y: float) -> str:
    max_icons = 5
    icon_size = 16.0
    icon_gap = 8.0

    count = min(len(icon_names), max_icons)
    if count <= 0:
        return ""

    total_width = (count * icon_size) + ((count - 1) * icon_gap)
    start_x = LABEL_CENTER_X - (total_width / 2.0)

    snippets: List[str] = []
    for idx, icon_name in enumerate(icon_names[:count]):
        x = start_x + idx * (icon_size + icon_gap)
        inline_svg = _build_inline_icon_svg(icon_name, x=x, y=icon_y, size=icon_size, index=idx)
        if not inline_svg:
            continue
        snippets.append(inline_svg)

    return "".join(snippets)


def _build_material_tspans(lines: List[Tuple[str, str]], line_height: float = 16.08) -> str:
    result: List[str] = []
    y = 0.0
    for line, css_class in lines:
        result.append(
            f'<tspan class="{css_class}"><tspan x="0" y="{y:.2f}">{escape(line)}</tspan></tspan>'
        )
        y += line_height
    return "".join(result)


def _build_simple_tspans(lines: List[str], line_height: float, x: float = 0.0) -> str:
    result: List[str] = []
    y = 0.0
    for line in lines:
        result.append(f'<tspan x="{x:.2f}" y="{y:.2f}">{escape(line)}</tspan>')
        y += line_height
    return "".join(result)


def _wrap_distributor_lines(lines: List[str], max_chars: int = DISTRIBUTOR_WRAP_MAX_CHARS) -> List[str]:
    wrapped: List[str] = []
    for line in lines:
        normalized = re.sub(r"\s+", " ", line.strip())
        if not normalized:
            continue
        segments = textwrap.wrap(
            normalized,
            width=max_chars,
            break_long_words=False,
            break_on_hyphens=False,
        )
        if segments:
            wrapped.extend(segments)
        else:
            wrapped.append(normalized)
    return wrapped


@lru_cache(maxsize=1)
def _load_outer_cover_measure_font():
    if ImageFont is None:
        return None
    if not os.path.exists(OUTER_COVER_MEASURE_FONT_PATH):
        return None
    try:
        return ImageFont.truetype(OUTER_COVER_MEASURE_FONT_PATH, OUTER_COVER_MEASURE_FONT_SIZE)
    except Exception:
        return None


@lru_cache(maxsize=None)
def _load_label4_measure_font(css_class: str):
    if ImageFont is None:
        return None
    font_path = MATERIAL_MEASURE_FONT_PATHS.get(css_class)
    if not font_path or not os.path.exists(font_path):
        return None
    try:
        return ImageFont.truetype(font_path, MATERIAL_MEASURE_FONT_SIZE)
    except Exception:
        return None


def _measure_label4_text_width(text: str, css_class: str) -> float:
    normalized = text.strip()
    if not normalized:
        return 0.0

    font = _load_label4_measure_font(css_class)
    font_size = MATERIAL_FONT_SIZE_BY_CLASS.get(css_class, 13.4)
    if font is not None:
        try:
            return float(font.getlength(normalized)) * (font_size / MATERIAL_MEASURE_FONT_SIZE)
        except Exception:
            pass

    fallback_char_width = MATERIAL_FALLBACK_CHAR_WIDTH_BY_CLASS.get(css_class, 0.34)
    return len(normalized) * fallback_char_width * font_size


def _split_label4_long_token(token: str, css_class: str, max_width: float) -> List[str]:
    chunks: List[str] = []
    current = ""

    for char in token:
        candidate = f"{current}{char}"
        if current and _measure_label4_text_width(candidate, css_class) > max_width:
            chunks.append(current)
            current = char
        else:
            current = candidate

    if current:
        chunks.append(current)

    return chunks or [token]


def _wrap_label4_text_line(text: str, css_class: str, max_width: float) -> List[str]:
    normalized = re.sub(r"\s+", " ", text.strip())
    if not normalized:
        return []

    if _measure_label4_text_width(normalized, css_class) <= max_width:
        return [normalized]

    wrapped: List[str] = []
    current = ""

    for token in normalized.split(" "):
        if not token:
            continue

        if _measure_label4_text_width(token, css_class) > max_width:
            if current:
                wrapped.append(current)
                current = ""
            wrapped.extend(_split_label4_long_token(token, css_class, max_width))
            continue

        candidate = token if not current else f"{current} {token}"
        if _measure_label4_text_width(candidate, css_class) <= max_width:
            current = candidate
        else:
            wrapped.append(current)
            current = token

    if current:
        wrapped.append(current)

    return wrapped


def _wrap_material_lines(
    lines: List[Tuple[str, str]],
    max_width: float = MATERIAL_TEXT_MAX_WIDTH,
) -> List[Tuple[str, str]]:
    wrapped: List[Tuple[str, str]] = []

    for text, css_class in lines:
        for line in _wrap_label4_text_line(text, css_class, max_width):
            wrapped.append((line, css_class))

    return wrapped


def _wrap_label4_washing_text_lines(lines: List[str]) -> List[str]:
    wrapped: List[str] = []
    for line in lines:
        wrapped.extend(_wrap_label4_text_line(line, WASH_TEXT_CLASS, WASH_TEXT_MAX_WIDTH))
    return wrapped


def _measure_outer_cover_text_width(text: str) -> float:
    font = _load_outer_cover_measure_font()
    if font is not None:
        try:
            return float(font.getlength(text)) / OUTER_COVER_MEASURE_FONT_SIZE
        except Exception:
            pass
    return len(text) * OUTER_COVER_TITLE_FALLBACK_CHAR_WIDTH_FACTOR


def _fit_outer_cover_title_font_size(text: str) -> float:
    normalized = text.strip()
    if not normalized:
        return OUTER_COVER_TITLE_BASE_FONT_SIZE

    unit_width = _measure_outer_cover_text_width(normalized)
    if unit_width <= 0:
        return OUTER_COVER_TITLE_BASE_FONT_SIZE

    estimated_width = unit_width * OUTER_COVER_TITLE_BASE_FONT_SIZE
    if estimated_width <= OUTER_COVER_TITLE_MAX_TEXT_WIDTH:
        return OUTER_COVER_TITLE_BASE_FONT_SIZE

    scaled = OUTER_COVER_TITLE_MAX_TEXT_WIDTH / unit_width
    return round(max(OUTER_COVER_TITLE_MIN_FONT_SIZE, scaled), 2)


def _get_label4_header_titles(customized_wash_label: str) -> Tuple[str, str, float, float]:
    normalized = customized_wash_label.strip()
    if normalized[:1].lower() == "y":
        title_en = "Outer Cover (Visible Surface):"
        title_fr = "Couverture Extérieure (Surface Visible)"
        return (
            title_en,
            title_fr,
            _fit_outer_cover_title_font_size(title_en),
            _fit_outer_cover_title_font_size(title_fr),
        )
    title_en = "Outer Covering"
    title_fr = "(Recouverture exterieure)"
    return (
        title_en,
        title_fr,
        OUTER_COVER_TITLE_BASE_FONT_SIZE,
        OUTER_COVER_TITLE_BASE_FONT_SIZE,
    )


def _replace_label4_variables(
    svg_content: str,
    material_tspans: str,
    washing_text_tspans: str,
    washing_icons: str,
    distributor_tspans: str,
    made_for_company: str,
    origin_country: str,
    embedded_font_faces: str,
    washing_text_y: float,
    made_for_y: float,
    made_for_company_y: float,
    distributor_title_y: float,
    distributor_y: float,
    made_in_y: float,
    label_height: float,
    label_bottom: float,
    outer_cover_title_en: str,
    outer_cover_title_fr: str,
    outer_cover_title_en_font_size: float,
    outer_cover_title_fr_font_size: float,
) -> str:
    svg_content = svg_content.replace("{{material_lines}}", material_tspans)
    svg_content = svg_content.replace("{{washing_text_lines}}", washing_text_tspans)
    svg_content = svg_content.replace("{{washing_icons}}", washing_icons)
    svg_content = svg_content.replace("{{distributor_lines}}", distributor_tspans)
    svg_content = svg_content.replace("{{made_for_company}}", escape(made_for_company))
    svg_content = svg_content.replace("{{origin_country}}", escape(origin_country))
    svg_content = svg_content.replace("{{embedded_font_faces}}", embedded_font_faces)
    svg_content = svg_content.replace("{{washing_text_y}}", f"{washing_text_y:.2f}")
    svg_content = svg_content.replace("{{made_for_y}}", f"{made_for_y:.2f}")
    svg_content = svg_content.replace("{{made_for_company_y}}", f"{made_for_company_y:.2f}")
    svg_content = svg_content.replace("{{distributor_title_y}}", f"{distributor_title_y:.2f}")
    svg_content = svg_content.replace("{{distributor_y}}", f"{distributor_y:.2f}")
    svg_content = svg_content.replace("{{made_in_y}}", f"{made_in_y:.2f}")
    svg_content = svg_content.replace("{{label_height}}", f"{label_height:.2f}")
    svg_content = svg_content.replace("{{label_bottom}}", f"{label_bottom:.2f}")
    svg_content = svg_content.replace("{{outer_cover_title_en}}", escape(outer_cover_title_en))
    svg_content = svg_content.replace("{{outer_cover_title_fr}}", escape(outer_cover_title_fr))
    svg_content = svg_content.replace(
        "{{outer_cover_title_en_font_size}}", f"{outer_cover_title_en_font_size:.2f}"
    )
    svg_content = svg_content.replace(
        "{{outer_cover_title_fr_font_size}}", f"{outer_cover_title_fr_font_size:.2f}"
    )
    return svg_content


def _last_line_baseline(start_y: float, line_count: int, line_height: float) -> float:
    if line_count <= 1:
        return start_y
    return start_y + ((line_count - 1) * line_height)


def generate_label4_from_dataframe(
    template_content: str,
    df: pd.DataFrame,
    generate_pdf: bool = True,
) -> Tuple[List[Tuple[str, bytes]], List[str]]:
    """Generate Label 4 PDFs from a dataframe."""
    columns = df.columns.tolist()

    pdf_files: List[Tuple[str, bytes]] = []
    warnings: List[str] = []

    if len(columns) < 8:
        warnings.append("label4 is not generated, reason: input file requires at least 8 columns (A-H).")
        return pdf_files, warnings

    code_col = columns[0]
    firm_col = columns[4]
    origin_col = columns[5]
    material_col = columns[6]
    washing_guide_col = columns[7]
    customized_wash_label_col = columns[8] if len(columns) > 8 else None

    embedded_fonts = _build_embedded_font_faces()

    for index, row in df.iterrows():
        identifier = str(row[code_col]) if pd.notna(row[code_col]) else f"label_{index}"
        firm = str(row[firm_col]).strip() if pd.notna(row[firm_col]) else ""
        origin_raw = str(row[origin_col]).strip() if pd.notna(row[origin_col]) else ""
        material_text = str(row[material_col]).strip() if pd.notna(row[material_col]) else ""
        washing_guide_text = str(row[washing_guide_col]).strip() if pd.notna(row[washing_guide_col]) else ""
        customized_wash_label = (
            str(row[customized_wash_label_col]).strip()
            if customized_wash_label_col and pd.notna(row[customized_wash_label_col])
            else ""
        )

        if not material_text:
            warnings.append(f"{identifier} label4 is not generated, reason: washing material (column G) is empty.")
            continue

        if not washing_guide_text:
            warnings.append(f"{identifier} label4 is not generated, reason: washing guide (column H) is empty.")
            continue

        material_lines, material_alerts = _build_material_lines(material_text)
        if material_alerts:
            for alert in material_alerts:
                warnings.append(f"{identifier} label4 is not generated, reason: {alert}")
            continue

        if not material_lines:
            warnings.append(f"{identifier} label4 is not generated, reason: no valid material lines found.")
            continue

        wrapped_material_lines = _wrap_material_lines(material_lines)
        if not wrapped_material_lines:
            warnings.append(f"{identifier} label4 is not generated, reason: no valid wrapped material lines found.")
            continue

        icon_keys, washing_text_lines, washing_alerts = _parse_washing_guides(washing_guide_text)
        if washing_alerts:
            for alert in washing_alerts:
                warnings.append(f"{identifier} label4 is not generated, reason: {alert}")
            continue

        if not icon_keys:
            warnings.append(f"{identifier} label4 is not generated, reason: no mappable washing icons found.")
            continue

        if len(icon_keys) > 5:
            warnings.append(f"{identifier} label4 warning: more than 5 washing icons; only first 5 are rendered.")

        made_for_company = get_label4_made_for_company(firm)
        distributor_lines = get_label4_distributor_lines(firm)
        if not made_for_company or not distributor_lines:
            warnings.append(f"{identifier} label4 is not generated, reason: unknown firm '{firm}'.")
            continue

        wrapped_distributor_lines = _wrap_distributor_lines(distributor_lines)
        if not wrapped_distributor_lines:
            warnings.append(f"{identifier} label4 is not generated, reason: distributor lines are empty.")
            continue

        origin_key = origin_raw.upper()
        origin_country = ORIGIN_COUNTRY_MAP.get(origin_key, origin_key)
        (
            outer_cover_title_en,
            outer_cover_title_fr,
            outer_cover_title_en_font_size,
            outer_cover_title_fr_font_size,
        ) = _get_label4_header_titles(customized_wash_label)

        material_end_y = _last_line_baseline(
            MATERIAL_START_Y,
            len(wrapped_material_lines),
            MATERIAL_LINE_HEIGHT,
        )
        washing_text_y = material_end_y + SECTION_SPACING

        wrapped_washing_text_lines = _wrap_label4_washing_text_lines(washing_text_lines)
        washing_line_count = max(len(wrapped_washing_text_lines), 1)
        washing_end_y = _last_line_baseline(
            washing_text_y,
            washing_line_count,
            WASH_LINE_HEIGHT,
        )

        icon_y = washing_end_y + SECTION_SPACING
        made_for_y = icon_y + ICON_HEIGHT + SECTION_SPACING
        made_for_company_y = made_for_y + SECTION_SPACING
        distributor_title_y = made_for_company_y + SECTION_SPACING
        distributor_y = distributor_title_y + 10.0

        distributor_end_y = _last_line_baseline(
            distributor_y,
            len(wrapped_distributor_lines),
            DISTRIBUTOR_LINE_HEIGHT,
        )

        made_in_y = distributor_end_y + SECTION_SPACING
        label_bottom = made_in_y + SECTION_SPACING
        label_height = label_bottom - LABEL_TOP_Y

        svg_content = _replace_label4_variables(
            template_content,
            material_tspans=_build_material_tspans(wrapped_material_lines),
            washing_text_tspans=_build_simple_tspans(
                wrapped_washing_text_lines,
                line_height=WASH_LINE_HEIGHT,
            ),
            washing_icons=_build_washing_icons(icon_keys, icon_y=icon_y),
            distributor_tspans=_build_simple_tspans(wrapped_distributor_lines, line_height=DISTRIBUTOR_LINE_HEIGHT),
            made_for_company=made_for_company,
            origin_country=origin_country,
            embedded_font_faces=embedded_fonts,
            washing_text_y=washing_text_y,
            made_for_y=made_for_y,
            made_for_company_y=made_for_company_y,
            distributor_title_y=distributor_title_y,
            distributor_y=distributor_y,
            made_in_y=made_in_y,
            label_height=label_height,
            label_bottom=label_bottom,
            outer_cover_title_en=outer_cover_title_en,
            outer_cover_title_fr=outer_cover_title_fr,
            outer_cover_title_en_font_size=outer_cover_title_en_font_size,
            outer_cover_title_fr_font_size=outer_cover_title_fr_font_size,
        )

        safe_name = sanitize_filename(identifier)
        pdf_filename = f"{safe_name}-label4.pdf"

        if HAS_CAIROSVG and generate_pdf:
            pdf_bytes = _convert_svg_to_pdf(svg_content)
            if pdf_bytes:
                pdf_files.append((pdf_filename, pdf_bytes))

    return pdf_files, warnings
