import unittest

from generate_label4 import (
    MATERIAL_TEXT_MAX_WIDTH,
    _build_material_lines,
    _measure_label4_text_width,
    _wrap_material_lines,
)
from term_config import find_label4_material_match, find_label4_title_match
from term_config import LABEL4_TITLE_MAP, MATERIALS_DICT, PARTS_DICT


class Label4MaterialMappingTests(unittest.TestCase):
    def test_case_insensitive_exact_match(self) -> None:
        match_key, french = find_label4_material_match("polyester")
        self.assertEqual(match_key, "Polyester")
        self.assertEqual(french, "Polyester")

        title_key, title_french, line_type = find_label4_title_match("polyester")
        self.assertEqual(title_key, "Polyester")
        self.assertEqual(title_french, "Polyester")
        self.assertEqual(line_type, "Material")

        lines, alerts = _build_material_lines("polyester")
        self.assertEqual(lines, [("Polyester(Polyester)", "cls-25")])
        self.assertEqual(alerts, [])

    def test_material_line_with_percentage_gets_french_translation(self) -> None:
        lines, alerts = _build_material_lines("45%COTTON")
        self.assertEqual(lines, [("45%Cotton(Coton)", "cls-25")])
        self.assertEqual(alerts, [])

    def test_material_dictionary_terms_are_available_for_label4(self) -> None:
        label4_materials_by_uppercase = {
            key.upper(): meta
            for key, meta in LABEL4_TITLE_MAP.items()
            if meta["type"] == "Material"
        }

        missing_terms = [
            key
            for key in MATERIALS_DICT
            if key not in label4_materials_by_uppercase
        ]
        self.assertEqual(missing_terms, [])

        for key, french in MATERIALS_DICT.items():
            self.assertEqual(label4_materials_by_uppercase[key]["french"], french)

    def test_part_dictionary_terms_are_available_for_label4(self) -> None:
        label4_parts_by_uppercase = {
            key.upper(): meta
            for key, meta in LABEL4_TITLE_MAP.items()
            if meta["type"] == "Part Title"
        }

        missing_terms = [
            key
            for key in PARTS_DICT
            if key not in label4_parts_by_uppercase
        ]
        self.assertEqual(missing_terms, [])

        for key, french in PARTS_DICT.items():
            self.assertEqual(label4_parts_by_uppercase[key]["french"], french)

    def test_cotton_canvas_material_line_gets_french_translation(self) -> None:
        lines, alerts = _build_material_lines("100% Cotton Canvas")
        self.assertEqual(lines, [("100% Cotton Canvas(Toile de coton)", "cls-25")])
        self.assertEqual(alerts, [])

    def test_overlong_bilingual_material_line_wraps_to_fit_label_width(self) -> None:
        lines, alerts = _build_material_lines("23%VISCOSE")
        self.assertEqual(alerts, [])

        wrapped = _wrap_material_lines(lines)

        self.assertGreater(len(wrapped), 1)
        self.assertEqual(
            " ".join(line for line, _css_class in wrapped),
            "23%Viscose(La fibre artificielle de viscose)",
        )
        for line, css_class in wrapped:
            self.assertEqual(css_class, "cls-25")
            self.assertLessEqual(_measure_label4_text_width(line, css_class), MATERIAL_TEXT_MAX_WIDTH)

    def test_trailing_colon_compatibility(self) -> None:
        title_key, title_french, line_type = find_label4_title_match("PRIMARY FABRIC:")
        self.assertEqual(title_key, "Primary Fabric")
        self.assertEqual(title_french, "Tissu Principal")
        self.assertEqual(line_type, "Sub-part Title")

        lines, alerts = _build_material_lines("PRIMARY FABRIC:")
        self.assertEqual(lines, [("Primary Fabric(Tissu Principal):", "cls-25")])
        self.assertEqual(alerts, [])

    def test_no_alias_behavior(self) -> None:
        lines, alerts = _build_material_lines("Second Fabric:")
        self.assertEqual(lines, [("Second Fabric:", "cls-25")])
        self.assertEqual(alerts, [])

    def test_unknown_passthrough(self) -> None:
        lines, alerts = _build_material_lines("Custom Blend:")
        self.assertEqual(lines, [("Custom Blend:", "cls-25")])
        self.assertEqual(alerts, [])

    def test_long_part_title_wraps_french_to_next_line(self) -> None:
        lines, alerts = _build_material_lines("Seat Cushion Cover")
        self.assertEqual(
            lines,
            [
                ("Seat Cushion Cover", "cls-20"),
                ("(Housse de coussin de siege):", "cls-19"),
            ],
        )
        self.assertEqual(alerts, [])

    def test_sub_part_title_without_colon_gets_colon(self) -> None:
        lines, alerts = _build_material_lines("Secondary Fabric")
        self.assertEqual(lines, [("Secondary Fabric(Tissu Secondaire):", "cls-25")])
        self.assertEqual(alerts, [])

    def test_inline_sub_part_title_with_material_gets_french_title(self) -> None:
        lines, alerts = _build_material_lines("Primary Fabric: 100%POLYESTER")
        self.assertEqual(
            lines,
            [
                ("Primary Fabric(Tissu Principal):", "cls-25"),
                ("100%Polyester(Polyester)", "cls-25"),
            ],
        )
        self.assertEqual(alerts, [])

    def test_inline_frame_with_fabric_blend(self) -> None:
        lines, alerts = _build_material_lines("Frame:96% Polyester,4%Linen")
        self.assertEqual(
            lines,
            [
                ("Frame(House de structure):", "cls-20"),
                ("96% Polyester(Polyester)", "cls-25"),
                ("4%Linen(Lin)", "cls-25"),
            ],
        )
        self.assertEqual(alerts, [])

    def test_leather_frame_and_material_subtitles(self) -> None:
        lines, alerts = _build_material_lines(
            "Frame:100%Leather\n\u3000\nSeat Cushion:\n"
            "Primary Material:100% Leather\n"
            "Secondary Material:100% Polypropylene"
        )
        self.assertEqual(
            lines,
            [
                ("Frame(House de structure):", "cls-20"),
                ("100%Leather(Cuir)", "cls-25"),
                ("Seat Cushion(Un coussin de siège):", "cls-20"),
                ("Primary Material(Matériau principal):", "cls-25"),
                ("100% Leather(Cuir)", "cls-25"),
                ("Secondary Material(Matériau secondaire):", "cls-25"),
                ("100% Polypropylene(Polypropylène)", "cls-25"),
            ],
        )
        self.assertEqual(alerts, [])

    def test_mixed_multiline_input(self) -> None:
        text = (
            "Seat Cushion Cover\n"
            "Primary Fabric:\n"
            "100%POLYESTER\n"
            "Second Fabric:\n"
            "100%POLYESTER\n"
            "Base Cushion Cover\n"
            "Custom Blend:"
        )
        lines, alerts = _build_material_lines(text)
        self.assertEqual(
            lines,
            [
                ("Seat Cushion Cover", "cls-20"),
                ("(Housse de coussin de siege):", "cls-19"),
                ("Primary Fabric(Tissu Principal):", "cls-25"),
                ("100%Polyester(Polyester)", "cls-25"),
                ("Second Fabric:", "cls-25"),
                ("100%Polyester(Polyester)", "cls-25"),
                ("Base Cushion Cover", "cls-20"),
                ("(Housse de coussin de base):", "cls-19"),
                ("Custom Blend:", "cls-25"),
            ],
        )
        self.assertEqual(alerts, [])


if __name__ == "__main__":
    unittest.main()
