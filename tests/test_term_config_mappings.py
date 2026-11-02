import unittest

from term_config import find_label4_title_match, find_material_match, find_part_match
from validation import parse_material_text


class TermConfigMappingTests(unittest.TestCase):
    def test_duck_feather_is_a_shared_material(self) -> None:
        parts, alerts, unmapped = parse_material_text("Body\n100% DUCK FEATHER")
        self.assertFalse(unmapped)
        self.assertEqual(alerts, [])
        self.assertEqual(parts[0]["materials"][0]["text"], "100% DUCK FEATHER(plume de canard)")
        self.assertEqual(find_label4_title_match("DUCK FEATHER"), ("Duck Feather", "plume de canard", "Material"))

    def test_new_material_terms_match_requested_translations(self) -> None:
        self.assertEqual(
            find_material_match("VISCOELASTIC POLYURETHANE FOAM"),
            ("VISCOELASTIC POLYURETHANE FOAM", "Mousse de polyuréthane viscoélastique"),
        )
        self.assertEqual(
            find_material_match("POLYURETHANE FOAM"),
            ("POLYURETHANE FOAM", "Mousse de polyuréthane"),
        )
        self.assertEqual(
            find_material_match("INNERSPRING UNIT"),
            ("INNERSPRING UNIT", "Unité de ressorts"),
        )
        self.assertEqual(
            find_material_match("Lyocell"),
            ("LYOCELL", "Lyocell"),
        )
        self.assertEqual(
            find_material_match("Cotton Canvas"),
            ("COTTON CANVAS", "Toile de coton"),
        )

    def test_new_part_terms_match_requested_translations(self) -> None:
        self.assertEqual(
            find_part_match("Top fabric"),
            ("TOP FABRIC", "Tissu du dessus"),
        )
        self.assertEqual(
            find_part_match("Border fabric"),
            ("BORDER FABRIC", "Tissu de bordure"),
        )

    def test_parse_material_text_uses_new_shared_terms(self) -> None:
        parts, alerts, has_unmapped_terms = parse_material_text(
            "Top fabric\n"
            "60% Lyocell\n"
            "40% Cotton Canvas\n"
            "Border fabric\n"
            "100% VISCOELASTIC POLYURETHANE FOAM"
        )

        self.assertFalse(has_unmapped_terms)
        self.assertEqual(alerts, [])
        self.assertEqual(
            parts,
            [
                {
                    "title": "Top fabric(Tissu du dessus):",
                    "materials": [
                        {"text": "60% Lyocell(Lyocell)", "has_percentage": True},
                        {"text": "40% Cotton Canvas(Toile de coton)", "has_percentage": True},
                    ],
                },
                {
                    "title": "Border fabric(Tissu de bordure):",
                    "materials": [
                        {
                            "text": "100% VISCOELASTIC POLYURETHANE FOAM(Mousse de polyuréthane viscoélastique)",
                            "has_percentage": True,
                        }
                    ],
                },
            ],
        )


if __name__ == "__main__":
    unittest.main()
