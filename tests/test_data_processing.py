import json
import sys
import tempfile
import unittest
from pathlib import Path

SOURCE = Path(__file__).resolve().parents[1] / "source"
sys.path.insert(0, str(SOURCE))

from data_processing import safe_float_conversion, safe_int_conversion  # noqa: E402
from export_data import flatten_dict, save_to_csv, save_to_json, save_to_xml  # noqa: E402


class DataProcessingTests(unittest.TestCase):
    def test_safe_float_conversion_accepts_grouping_and_defaults_invalid_values(self):
        self.assertEqual(1234.5, safe_float_conversion("1,234.5"))
        self.assertEqual(0.0, safe_float_conversion("not-a-number"))

    def test_safe_int_conversion_handles_supported_types(self):
        self.assertEqual(42, safe_int_conversion("42"))
        self.assertEqual(0, safe_int_conversion("   "))
        self.assertEqual(3, safe_int_conversion(3.9))
        self.assertEqual(1, safe_int_conversion(True))

    def test_flatten_dict_preserves_nested_list_positions(self):
        self.assertEqual(
            {"section:value": 2, "rows:0:name": "alpha"},
            flatten_dict({"section": {"value": 2}, "rows": [{"name": "alpha"}]}),
        )

    def test_structured_exports_write_parseable_outputs(self):
        data = [{"Identification": {"Employee": "Example"}, "Earnings": {"Regular": {"Amount": 10.0}}}]
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            json_path = root / "output.json"
            csv_path = root / "output.csv"
            xml_path = root / "output.xml"
            self.assertTrue(save_to_json(data, json_path))
            self.assertTrue(save_to_csv(data, csv_path))
            self.assertTrue(save_to_xml(data, xml_path))
            self.assertEqual(data, json.loads(json_path.read_text(encoding="utf-8")))
            self.assertIn("Identification:Employee", csv_path.read_text(encoding="utf-8"))
            self.assertIn("<root>", xml_path.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
