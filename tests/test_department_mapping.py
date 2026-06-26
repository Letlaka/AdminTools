from __future__ import annotations

import sys
from pathlib import Path
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from lib.department_mapping import DepartmentMatcher, load_dept_codes, load_departments


class DepartmentMappingTests(unittest.TestCase):
    def test_department_files_load(self) -> None:
        root = Path(__file__).resolve().parents[1]
        departments = load_departments(root / "config" / "dept_list.txt")
        code_map = load_dept_codes(root / "config" / "dept_codes.txt")
        self.assertIn("Education", departments)
        self.assertIn("Department of Health", departments)
        self.assertNotIn('"Department of Health"', departments)
        self.assertEqual(code_map["gdh"], "Department of Health")

    def test_match_uses_code_map(self) -> None:
        matcher = DepartmentMatcher(
            departments=["Health", "Education"],
            code_map={"gdh": "Health"},
        )
        match = matcher.match_values(
            [
                ("OUPath", "OU=Regional,OU=GDH,DC=example,DC=local"),
                ("DNSHostName", "pc01.example.local"),
            ]
        )
        self.assertIsNotNone(match)
        assert match is not None
        self.assertEqual(match.department, "Health")

    def test_match_uses_code_prefix_from_device_name(self) -> None:
        matcher = DepartmentMatcher(
            departments=["Department of Health", "Education"],
            code_map={"gdh": "Department of Health", "gde": "Education"},
        )

        health_match = matcher.match_values(
            [
                ("Name", "GDHJDH_STR02"),
                ("DNSHostName", "GDHJDH_STR02.example.corp.local"),
            ]
        )
        education_match = matcher.match_values(
            [
                ("Name", "GDEGE7TH-M1282"),
                ("DNSHostName", "GDEGE7TH-M1282.example.corp.local"),
            ]
        )

        self.assertIsNotNone(health_match)
        self.assertIsNotNone(education_match)
        assert health_match is not None
        assert education_match is not None
        self.assertEqual(health_match.department, "Department of Health")
        self.assertEqual(health_match.matched_by, "code-prefix:gdh")
        self.assertEqual(education_match.department, "Education")
        self.assertEqual(education_match.matched_by, "code-prefix:gde")

    def test_match_uses_citrix_special_case(self) -> None:
        matcher = DepartmentMatcher(
            departments=["e-Government", "Health"],
            code_map={},
        )
        match = matcher.match_values(
            [
                ("Description", "Citrix virtual desktop host"),
                ("OUPath", "OU=Platforms,DC=example,DC=local"),
            ]
        )
        self.assertIsNotNone(match)
        assert match is not None
        self.assertEqual(match.department, "e-Government")

    def test_full_department_phrase_beats_generic_earlier_token(self) -> None:
        matcher = DepartmentMatcher(
            departments=["Office of the Premier", "Department of Health", "Community Safety"],
            code_map={},
        )

        match = matcher.match_values(
            [
                (
                    "OUPath",
                    "example.corp.local/Department Of Health/Facilities/"
                    "Johannesburg F Health sub-District/Head Office/computers",
                ),
                ("DNSHostName", "health45sapdist.example.corp.local"),
            ]
        )

        self.assertIsNotNone(match)
        assert match is not None
        self.assertEqual(match.department, "Department of Health")
        self.assertEqual(match.matched_by, "name:department of health")

    def test_quoted_multi_word_department_is_not_split_into_word_tokens(self) -> None:
        matcher = DepartmentMatcher(
            departments=["Office of the Premier", "Community Safety"],
            code_map={},
        )

        match = matcher.match_values(
            [
                (
                    "OUPath",
                    "example.corp.local/Department Of Community Safety/"
                    "Workstations/Vereeniging office",
                ),
                ("DNSHostName", "dcs-ver-kiosk9.example.corp.local"),
            ]
        )

        self.assertIsNotNone(match)
        assert match is not None
        self.assertEqual(match.department, "Community Safety")
        self.assertEqual(match.matched_by, "name:community safety")


if __name__ == "__main__":
    unittest.main()

