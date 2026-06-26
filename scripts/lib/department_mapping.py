from __future__ import annotations

from dataclasses import dataclass
import re
import shlex
from pathlib import Path

SPECIAL_DEPARTMENT_MATCHES: tuple[tuple[str, str], ...] = (
    ("citrix", "e-Government"),
    ("domain controllers", "Domain Controllers"),
)


def normalize(value: str | None) -> str:
    if not value:
        return ""
    normalized = value.lower()
    normalized = re.sub(r"[^a-z0-9]+", " ", normalized)
    normalized = re.sub(r"\s+", " ", normalized)
    return normalized.strip()


def clean_config_value(value: str) -> str:
    cleaned = value.strip()
    if len(cleaned) >= 2 and cleaned[0] == cleaned[-1] and cleaned[0] in {'"', "'"}:
        return cleaned[1:-1].strip()
    return cleaned


def contains_normalized_phrase(haystack: str, needle: str) -> bool:
    if not haystack or not needle:
        return False
    return f" {needle} " in f" {haystack} "


def is_identity_field(field_name: str) -> bool:
    return field_name.lower() in {"name", "cn", "dnshostname"}


def starts_with_code_token(normalized_value: str, code: str) -> bool:
    if not normalized_value or not code or " " in code:
        return False
    return any(token.startswith(code) for token in normalized_value.split())


def load_departments(path: Path) -> list[str]:
    departments: list[str] = []
    with path.open("r", encoding="utf-8") as handle:
        for line in handle:
            name = clean_config_value(line)
            if name and not name.startswith("#"):
                departments.append(name)
    return departments


def load_dept_codes(path: Path) -> dict[str, str]:
    mapping: dict[str, str] = {}
    if not path.exists():
        return mapping

    with path.open("r", encoding="utf-8") as handle:
        for line in handle:
            line = line.strip()
            if not line or line.startswith("#"):
                continue

            if ":" in line:
                key, value = line.split(":", 1)
            else:
                parts = shlex.split(line)
                key, value = parts[0], " ".join(parts[1:]) if len(parts) > 1 else ""

            key = normalize(clean_config_value(key))
            value = clean_config_value(value)
            if key and value:
                mapping[key] = value

    return mapping


def build_dept_info(departments: list[str], min_token_len: int = 3) -> list[dict[str, object]]:
    info: list[dict[str, object]] = []
    for department in departments:
        norm = normalize(department)
        tokens = [norm] if norm and " " not in norm and len(norm) >= min_token_len else []
        info.append({"name": department, "norm": norm, "tokens": tokens})
    return info


@dataclass(frozen=True)
class DepartmentMatch:
    department: str
    matched_by: str


class DepartmentMatcher:
    def __init__(
        self,
        departments: list[str],
        code_map: dict[str, str] | None = None,
        special_matches: tuple[tuple[str, str], ...] = SPECIAL_DEPARTMENT_MATCHES,
    ) -> None:
        self.departments = departments
        self.code_map = code_map or {}
        self.special_matches = special_matches
        self.dept_info = build_dept_info(departments)

    def match_values(self, values: list[tuple[str, str | None]]) -> DepartmentMatch | None:
        normalized_parts: list[tuple[str, str]] = []
        for field_name, raw_value in values:
            if raw_value:
                normalized_value = normalize(raw_value)
                if normalized_value:
                    normalized_parts.append((field_name, normalized_value))

        if not normalized_parts:
            return None

        combined = " ".join(value for _, value in normalized_parts)

        for match_text, department_name in self.special_matches:
            normalized_match = normalize(match_text)
            if contains_normalized_phrase(combined, normalized_match):
                return DepartmentMatch(department=department_name, matched_by=f"special:{normalized_match}")

        for key, value in sorted(self.code_map.items(), key=lambda item: len(item[0]), reverse=True):
            if contains_normalized_phrase(combined, key):
                return DepartmentMatch(department=value, matched_by=f"code:{key}")

            for field_name, normalized_value in normalized_parts:
                if is_identity_field(field_name) and starts_with_code_token(normalized_value, key):
                    return DepartmentMatch(department=value, matched_by=f"code-prefix:{key}")

        for department in self.dept_info:
            department_name = str(department["name"])
            department_norm = str(department["norm"])
            if contains_normalized_phrase(combined, department_norm):
                return DepartmentMatch(department=department_name, matched_by=f"name:{department_norm}")

        for department in self.dept_info:
            department_name = str(department["name"])
            department_tokens = list(department["tokens"])
            for token in department_tokens:
                if contains_normalized_phrase(combined, token):
                    return DepartmentMatch(department=department_name, matched_by=f"token:{token}")

        return None

