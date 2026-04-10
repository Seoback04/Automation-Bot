from __future__ import annotations

import json
from pathlib import Path
from typing import Any


class ProfileStore:
    """Loads, normalizes, and persists applicant profile data."""

    BASIC_KEY_MAP = {
        "full_name": "name",
        "first_name": "first_name",
        "last_name": "last_name",
        "email": "email",
        "phone": "phone",
        "location": "location",
        "city": "city",
        "country": "country",
        "linkedin": "linkedin",
        "github": "github",
        "website": "website",
        "resume_url": "resume_url",
        "resume_path": "resume_path",
        "summary": "summary",
    }

    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)

    def load(self) -> dict[str, Any]:
        if not self.path.exists():
            profile = self._default_profile()
            self.save(profile)
            return profile

        with self.path.open("r", encoding="utf-8") as file:
            profile = json.load(file)
        self._ensure_structure(profile)
        return profile

    def save(self, profile: dict[str, Any]) -> None:
        self._ensure_structure(profile)
        with self.path.open("w", encoding="utf-8") as file:
            json.dump(profile, file, indent=2)

    def get_learned_answers(self, profile: dict[str, Any]) -> dict[str, str]:
        self._ensure_structure(profile)
        learned = profile["memory"]["learned_answers"]
        return {str(k): str(v) for k, v in learned.items()}

    def get_custom_answers(self, profile: dict[str, Any]) -> dict[str, str]:
        self._ensure_structure(profile)
        custom = profile["memory"]["custom_fields"]
        return {str(k): str(v) for k, v in custom.items()}

    def lookup_answer(self, profile: dict[str, Any], field: dict[str, Any]) -> str:
        self._ensure_structure(profile)
        memory = profile["memory"]

        candidates = self._field_alias_candidates(field)
        learned = memory["learned_answers"]
        for key in candidates:
            value = str(learned.get(key, "")).strip()
            if value:
                return value

        custom = memory["custom_fields"]
        labels = memory["custom_field_labels"]
        aliases = memory["field_aliases"]
        normalized_custom = {self._normalize_key(key): str(value).strip() for key, value in custom.items()}
        for key in candidates:
            normalized = self._normalize_key(key)
            aliased_key = str(aliases.get(normalized, "")).strip()
            if aliased_key:
                aliased_value = str(learned.get(aliased_key, "")).strip()
                if aliased_value:
                    return aliased_value
            if normalized in normalized_custom and normalized_custom[normalized]:
                return normalized_custom[normalized]

        for custom_key, label in labels.items():
            normalized_label = self._normalize_key(label)
            if normalized_label and normalized_label in {self._normalize_key(item) for item in candidates}:
                value = str(custom.get(custom_key, "")).strip()
                if value:
                    return value

        documents = memory["documents"]
        normalized_candidates = {self._normalize_key(item) for item in candidates}
        for doc_key, meta in documents.items():
            doc_path = str(meta.get("path", "")).strip()
            if not doc_path:
                continue
            doc_label = self._normalize_key(str(meta.get("label", "")))
            if self._normalize_key(doc_key) in normalized_candidates or (doc_label and doc_label in normalized_candidates):
                return doc_path

        return ""

    def remember_answer(
        self,
        profile: dict[str, Any],
        field_key: str,
        value: str,
        label: str,
    ) -> None:
        clean_value = str(value).strip()
        if not clean_value:
            return

        self._ensure_structure(profile)
        profile["memory"]["learned_answers"][field_key] = clean_value

        basic_key = self.BASIC_KEY_MAP.get(field_key)
        if basic_key:
            profile["basics"][basic_key] = clean_value
            return

        if field_key.startswith("custom:"):
            custom_key = field_key.split("custom:", 1)[1]
            profile["memory"]["custom_fields"][custom_key] = clean_value
            if label.strip():
                profile["memory"]["custom_field_labels"][custom_key] = label.strip()

    def remember_field_answer(
        self,
        profile: dict[str, Any],
        field: dict[str, Any],
        value: str,
    ) -> None:
        clean_value = str(value).strip()
        if not clean_value:
            return

        self._ensure_structure(profile)
        memory = profile["memory"]
        aliases = self._field_alias_candidates(field)
        if not aliases:
            aliases = [str(field.get("key", "")).strip()]

        primary_key = aliases[0]
        label = str(field.get("label") or primary_key).strip()
        self.remember_answer(profile, primary_key, clean_value, label)

        for alias in aliases[1:]:
            alias = str(alias).strip()
            if alias:
                memory["learned_answers"][alias] = clean_value

        for alias in aliases:
            normalized = self._normalize_key(alias)
            if normalized:
                memory["field_aliases"][normalized] = primary_key

    def remember_document(
        self,
        profile: dict[str, Any],
        category: str,
        path: str,
        label: str = "",
    ) -> None:
        clean_path = str(path).strip()
        clean_category = str(category).strip().lower()
        if not clean_path or not clean_category:
            return

        self._ensure_structure(profile)
        documents = profile["memory"]["documents"]
        documents[clean_category] = {
            "path": clean_path,
            "label": label.strip() or clean_category.replace("_", " ").title(),
        }

    def record_application_step(
        self,
        profile: dict[str, Any],
        step_data: dict[str, Any],
    ) -> None:
        self._ensure_structure(profile)
        history = profile["memory"]["application_history"]
        history.append(step_data)
        if len(history) > 50:
            del history[:-50]

    @staticmethod
    def _ensure_structure(profile: dict[str, Any]) -> None:
        if not isinstance(profile.get("basics"), dict):
            profile["basics"] = {}
        basics = profile["basics"]
        basics.setdefault("name", "")
        basics.setdefault("first_name", "")
        basics.setdefault("last_name", "")
        basics.setdefault("email", "")
        basics.setdefault("phone", "")
        basics.setdefault("location", "")
        basics.setdefault("city", "")
        basics.setdefault("country", "")
        basics.setdefault("linkedin", "")
        basics.setdefault("github", "")
        basics.setdefault("website", "")
        basics.setdefault("resume_url", "")
        basics.setdefault("resume_path", "")
        basics.setdefault("summary", "")
        if not isinstance(profile.get("experience"), list):
            profile["experience"] = []
        if not isinstance(profile.get("skills"), list):
            profile["skills"] = []
        if not isinstance(profile.get("preferences"), dict):
            profile["preferences"] = {}
        preferences = profile["preferences"]
        preferences.setdefault("work_authorized", "")
        preferences.setdefault("requires_sponsorship", "")
        preferences.setdefault("salary_expectation", "")
        preferences.setdefault("notice_period", "")
        if not isinstance(profile.get("memory"), dict):
            profile["memory"] = {}
        if not isinstance(profile.get("job_preferences"), dict):
            profile["job_preferences"] = {}
        memory = profile["memory"]
        if not isinstance(memory.get("learned_answers"), dict):
            memory["learned_answers"] = {}
        if not isinstance(memory.get("custom_fields"), dict):
            memory["custom_fields"] = {}
        if not isinstance(memory.get("custom_field_labels"), dict):
            memory["custom_field_labels"] = {}
        if not isinstance(memory.get("field_aliases"), dict):
            memory["field_aliases"] = {}
        if not isinstance(memory.get("documents"), dict):
            memory["documents"] = {}
        if not isinstance(memory.get("application_history"), list):
            memory["application_history"] = []

    @staticmethod
    def _default_profile() -> dict[str, Any]:
        return {
            "basics": {
                "name": "",
                "first_name": "",
                "last_name": "",
                "email": "",
                "phone": "",
                "location": "",
                "city": "",
                "country": "",
                "linkedin": "",
                "github": "",
                "website": "",
                "resume_url": "",
                "resume_path": "",
                "summary": "",
            },
            "experience": [],
            "skills": [],
            "preferences": {
                "work_authorized": "",
                "requires_sponsorship": "",
                "salary_expectation": "",
                "notice_period": "",
            },
            "memory": {
                "learned_answers": {},
                "custom_fields": {},
                "custom_field_labels": {},
                "field_aliases": {},
                "documents": {},
                "application_history": [],
            },
            "job_preferences": {
                "role": "",
                "location": "",
            },
        }

    @classmethod
    def _field_alias_candidates(cls, field: dict[str, Any]) -> list[str]:
        raw_values = [
            str(field.get("key", "")).strip(),
            str(field.get("name", "")).strip(),
            str(field.get("id", "")).strip(),
            str(field.get("label", "")).strip(),
            str(field.get("placeholder", "")).strip(),
            str(field.get("aria_label", "")).strip(),
            str(field.get("section", "")).strip(),
        ]
        candidates: list[str] = []
        seen: set[str] = set()
        for value in raw_values:
            if not value:
                continue
            normalized = cls._normalize_key(value)
            for candidate in (value, normalized):
                if candidate and candidate not in seen:
                    seen.add(candidate)
                    candidates.append(candidate)
        return candidates

    @staticmethod
    def _normalize_key(value: str) -> str:
        cleaned = "".join(ch.lower() if ch.isalnum() else "_" for ch in str(value))
        cleaned = "_".join(part for part in cleaned.split("_") if part)
        return cleaned.strip("_")
