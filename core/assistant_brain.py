from __future__ import annotations

from dataclasses import dataclass, field
import json
import os
import re
from typing import Any
from urllib.request import Request, urlopen

from config import DEFAULT_MODEL

__all__ = ["SearchDirection", "AssistantBrain"]


@dataclass(slots=True)
class SearchDirection:
    role: str = ""
    location: str = ""
    include_keywords: list[str] = field(default_factory=list)
    exclude_keywords: list[str] = field(default_factory=list)
    target_sites: list[str] = field(default_factory=list)
    avoid_sites: list[str] = field(default_factory=list)
    remote_preference: str = ""
    seniority: str = ""

    def summary(self) -> str:
        parts = [
            f"Role: {self.role or 'Not set'}",
            f"Location: {self.location or 'Not set'}",
            f"Include: {', '.join(self.include_keywords) if self.include_keywords else 'None'}",
            f"Exclude: {', '.join(self.exclude_keywords) if self.exclude_keywords else 'None'}",
            f"Target sites: {', '.join(self.target_sites) if self.target_sites else 'Any trusted site'}",
            f"Avoid sites: {', '.join(self.avoid_sites) if self.avoid_sites else 'None'}",
            f"Remote: {self.remote_preference or 'No preference'}",
            f"Seniority: {self.seniority or 'Any'}",
        ]
        return "\n".join(parts)

    def to_preferences(self) -> dict[str, str]:
        return {
            "role": self.role,
            "location": self.location,
            "include_keywords": " ".join(self.include_keywords),
            "exclude_keywords": " ".join(self.exclude_keywords),
            "target_sites": ",".join(self.target_sites),
            "avoid_sites": ",".join(self.avoid_sites),
            "remote_preference": self.remote_preference,
            "seniority": self.seniority,
        }


class AssistantBrain:
    def __init__(self) -> None:
        self.openai_api_key = os.getenv("OPENAI_API_KEY", "").strip()
        self.gemini_api_key = os.getenv("GEMINI_API_KEY", "").strip()

    def provider_label(self) -> str:
        if self.openai_api_key:
            return "OpenAI API"
        if self.gemini_api_key:
            return "Gemini API"
        return "Local planner"

    def update_direction(self, message: str, current: SearchDirection) -> tuple[SearchDirection, list[str]]:
        remote = self._try_model_parse(message, current)
        if remote is not None:
            return remote
        return self._local_parse(message, current)

    def next_step_guidance(self, direction: SearchDirection) -> str:
        if not direction.role:
            return "Step 1: tell me the role you want, for example 'Find QA tester jobs'."
        if not direction.location:
            return "Step 2: tell me the location, for example 'in Auckland' or 'remote in New Zealand'."
        if not direction.include_keywords and not direction.target_sites:
            return "Step 3: refine the search with keywords or target sites, for example 'focus on automation and API testing on LinkedIn and Greenhouse'."
        return "Step 4: ask me to search, review matches, then choose one to start the application."

    def _local_parse(self, message: str, current: SearchDirection) -> tuple[SearchDirection, list[str]]:
        direction = SearchDirection(
            role=current.role,
            location=current.location,
            include_keywords=list(current.include_keywords),
            exclude_keywords=list(current.exclude_keywords),
            target_sites=list(current.target_sites),
            avoid_sites=list(current.avoid_sites),
            remote_preference=current.remote_preference,
            seniority=current.seniority,
        )
        notes: list[str] = []
        text = " ".join(message.replace("\n", " ").split())
        lowered = text.lower()

        role_patterns = [
            r"(?:find|search for|look for)\s+(.+?)\s+jobs?(?:\s+in|\s+at|\s*$)",
            r"(?:role|job title)\s*(?:is|=|:)?\s*(.+)",
        ]
        for pattern in role_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                role = match.group(1).strip(" ,.")
                if role and len(role) <= 80:
                    direction.role = role
                    notes.append(f"role set to {role}")
                    break

        location_match = re.search(r"(?:in|location)\s+(.*?)(?:\s+with|\s+focus|\s+avoid|\s+exclude|$)", text, re.IGNORECASE)
        if location_match:
            location = location_match.group(1).strip(" ,.")
            if location and len(location) <= 80:
                direction.location = location
                notes.append(f"location set to {location}")

        include_match = re.search(r"(?:focus on|include|with keywords?)\s+(.+?)(?:\s+avoid|\s+exclude|\s+on\s+|$)", text, re.IGNORECASE)
        if include_match:
            direction.include_keywords = self._split_terms(include_match.group(1))
            if direction.include_keywords:
                notes.append(f"include keywords set to {', '.join(direction.include_keywords)}")

        exclude_match = re.search(r"(?:avoid|exclude)\s+(.+?)(?:\s+on\s+|$)", text, re.IGNORECASE)
        if exclude_match:
            direction.exclude_keywords = self._split_terms(exclude_match.group(1))
            if direction.exclude_keywords:
                notes.append(f"exclude keywords set to {', '.join(direction.exclude_keywords)}")

        site_match = re.search(r"(?:on|target sites?)\s+(.+)", text, re.IGNORECASE)
        if site_match:
            sites = self._split_terms(site_match.group(1))
            normalized_sites = [self._normalize_site(site) for site in sites]
            normalized_sites = [site for site in normalized_sites if site]
            if normalized_sites:
                direction.target_sites = normalized_sites
                notes.append(f"target sites set to {', '.join(normalized_sites)}")

        if "remote" in lowered:
            direction.remote_preference = "remote"
            notes.append("remote preference set to remote")
        if "hybrid" in lowered:
            direction.remote_preference = "hybrid"
            notes.append("remote preference set to hybrid")
        if "onsite" in lowered or "on-site" in lowered:
            direction.remote_preference = "onsite"
            notes.append("remote preference set to onsite")

        for label in ("intern", "junior", "mid", "senior", "lead", "manager"):
            if label in lowered:
                direction.seniority = label
                notes.append(f"seniority set to {label}")
                break

        return direction, notes

    def _try_model_parse(self, message: str, current: SearchDirection) -> tuple[SearchDirection, list[str]] | None:
        prompt = (
            "Extract or update a structured job-search direction from the user message.\n"
            "Return strict JSON with keys: role, location, include_keywords, exclude_keywords, "
            "target_sites, avoid_sites, remote_preference, seniority, notes.\n"
            "Use arrays for keyword/site fields. Preserve unspecified current values by repeating them.\n\n"
            f"Current direction:\n{current.summary()}\n\n"
            f"User message:\n{message}"
        )
        try:
            payload = self._call_openai(prompt) if self.openai_api_key else self._call_gemini(prompt) if self.gemini_api_key else None
            if not payload:
                return None
            parsed = json.loads(payload)
            direction = SearchDirection(
                role=str(parsed.get("role", current.role)).strip(),
                location=str(parsed.get("location", current.location)).strip(),
                include_keywords=self._clean_list(parsed.get("include_keywords", current.include_keywords)),
                exclude_keywords=self._clean_list(parsed.get("exclude_keywords", current.exclude_keywords)),
                target_sites=self._clean_list(parsed.get("target_sites", current.target_sites)),
                avoid_sites=self._clean_list(parsed.get("avoid_sites", current.avoid_sites)),
                remote_preference=str(parsed.get("remote_preference", current.remote_preference)).strip(),
                seniority=str(parsed.get("seniority", current.seniority)).strip(),
            )
            notes = self._clean_list(parsed.get("notes", []))
            return direction, notes
        except Exception:
            return None

    def _call_openai(self, prompt: str) -> str | None:
        request = Request(
            "https://api.openai.com/v1/responses",
            headers={
                "Authorization": f"Bearer {self.openai_api_key}",
                "Content-Type": "application/json",
            },
            data=json.dumps(
                {
                    "model": DEFAULT_MODEL,
                    "input": prompt,
                    "temperature": 0.2,
                }
            ).encode("utf-8"),
            method="POST",
        )
        with urlopen(request, timeout=20) as response:
            payload = json.loads(response.read().decode("utf-8"))
        if isinstance(payload.get("output_text"), str):
            return payload["output_text"]
        output = payload.get("output", [])
        texts: list[str] = []
        for item in output:
            if str(item.get("type", "")).lower() != "message":
                continue
            for content in item.get("content", []):
                text = str(content.get("text", "")).strip()
                if text:
                    texts.append(text)
        return "\n".join(texts) if texts else None

    def _call_gemini(self, prompt: str) -> str | None:
        request = Request(
            "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent",
            headers={
                "x-goog-api-key": self.gemini_api_key,
                "Content-Type": "application/json",
            },
            data=json.dumps(
                {
                    "contents": [{"parts": [{"text": prompt}]}],
                    "generationConfig": {"temperature": 0.2},
                }
            ).encode("utf-8"),
            method="POST",
        )
        with urlopen(request, timeout=20) as response:
            payload = json.loads(response.read().decode("utf-8"))
        candidates = payload.get("candidates", [])
        if not candidates:
            return None
        parts = candidates[0].get("content", {}).get("parts", [])
        texts = [str(part.get("text", "")) for part in parts if str(part.get("text", "")).strip()]
        return "\n".join(texts) if texts else None

    @staticmethod
    def _split_terms(raw: str) -> list[str]:
        return [part.strip(" ,.") for part in re.split(r",| and ", raw) if part.strip(" ,.")]

    @staticmethod
    def _normalize_site(site: str) -> str:
        lowered = site.lower().strip()
        mapping = {
            "linkedin": "linkedin.com",
            "indeed": "indeed.com",
            "greenhouse": "greenhouse.io",
            "lever": "lever.co",
            "workday": "workday",
            "seek": "seek",
        }
        return mapping.get(lowered, lowered)

    @staticmethod
    def _clean_list(value: Any) -> list[str]:
        if isinstance(value, list):
            return [str(item).strip() for item in value if str(item).strip()]
        return []
