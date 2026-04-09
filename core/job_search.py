from __future__ import annotations

from dataclasses import dataclass
from html import unescape
import re
from typing import Any
from urllib.parse import parse_qs, quote_plus, urlparse
from urllib.request import Request, urlopen


@dataclass(slots=True)
class JobMatch:
    title: str
    url: str
    snippet: str
    score: float


class JobSearchEngine:
    """Searches public job pages and ranks likely matches."""

    TRUSTED_DOMAINS = ("linkedin.com", "indeed.com", "greenhouse.io", "lever.co", "workday", "jobs")
    BLOCKED_URL_TERMS = (
        "/search",
        "google.com",
        "duckduckgo.com",
        "bing.com",
        "yahoo.com",
        "seek.com",
        "jora.com",
        "jooble.org",
        "jobs?q=",
        "/jobs?",
    )
    JOB_PATH_TERMS = ("job", "jobs", "careers", "career", "apply", "position", "vacancy", "opening")

    def search(
        self,
        preferences: dict[str, str],
        max_results: int = 8,
    ) -> list[JobMatch]:
        query = self._build_query(preferences)
        matches = self._search_duckduckgo(query, max_results=max_results)
        if matches:
            return matches
        return self._fallback_links(preferences, max_results=max_results)

    def pick_best(self, matches: list[JobMatch]) -> JobMatch | None:
        if not matches:
            return None
        return sorted(matches, key=lambda item: item.score, reverse=True)[0]

    def pick_smart_matches(self, matches: list[JobMatch], max_results: int = 5) -> list[JobMatch]:
        ranked = sorted(matches, key=lambda item: item.score, reverse=True)
        filtered = [item for item in ranked if self._is_likely_job_posting(item.url, item.title, item.snippet)]
        return (filtered or ranked)[:max_results]

    @staticmethod
    def _build_query(preferences: dict[str, str]) -> str:
        role = (preferences.get("role") or "").strip()
        location = (preferences.get("location") or "").strip()
        include_keywords = (preferences.get("include_keywords") or "").strip()
        exclude_keywords = (preferences.get("exclude_keywords") or "").strip()
        remote_preference = (preferences.get("remote_preference") or "").strip()
        seniority = (preferences.get("seniority") or "").strip()
        target_sites = [part.strip() for part in str(preferences.get("target_sites") or "").split(",") if part.strip()]
        avoid_sites = [part.strip() for part in str(preferences.get("avoid_sites") or "").split(",") if part.strip()]

        query_parts = [role, seniority, location, remote_preference, include_keywords, "jobs", "apply"]
        query_parts.extend([f"site:{site}" for site in target_sites])
        query_parts.extend([f"-site:{site}" for site in avoid_sites])
        if exclude_keywords:
            query_parts.extend(f"-{term}" for term in exclude_keywords.split())
        return " ".join(part for part in query_parts if part)

    def _search_duckduckgo(self, query: str, max_results: int) -> list[JobMatch]:
        url = f"https://duckduckgo.com/html/?q={quote_plus(query)}"
        req = Request(url, headers={"User-Agent": "Mozilla/5.0"})

        try:
            with urlopen(req, timeout=10) as response:
                body = response.read().decode("utf-8", errors="ignore")
        except Exception:
            return []

        return self._parse_results(html=body, query=query, max_results=max_results)

    def _parse_results(self, html: str, query: str, max_results: int) -> list[JobMatch]:
        anchor_pattern = re.compile(
            r'<a[^>]*class="[^"]*result__a[^"]*"[^>]*href="([^"]+)"[^>]*>(.*?)</a>',
            re.IGNORECASE | re.DOTALL,
        )
        snippet_pattern = re.compile(
            r'<a[^>]*class="[^"]*result__a[^"]*"[^>]*>.*?</a>.*?<a[^>]*class="[^"]*result__snippet[^"]*"[^>]*>(.*?)</a>',
            re.IGNORECASE | re.DOTALL,
        )

        anchors = anchor_pattern.findall(html)
        snippets = snippet_pattern.findall(html)
        query_terms = [term for term in re.split(r"\s+", query.lower()) if term]

        results: list[JobMatch] = []
        for index, (href, title_html) in enumerate(anchors[: max_results * 3]):
            title = self._strip_html(title_html)
            final_url = self._extract_redirect_target(href)
            snippet = self._strip_html(snippets[index]) if index < len(snippets) else ""
            score = self._score(title=title, url=final_url, snippet=snippet, query_terms=query_terms)
            if final_url.startswith("http"):
                results.append(JobMatch(title=title, url=final_url, snippet=snippet, score=score))

        results.sort(key=lambda item: item.score, reverse=True)
        deduped: list[JobMatch] = []
        seen_urls: set[str] = set()
        for item in results:
            normalized = item.url.lower().rstrip("/")
            if normalized in seen_urls:
                continue
            seen_urls.add(normalized)
            deduped.append(item)
        return deduped[:max_results]

    def _score(self, title: str, url: str, snippet: str, query_terms: list[str]) -> float:
        haystack = f"{title} {snippet} {url}".lower()
        score = 0.0
        for term in query_terms:
            if term and term in haystack:
                score += 1.0
        if any(domain in url.lower() for domain in self.TRUSTED_DOMAINS):
            score += 2.0
        if "job" in haystack or "career" in haystack:
            score += 0.5
        if self._is_likely_job_posting(url, title, snippet):
            score += 2.5
        else:
            score -= 1.5
        return score

    def _is_likely_job_posting(self, url: str, title: str, snippet: str) -> bool:
        lowered_url = url.lower()
        lowered_text = f"{title} {snippet}".lower()
        if any(term in lowered_url for term in self.BLOCKED_URL_TERMS):
            return False
        if any(term in lowered_url for term in self.JOB_PATH_TERMS):
            return True
        return any(term in lowered_text for term in ("apply", "job", "career", "position", "opening"))

    @staticmethod
    def _extract_redirect_target(href: str) -> str:
        if "duckduckgo.com/l/?" not in href:
            return href
        parsed = urlparse(href)
        params = parse_qs(parsed.query)
        target = params.get("uddg", [""])[0]
        return unescape(target) or href

    def _fallback_links(self, preferences: dict[str, str], max_results: int) -> list[JobMatch]:
        role = quote_plus((preferences.get("role") or "software engineer").strip())
        location = quote_plus((preferences.get("location") or "").strip())
        query = f"{role}%20{location}".strip("%20")

        candidates = [
            ("LinkedIn Jobs", f"https://www.linkedin.com/jobs/search/?keywords={query}", 3.0),
            ("Indeed Jobs", f"https://www.indeed.com/jobs?q={query}", 2.8),
            ("Google Jobs", f"https://www.google.com/search?q={query}+jobs", 2.5),
            ("RemoteOK Jobs", f"https://remoteok.com/remote-{role}-jobs", 2.3),
        ]
        results = [JobMatch(title=t, url=u, snippet="Generated fallback search URL.", score=s) for t, u, s in candidates]
        return results[:max_results]

    @staticmethod
    def _strip_html(text: str) -> str:
        no_tags = re.sub(r"<[^>]+>", " ", text)
        return " ".join(unescape(no_tags).split())
