from __future__ import annotations

import json
import time
from pathlib import Path
from typing import Any
import os
from urllib.error import URLError
from urllib.request import Request, urlopen

from selenium.webdriver.chrome.webdriver import WebDriver as Chrome
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.common.exceptions import SessionNotCreatedException
from webdriver_manager.chrome import ChromeDriverManager

from config import AUTOMATION_BROWSER_PROFILE_DIR, PAGE_SETTLE_SECONDS
from utils.selectors import (
    CONTINUE_TERMS,
    EASY_APPLY_TERMS,
    EXTERNAL_APPLY_TERMS,
    SUBMIT_TERMS,
)

__all__ = ["BrowserSession", "start_browser"]


class BrowserSession:
    """Manages a Selenium browser session for job application automation."""

    DEBUGGER_ADDRESS = "127.0.0.1:9222"
    JOB_SITE_HINTS = (
        "linkedin.com",
        "seek.",
        "indeed.com",
        "greenhouse.io",
        "lever.co",
        "workday",
        "smartrecruiters",
        "roberthalf.com",
        "jora.com",
        "jooble.org",
    )

    def __init__(self, headless: bool = False) -> None:
        self.headless = headless
        self.driver: Chrome | None = None
        self._brave_path = self._find_brave_path()
        self._chrome_driver_path = self._find_chromedriver_path()
        self._attached_to_existing = False

    # ------------------------------------------------------------------
    # Lifecycle
    # ------------------------------------------------------------------

    def start(self, attach_to_existing: bool = False) -> bool:
        use_profile = not self.headless
        self._attached_to_existing = False

        if attach_to_existing and not self.headless and self._debugger_available():
            try:
                self.driver = self._create_driver(use_profile=False, debugger_address=self.DEBUGGER_ADDRESS)
                self._attached_to_existing = True
                return True
            except Exception:
                self.driver = None

        try:
            self.driver = self._create_driver(use_profile=use_profile)
        except SessionNotCreatedException:
            self.driver = self._create_driver(use_profile=False)
        return False

    @property
    def attached_to_existing(self) -> bool:
        return self._attached_to_existing

    def stop(self, keep_browser_open: bool = False) -> None:
        if self.driver:
            if keep_browser_open or self._attached_to_existing:
                try:
                    service = getattr(self.driver, "service", None)
                    if service is not None:
                        service.stop()
                except Exception:
                    pass
                finally:
                    self.driver = None
                return
            try:
                self.driver.quit()
            except Exception:
                pass
            finally:
                self.driver = None

    # ------------------------------------------------------------------
    # Navigation
    # ------------------------------------------------------------------

    def open(self, url: str) -> None:
        driver = self._require_driver()
        self._switch_to_live_window(driver)
        driver.get(url)

    def open_in_new_tab(self, url: str) -> None:
        driver = self._require_driver()
        self._switch_to_live_window(driver)
        driver.switch_to.new_window("tab")
        driver.get(url)

    def prepare_for_job_search(self, url: str) -> None:
        driver = self._require_driver()
        self._switch_to_live_window(driver)
        if self._attached_to_existing:
            self.open_in_new_tab(url)
        else:
            self.open(url)

    def wait_for_page_settle(self, seconds: float | None = None) -> None:
        wait_time = seconds if seconds is not None else PAGE_SETTLE_SECONDS
        time.sleep(wait_time)
        driver = self._require_driver()
        try:
            WebDriverWait(driver, 5).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
        except Exception:
            pass

    def list_tabs(self, job_only: bool = False) -> list[dict[str, Any]]:
        driver = self._require_driver()
        tabs: list[dict[str, Any]] = []
        current_handle = driver.current_window_handle

        for handle in driver.window_handles:
            try:
                driver.switch_to.window(handle)
                title = driver.title or ""
                url = driver.current_url or ""
                is_job = self._looks_like_job_page(title=title, url=url)
                if job_only and not is_job:
                    continue
                tabs.append({
                    "handle": handle,
                    "title": title,
                    "url": url,
                    "is_job": is_job,
                })
            except Exception:
                continue

        driver.switch_to.window(current_handle)
        return tabs

    def switch_to_tab(self, handle: str) -> None:
        self._require_driver().switch_to.window(handle)

    def get_current_page_summary(self) -> dict[str, str]:
        driver = self._require_driver()
        return {
            "title": driver.title or "",
            "url": driver.current_url or "",
        }

    def has_application_form(self) -> bool:
        fields = self.collect_inputs()
        if len(fields) >= 2:
            return True
        return any(field.get("type") in {"file", "email", "tel"} or field.get("tag") == "textarea" for field in fields)

    def extract_role_suggestions(self, max_items: int = 12) -> list[str]:
        driver = self._require_driver()
        suggestions: list[str] = []
        seen: set[str] = set()

        def add_candidate(value: str) -> None:
            clean = " ".join(str(value).split()).strip(" -|")
            if len(clean) < 4 or len(clean) > 120:
                return
            lowered = clean.lower()
            if lowered in seen:
                return
            seen.add(lowered)
            suggestions.append(clean)

        try:
            title = driver.title or ""
            for part in title.replace("|", "-").split("-"):
                add_candidate(part)
        except Exception:
            pass

        selectors = [
            "h1",
            "h2",
            "h3",
            "[data-job-title]",
            "[class*='job'][class*='title']",
            "[class*='Job'][class*='Title']",
            "[class*='position']",
            "[class*='role']",
        ]
        for selector in selectors:
            try:
                for element in driver.find_elements(By.CSS_SELECTOR, selector)[:30]:
                    text = element.text.strip()
                    if self._looks_like_role_text(text):
                        add_candidate(text)
            except Exception:
                continue

        for field in self.collect_inputs():
            field_text = " ".join(
                str(field.get(part, "")).strip()
                for part in ("label", "name", "placeholder", "aria_label")
            ).strip()
            if self._looks_like_role_text(field_text):
                add_candidate(field_text)
            if str(field.get("tag", "")).lower() == "select":
                field_meta = field_text.lower()
                if any(term in field_meta for term in ("role", "position", "title")):
                    for option in field.get("options", [])[:20]:
                        if self._looks_like_role_text(str(option)):
                            add_candidate(str(option))

        return suggestions[:max_items]

    def find_best_matching_job_link(self, role: str) -> dict[str, str] | None:
        driver = self._require_driver()
        terms = [term.lower() for term in role.split() if len(term) >= 3]
        if not terms:
            return None

        best_score = 0.0
        best: dict[str, str] | None = None
        for element in driver.find_elements(By.TAG_NAME, "a")[:400]:
            try:
                text = " ".join(element.text.split()).strip()
                href = (element.get_attribute("href") or "").strip()
                if not text or not href.startswith("http"):
                    continue

                lowered = f"{text} {href}".lower()
                score = sum(2.0 for term in terms if term in text.lower())
                score += sum(1.0 for term in terms if term in href.lower())
                if self._looks_like_job_page(text, href):
                    score += 2.0
                if "apply" in lowered or "job" in lowered:
                    score += 1.0
                if score > best_score:
                    best_score = score
                    best = {"title": text, "url": href}
            except Exception:
                continue

        return best

    def open_best_matching_job(self, role: str) -> dict[str, str] | None:
        best = self.find_best_matching_job_link(role)
        if best is None:
            return None
        self.open(best["url"])
        self.wait_for_page_settle()
        return best

    # ------------------------------------------------------------------
    # Detection
    # ------------------------------------------------------------------

    def detect_apply_type(self) -> str:
        body_text = self._require_driver().find_element(By.TAG_NAME, "body").text.lower()
        for term in EASY_APPLY_TERMS:
            if term in body_text:
                return "easy_apply"
        return "external_apply"

    # ------------------------------------------------------------------
    # Clicking
    # ------------------------------------------------------------------

    def click_apply_entry(self, apply_type: str) -> bool:
        terms = EASY_APPLY_TERMS if apply_type == "easy_apply" else EXTERNAL_APPLY_TERMS
        return self._click_button_by_terms(terms)

    def click_next_step(self) -> bool:
        return self._click_button_by_terms(CONTINUE_TERMS)

    def click_submit(self) -> bool:
        return self._click_button_by_terms(SUBMIT_TERMS)

    # ------------------------------------------------------------------
    # Page content
    # ------------------------------------------------------------------

    def extract_page_text(self) -> str:
        try:
            return self._require_driver().find_element(By.TAG_NAME, "body").text
        except Exception:
            return ""

    def collect_inputs(self) -> list[dict[str, Any]]:
        driver = self._require_driver()
        fields: list[dict[str, Any]] = []
        seen_ids: set[str] = set()

        for tag in ("input", "textarea", "select"):
            for element in driver.find_elements(By.TAG_NAME, tag):
                try:
                    if not element.is_displayed():
                        continue
                except Exception:
                    continue

                elem_id = element.get_attribute("id") or ""
                name = element.get_attribute("name") or ""
                unique_key = f"{tag}:{elem_id}:{name}"
                if unique_key in seen_ids:
                    continue
                seen_ids.add(unique_key)

                field_type = (element.get_attribute("type") or tag).lower()
                if field_type in ("hidden", "submit", "button", "image"):
                    continue
                if not element.is_enabled():
                    continue
                if element.get_attribute("readonly") is not None:
                    continue

                label_text = self._find_label(element)
                options = self._get_options(element, tag) if tag == "select" else []
                section_text = self._find_section_context(element)
                required = (
                    element.get_attribute("required") is not None
                    or (element.get_attribute("aria-required") or "").lower() == "true"
                )

                fields.append({
                    "tag": tag,
                    "type": field_type,
                    "name": name,
                    "id": elem_id,
                    "label": label_text,
                    "placeholder": element.get_attribute("placeholder") or "",
                    "aria_label": element.get_attribute("aria-label") or "",
                    "section": section_text,
                    "required": required,
                    "options": options,
                    "selector": f"#{elem_id}" if elem_id else "",
                    "xpath": self._build_xpath(element, tag, name, elem_id),
                    "current_value": element.get_attribute("value") or "",
                })

        return fields

    # ------------------------------------------------------------------
    # Form filling
    # ------------------------------------------------------------------

    def apply_fill_plan(self, fill_plan: list[dict[str, Any]]) -> None:
        driver = self._require_driver()
        for item in fill_plan:
            value = str(item.get("value", "")).strip()
            if not value:
                continue

            element = self._locate_field(driver, item)
            if element is None:
                item["applied"] = False
                continue

            try:
                tag = item.get("tag", "").lower()
                field_type = item.get("type", "").lower()

                if tag == "select":
                    self._fill_select(element, value)
                elif field_type == "file":
                    element.send_keys(value)
                elif field_type in ("checkbox", "radio"):
                    self._fill_check_or_radio(driver, item, value)
                else:
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
                    element.click()
                    try:
                        element.send_keys(Keys.CONTROL, "a")
                        element.send_keys(Keys.DELETE)
                    except Exception:
                        element.clear()
                    element.send_keys(value)
                item["applied"] = True
            except Exception:
                item["applied"] = False
                continue

    # ------------------------------------------------------------------
    # Screenshot
    # ------------------------------------------------------------------

    def save_screenshot(self, path: str | Path) -> None:
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        self._require_driver().save_screenshot(str(path))

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _require_driver(self) -> Chrome:
        if self.driver is None:
            raise RuntimeError("Browser not started. Call start() first.")
        return self.driver

    @staticmethod
    def _switch_to_live_window(driver: Chrome) -> None:
        handles = list(driver.window_handles)
        if not handles:
            driver.switch_to.new_window("tab")
            handles = list(driver.window_handles)
        if not handles:
            raise RuntimeError("No browser window is available.")

        try:
            current_handle = driver.current_window_handle
            if current_handle in handles:
                driver.switch_to.window(current_handle)
                return
        except Exception:
            pass

        for handle in handles:
            try:
                driver.switch_to.window(handle)
                return
            except Exception:
                continue

        raise RuntimeError("Could not switch to a live browser window.")

    @staticmethod
    def _find_brave_path() -> str:
        """Locate Brave browser executable on Windows."""
        common_paths = [
            os.path.join(os.getenv("PROGRAMFILES", ""), "BraveSoftware", "Brave-Browser", "Application", "brave.exe"),
            os.path.join(os.getenv("PROGRAMFILES(X86)", ""), "BraveSoftware", "Brave-Browser", "Application", "brave.exe"),
            os.path.join(os.getenv("LOCALAPPDATA", ""), "BraveSoftware", "Brave-Browser", "Application", "brave.exe"),
        ]
        for path in common_paths:
            if os.path.exists(path):
                return path
        return ""

    @staticmethod
    def _find_chromedriver_path() -> str:
        cache_dir = Path(".selenium-cache/chromedriver")
        if cache_dir.exists():
            matches = sorted(cache_dir.rglob("chromedriver.exe"), reverse=True)
            if matches:
                return str(matches[0].resolve())
        return ""

    @classmethod
    def _debugger_available(cls) -> bool:
        try:
            request = Request(f"http://{cls.DEBUGGER_ADDRESS}/json/version", headers={"User-Agent": "Mozilla/5.0"})
            with urlopen(request, timeout=2) as response:
                payload = json.loads(response.read().decode("utf-8", errors="ignore"))
            return bool(payload.get("Browser"))
        except (OSError, TimeoutError, ValueError, URLError):
            return False

    def _create_driver(self, use_profile: bool, debugger_address: str | None = None) -> Chrome:
        opts = Options()

        if debugger_address:
            opts.add_experimental_option("debuggerAddress", debugger_address)
        else:
            # Prefer Brave as the default browser for automation.
            # If Brave is installed, use its executable path. Otherwise, fall back to Chrome.
            if self._brave_path:
                opts.binary_location = self._brave_path

            if self.headless:
                opts.add_argument("--headless=new")

            opts.add_argument("--disable-gpu")
            opts.add_argument("--start-maximized")
            opts.add_argument("--disable-blink-features=AutomationControlled")
            opts.add_argument("--remote-debugging-port=9222")
            opts.add_argument("--no-first-run")
            opts.add_argument("--no-default-browser-check")
            opts.add_experimental_option("excludeSwitches", ["enable-automation"])
            if not self.headless:
                opts.add_experimental_option("detach", True)

            if use_profile and self._brave_path:
                AUTOMATION_BROWSER_PROFILE_DIR.mkdir(parents=True, exist_ok=True)
                opts.add_argument(f"--user-data-dir={AUTOMATION_BROWSER_PROFILE_DIR}")
                opts.add_argument("--profile-directory=Default")

        if self._chrome_driver_path:
            service = Service(self._chrome_driver_path)
            return Chrome(service=service, options=opts)
        else:
            service = Service(ChromeDriverManager().install())
            return Chrome(service=service, options=opts)

    def _click_button_by_terms(self, terms: list[str]) -> bool:
        driver = self._require_driver()
        clickable = driver.find_elements(
            By.CSS_SELECTOR,
            "button, a[role='button'], div[role='button'], input[type='submit'], input[type='button']",
        )
        for element in clickable:
            try:
                if not element.is_displayed() or not element.is_enabled():
                    continue
                text = " ".join(
                    part.strip()
                    for part in [
                        element.text or "",
                        element.get_attribute("aria-label") or "",
                        element.get_attribute("value") or "",
                        element.get_attribute("title") or "",
                    ]
                    if part and part.strip()
                ).lower()
                for term in terms:
                    if term in text:
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
                        try:
                            element.click()
                        except Exception:
                            driver.execute_script("arguments[0].click();", element)
                        return True
            except Exception:
                continue
        return False

    @staticmethod
    def _find_label(element: WebElement) -> str:
        elem_id = element.get_attribute("id") or ""
        driver = element.parent
        if elem_id:
            try:
                labels = driver.find_elements(By.CSS_SELECTOR, f"label[for='{elem_id}']")
                if labels:
                    return labels[0].text.strip()
            except Exception:
                pass

            try:
                labelled_by = element.get_attribute("aria-labelledby") or ""
                for label_id in labelled_by.split():
                    label = driver.find_element(By.ID, label_id)
                    if label.text.strip():
                        return label.text.strip()
            except Exception:
                pass

        try:
            labels = element.find_elements(By.XPATH, "./ancestor::label[1]")
            if labels:
                return labels[0].text.strip()
        except Exception:
            pass

        try:
            parent = element.find_element(By.XPATH, "./ancestor::*[self::div or self::fieldset][1]")
            labels = parent.find_elements(By.TAG_NAME, "label")
            if labels:
                return labels[0].text.strip()
            legends = parent.find_elements(By.TAG_NAME, "legend")
            if legends:
                return legends[0].text.strip()
        except Exception:
            pass

        return element.get_attribute("aria-label") or element.get_attribute("placeholder") or ""

    @staticmethod
    def _get_options(element: WebElement, tag: str) -> list[str]:
        if tag != "select":
            return []
        try:
            select = Select(element)
            return [opt.text.strip() for opt in select.options if opt.text.strip()]
        except Exception:
            return []

    @staticmethod
    def _find_section_context(element: WebElement) -> str:
        section_queries = [
            "./ancestor::fieldset[1]//legend[1]",
            "./ancestor::section[1]//*[self::h1 or self::h2 or self::h3 or self::h4][1]",
            "./ancestor::*[self::div or self::form][1]//*[self::h1 or self::h2 or self::h3 or self::h4][1]",
        ]
        for query in section_queries:
            try:
                for match in element.find_elements(By.XPATH, query):
                    text = " ".join((match.text or "").split()).strip()
                    if text:
                        return text
            except Exception:
                continue

        try:
            parent = element.find_element(By.XPATH, "./ancestor::*[@aria-label or @aria-labelledby][1]")
            text = (parent.get_attribute("aria-label") or "").strip()
            if text:
                return text
        except Exception:
            pass

        return ""

    @staticmethod
    def _build_xpath(element: WebElement, tag: str, name: str, elem_id: str) -> str:
        if elem_id:
            return f"//{tag}[@id='{elem_id}']"
        if name:
            return f"//{tag}[@name='{name}']"
        return ""

    @staticmethod
    def _locate_field(driver: Chrome, item: dict[str, Any]) -> WebElement | None:
        selector = item.get("selector", "")
        xpath = item.get("xpath", "")
        name = item.get("name", "")

        if selector:
            try:
                return driver.find_element(By.CSS_SELECTOR, selector)
            except Exception:
                pass

        if xpath:
            try:
                return driver.find_element(By.XPATH, xpath)
            except Exception:
                pass

        if name:
            try:
                return driver.find_element(By.NAME, name)
            except Exception:
                pass

        return None

    @staticmethod
    def _fill_select(element: WebElement, value: str) -> None:
        select = Select(element)
        value_lower = value.lower()
        for option in select.options:
            if option.text.strip().lower() == value_lower:
                select.select_by_visible_text(option.text.strip())
                return
        for option in select.options:
            if value_lower in option.text.strip().lower():
                select.select_by_visible_text(option.text.strip())
                return

    @staticmethod
    def _fill_check_or_radio(driver: Chrome, item: dict[str, Any], value: str) -> None:
        name = item.get("name", "")
        field_id = item.get("id", "")
        value_lower = value.lower()
        elements: list[WebElement] = []
        if name:
            elements.extend(driver.find_elements(By.NAME, name))
        if not elements and field_id:
            try:
                elements.append(driver.find_element(By.ID, field_id))
            except Exception:
                pass
        for el in elements:
            label_text = ""
            try:
                label_text = BrowserSession._find_label(el).strip().lower()
            except Exception:
                pass

            el_value = (el.get_attribute("value") or "").lower()
            if value_lower in label_text or value_lower == el_value or el_value in value_lower:
                if not el.is_selected():
                    try:
                        el.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", el)
                return

    @classmethod
    def _looks_like_job_page(cls, title: str, url: str) -> bool:
        lowered = f"{title} {url}".lower()
        return any(hint in lowered for hint in cls.JOB_SITE_HINTS) or "job" in lowered or "career" in lowered or "apply" in lowered

    @staticmethod
    def _looks_like_role_text(text: str) -> bool:
        lowered = " ".join(text.split()).strip().lower()
        if len(lowered) < 4 or len(lowered) > 120:
            return False
        blocked_terms = ("sign in", "log in", "apply now", "next", "continue", "submit", "search jobs")
        if lowered in blocked_terms:
            return False
        role_terms = (
            "engineer",
            "developer",
            "tester",
            "qa",
            "analyst",
            "manager",
            "consultant",
            "specialist",
            "designer",
            "administrator",
            "coordinator",
            "architect",
            "lead",
            "intern",
        )
        return any(term in lowered for term in role_terms)


def start_browser() -> tuple[Any, Chrome, Any]:
    """Backward-compatible helper that returns (None, driver, driver).

    Old callers used ``p, browser, page = start_browser()``.
    With Selenium there is no separate playwright handle or page object,
    so the driver is returned in place of both *browser* and *page*.
    """
    session = BrowserSession(headless=False)
    session.start()
    driver = session.driver
    assert driver is not None, "Browser failed to start"
    return None, driver, driver
