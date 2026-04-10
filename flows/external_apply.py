from __future__ import annotations

from datetime import datetime
from typing import Any, Callable

from config import COVER_LETTERS_DIR
from core.easy_apply import EasyApplyBot
from core.profile_store import ProfileStore


MissingPromptFn = Callable[[list[dict[str, Any]]], bool]


class ExternalApplyFlow:
    """Handles standard external application pages."""

    def __init__(
        self,
        bot: EasyApplyBot | None = None,
        profile_store: ProfileStore | None = None,
    ) -> None:
        self.bot = bot or EasyApplyBot(profile_store=profile_store)

    def run(
        self,
        browser: Any,
        profile: dict[str, Any],
        resume_path: str,
        prompt_missing: MissingPromptFn,
    ) -> dict[str, Any]:
        if not browser.has_application_form():
            browser.click_apply_entry("external_apply")
            browser.wait_for_page_settle()

        page_text = browser.extract_page_text()
        fields = browser.collect_inputs()
        result = self.bot.prepare_application(
            profile=profile,
            page_text=page_text,
            fields=fields,
            extra_answers={"resume_path": resume_path, "resume_file": resume_path},
        )

        if result["missing_fields"]:
            changed = prompt_missing(result["missing_fields"])
            if changed:
                result = self.bot.prepare_application(
                    profile=profile,
                    page_text=page_text,
                    fields=fields,
                    extra_answers={"resume_path": resume_path, "resume_file": resume_path},
                )

        browser.apply_fill_plan(result["fill_plan"])
        self.bot.record_step(
            profile=profile,
            step_number=1,
            fields=fields,
            result=result,
        )
        self._save_cover_letter(result.get("answers", {}))
        return result

    @staticmethod
    def _save_cover_letter(answers: dict[str, Any]) -> None:
        letter = str(answers.get("cover_letter", "")).strip()
        if not letter:
            return

        COVER_LETTERS_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = COVER_LETTERS_DIR / f"cover_letter_{stamp}.txt"
        path.write_text(letter, encoding="utf-8")
