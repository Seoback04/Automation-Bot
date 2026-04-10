from __future__ import annotations

from datetime import datetime
from typing import Any, Callable

from config import COVER_LETTERS_DIR, MAX_EASY_APPLY_STEPS
from core.easy_apply import EasyApplyBot
from core.profile_store import ProfileStore


MissingPromptFn = Callable[[list[dict[str, Any]]], bool]


class EasyApplyFlow:
    """Handles multi-step Easy Apply flows."""

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
            browser.click_apply_entry("easy_apply")
            browser.wait_for_page_settle()

        final_result: dict[str, Any] = {"steps_completed": 0, "missing_fields": []}
        for step in range(1, MAX_EASY_APPLY_STEPS + 1):
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
                    # Recompute with newly learned answers.
                    result = self.bot.prepare_application(
                        profile=profile,
                        page_text=page_text,
                        fields=fields,
                        extra_answers={"resume_path": resume_path, "resume_file": resume_path},
                    )

            browser.apply_fill_plan(result["fill_plan"])
            self.bot.record_step(
                profile=profile,
                step_number=step,
                fields=fields,
                result=result,
            )
            final_result = result
            final_result["steps_completed"] = step

            if not browser.click_next_step():
                break
            browser.wait_for_page_settle()

        self._save_cover_letter(final_result.get("answers", {}))
        return final_result

    @staticmethod
    def _save_cover_letter(answers: dict[str, Any]) -> None:
        letter = str(answers.get("cover_letter", "")).strip()
        if not letter:
            return

        COVER_LETTERS_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = COVER_LETTERS_DIR / f"cover_letter_{stamp}.txt"
        path.write_text(letter, encoding="utf-8")
