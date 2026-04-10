from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from config import PROFILE_PATH
from core.ai_engine import generate_cover_letter, generate_answers
from core.browser import BrowserSession
from core.easy_apply import EasyApplyBot


def load_profile() -> dict[str, Any]:
    profile_path = PROFILE_PATH if PROFILE_PATH.exists() else Path("profile.json")
    with profile_path.open("r", encoding="utf-8") as file:
        return json.load(file)


def main() -> None:
    profile = load_profile()

    resume_path = input("Upload resume path: ").strip()
    job_pref = input("Enter job preference: ").strip()

    mode = input("1 = Paste link | 2 = Search: ").strip()

    browser_session = BrowserSession(headless=False)
    attached = browser_session.start(attach_to_existing=True)
    if attached:
        print("Attached to your existing Brave window.")
    else:
        print("Could not attach to an existing Brave window. Opened a separate automation window.")

    if mode == "1":
        url = input("Paste job link: ").strip()
        browser_session.open(url)
    else:
        print("Auto-search not implemented in this CLI version. Use app_gui.py for auto-search.")
        return

    input("Login if needed, then press ENTER...")

    browser_session.wait_for_page_settle()
    page_text = browser_session.extract_page_text()

    cover = generate_cover_letter(page_text, profile)
    answers_text = generate_answers(page_text)

    print("Generated Cover Letter:\n", cover)
    print("\nGenerated Answers:\n", answers_text)

    fields = browser_session.collect_inputs()
    bot = EasyApplyBot()
    result = bot.prepare_application(
        profile=profile, page_text=page_text, fields=fields, extra_answers={"resume_path": resume_path}
    )

    browser_session.apply_fill_plan(result["fill_plan"])

    print("\nAll fields filled.")
    if result["missing_fields"]:
        print(f"Warning: {len(result['missing_fields'])} field(s) remain unfilled.")

    input("Review and press ENTER to finish...")

    browser_session.stop()

if __name__ == "__main__":
    main()
