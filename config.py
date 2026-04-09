from __future__ import annotations

from pathlib import Path
import sys


if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
CORE_DIR = BASE_DIR / "core"
FLOWS_DIR = BASE_DIR / "flows"
UTILS_DIR = BASE_DIR / "utils"

DATA_DIR = BASE_DIR / "data"
RESUMES_DIR = DATA_DIR / "resumes"
COVER_LETTERS_DIR = DATA_DIR / "cover_letters"
SCREENSHOTS_DIR = DATA_DIR / "screenshots"
RUN_HISTORY_PATH = DATA_DIR / "run_history.json"
APP_STATE_PATH = DATA_DIR / "app_state.json"
APPLICATION_PIPELINE_PATH = DATA_DIR / "application_pipeline.json"
REPORTS_DIR = DATA_DIR / "reports"
AUTOMATION_BROWSER_PROFILE_DIR = DATA_DIR / "browser_profile"

PROFILE_PATH = BASE_DIR / "profile.json"

DEFAULT_MODEL = "gpt-4.1-mini"
MAX_EASY_APPLY_STEPS = 5
AUTO_SEARCH_LIMIT = 8
PAGE_SETTLE_SECONDS = 1.0
