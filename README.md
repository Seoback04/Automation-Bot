# Job Bot v2
Job Bot v2 is a desktop and CLI automation tool for pre-filling job applications with Selenium.

The main workflow runs from a Tkinter GUI, opens job pages in Brave, loads your saved profile and resume, detects Easy Apply or external apply flows, fills recognized fields, saves a review screenshot, and asks for confirmation before submit.

## Features
- Tkinter GUI for day-to-day use
- Brave browser automation with reuse of your existing Brave profile
- Manual job URL mode and auto-search mode
- Resume handling and profile persistence
- Local answer generation and cover letter generation
- Multi-step apply flow support
- Review screenshot before submission

## Requirements
- Windows
- Python 3.13 recommended
- Brave Browser installed
- Internet access for package installation and job search

## Installation
1. Create a virtual environment:

```powershell
python -m venv .venv
```

2. Activate it:

```powershell
.\.venv\Scripts\Activate.ps1
```

3. Install dependencies:

```powershell
python -m pip install --upgrade pip
pip install -r requirements.txt
```

The current dependency list is minimal:

```text
selenium>=4.20.0
```

Selenium 4 can manage the browser driver automatically in many cases.

## Browser behavior
The app is configured to open Brave instead of Chrome.

- It looks for `brave.exe` in standard Windows install paths.
- It reuses the default Brave profile so you do not need to log in every run.
- The profile path used is typically:
  - `%LOCALAPPDATA%\BraveSoftware\Brave-Browser\User Data`

If the session is not reused correctly:
- Close all running Brave windows
- Start the app again
- Make sure you are already logged in to the target site in your Brave default profile

## Project setup
The app reads profile data from `profile.json` in the project root.

At minimum, make sure the nested `basics` and `job_preferences` sections are filled. The GUI can also update some of these values for you.

Example structure:

```json
{
  "basics": {
    "name": "Your Name",
    "email": "you@example.com",
    "phone": "+64XXXXXXXX",
    "location": "Auckland",
    "linkedin": "https://linkedin.com/in/yourprofile",
    "github": "https://github.com/yourprofile",
    "website": "",
    "resume_url": "",
    "resume_path": "B:\\Bot\\data\\resumes\\your_resume.pdf",
    "summary": "Short professional summary"
  },
  "preferences": {
    "work_authorized": "Yes",
    "requires_sponsorship": "No",
    "salary_expectation": "",
    "notice_period": ""
  },
  "job_preferences": {
    "role": "Software Tester",
    "location": "Auckland"
  }
}
```

You can also use the GUI to pick a resume file directly. The app copies the selected file into `data/resumes/`.

## Running the GUI
You can start the GUI in any of these ways:

```powershell
python app_gui.py
```

```powershell
.\start_job_bot_gui.bat
```

```powershell
.\Launch_JobBot_GUI.bat
```

`Launch_JobBot_GUI.bat` installs dependencies first and then starts the GUI with `pythonw`.

## GUI workflow
1. Select `Profile JSON`.
2. Select `Resume File`.
3. Enter `Role`.
4. Optionally enter `Location`.
5. Choose one job source:
   - `Paste Job Link`
   - `Auto Search Jobs`
6. If using manual mode, enter the `Job URL`.
7. Optionally enable `Headless Browser`.
8. Click `Start`.

During execution, the app will:
- load and normalize profile data
- copy the selected resume into `data/resumes/`
- open the target job page in Brave
- detect the apply flow type
- extract fields from the page
- generate answers from your profile
- prompt for missing fields when necessary
- save a screenshot before submit

The GUI writes status messages to the `Execution Log`.

## Running the CLI
The CLI entry point is `main.py`:

```powershell
python main.py
```

Current CLI behavior:
- supports manual job URL mode
- loads the profile from `profile.json`
- opens the page in Brave
- parses page text
- builds a fill plan
- fills detected fields

The CLI path is more limited than the GUI. For auto-search and full flow handling, use `app_gui.py`.

## Generated files
The application writes artifacts into `data/`:

- `data/resumes/` for copied resume files
- `data/cover_letters/` for generated cover letters
- `data/screenshots/` for review screenshots

## Project structure
- `app_gui.py` — main desktop application
- `main.py` — CLI entry point
- `config.py` — shared paths and constants
- `core/` — browser, parsing, filling, search, profile storage, AI helpers
- `flows/` — Easy Apply and external apply flow orchestration
- `utils/` — selectors and helper utilities
- `data/` — resumes, cover letters, screenshots, sample files

## Troubleshooting
### Resume file not found
Use the `Browse` button and choose a valid file.

Supported file types in the GUI:
- `.pdf`
- `.doc`
- `.docx`
- `.txt`

### Job URL is required
If `Paste Job Link` is selected, you must enter a job URL before starting.

### Auto-search found nothing
Switch to manual URL mode and paste the job page directly.

### Brave opens without your logged-in session
- Close Brave fully
- Sign in to the target site in Brave
- Start the app again

### Form fields are missed
Some job sites use custom controls that are not detected reliably. In those cases:
- fill the remaining fields manually
- answer prompts for missing values when shown

## Notes
- The application is designed for review-before-submit workflows, not blind background submission.
- Site layouts change often, so some pages may need manual review or manual completion.
