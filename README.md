# MCQS-CREATER
# 📝 MCQ Generator

Type a topic (and optionally upload notes) and get up to **100 multiple-choice questions** with answers and explanations.

Built with **Streamlit** and **Google Gemini Flash**.

## Features

- Type any topic, or upload notes (PDF, DOCX or TXT) so the questions are based on your material. Scanned PDFs are read by Gemini directly.
- Sidebar options: difficulty (Easy / Medium / Hard / Mixed), number of questions (10-100, default 100), options per question (3-5, default 4), and language.
- Questions are generated in **several batches of 20-25** with a progress bar, because one huge request often fails or repeats itself.
- Duplicate and near-duplicate questions are removed. If some are removed, extra top-up batches fill the gap.
- Answer options are shuffled in code, so the correct answer is not always the same letter.
- **Practice mode:** answer on screen, submit, and see your score with the correct answers and explanations.
- **Study mode:** all questions with the correct answers marked.
- Download everything as **Markdown, CSV or JSON**.
- Clear errors for a bad API key, rate limit, wrong model name (with automatic fallback) and broken AI output (retries). If a later batch fails, you still keep the questions generated so far.

## Run locally

```bash
git clone https://github.com/<your-username>/mcq-generator.git
cd mcq-generator
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

Get a free Gemini API key at https://aistudio.google.com/apikey, then either:

- paste it into the "Gemini API key" box on the main page, or
- set it as an environment variable: `export GEMINI_API_KEY="your-key"`, or
- create `.streamlit/secrets.toml` containing `GEMINI_API_KEY = "your-key"`

## Deploy on Streamlit Community Cloud

1. Push this repo to GitHub.
2. Go to https://share.streamlit.io and sign in with GitHub.
3. Click **Create app**, choose the repo, branch `main`, and main file `app.py`.
4. Open **Advanced settings → Secrets** and add:
   ```toml
   GEMINI_API_KEY = "your-key"
   ```
5. Click **Deploy**.

## Project structure

```
app.py            # Streamlit app
requirements.txt  # Dependencies
README.md
.gitignore        # Keeps secrets out of Git
```

## Notes

- Generating 100 questions takes roughly 1-3 minutes (4 batches plus any top-ups).
- The free Gemini tier has rate limits. If you hit one, wait a minute and try again, or ask for fewer questions. Any questions already generated are kept.
- If the AI keeps repeating itself, you may get fewer than requested (the app tells you). Try a broader topic or add notes.
- Your topic and notes are sent to Google's Gemini API. Don't upload anything you aren't comfortable sharing.
- The default model is `gemini-flash-latest`; change it in the sidebar if you like. A wrong model name falls back to a working one automatically.
- Questions are AI-generated. Double-check important facts before using them for real exams.

## Troubleshooting

**`ModuleNotFoundError` / "package is not installed" on Streamlit Cloud**

The host did not install the packages. The app tries to install them itself on first load (about a minute), but the proper fix is:

1. In your GitHub repo, make sure a file named exactly `requirements.txt` (with an **s**) is in the **root**, next to `app.py`, not inside a folder.
   - On GitHub: **Add file → Create new file**, name it `requirements.txt`, paste the four lines below, then **Commit changes**.
2. It must contain these lines - note `python-docx`, not `docx`:
   ```
   streamlit>=1.40.0
   google-genai>=1.0.0
   pypdf>=5.0.0
   python-docx>=1.1.0
   ```
3. Open your app, click **Manage app → ⋮ → Reboot app**, and check the logs for pip errors.
