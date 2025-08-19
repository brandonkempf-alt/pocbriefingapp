# POC Briefing App (Local – Streamlit, Google Slides Template)

A local web app that:
1) Lists upcoming Google Calendar events that match a term (e.g., "POC")
2) Copies a **Google Slides** template and fills placeholders (text)
3) Generates a company summary with OpenAI
4) (Optional) Posts the brief to Slack

## Quick start
```bash
python3 -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env
```

### Google Cloud OAuth (one-time)
- Enable APIs: **Google Calendar API**, **Google Drive API**, **Google Slides API**
- Create OAuth client: **Desktop app**, download as `client_secret.json` in project root.

### Run
```bash
streamlit run app.py
```

## Placeholders in your **Slides** template
Use these exact tokens in text boxes (any slide):
- {{CompanyName}}
- {{POCEmails}}
- {{EventTitle}}
- {{EventTime}}

> We use Slides `presentations.batchUpdate` → `replaceAllText` to replace these tokens.

## Notes
- First run opens a browser for Google consent; token is cached under `.tokens/`.
- If your Calendar ID is not `me`, set `GOOGLE_CALENDAR_ID` in `.env`.
# pocbriefingapp
