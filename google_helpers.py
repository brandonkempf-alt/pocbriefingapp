# google_helpers.py
# Utilities for Google OAuth, Calendar, Drive, Slides, and company detection.

import os
import re
import datetime as dt
from collections import Counter
from typing import List, Dict, Optional

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# ---------------------------------------------------------------------
# OAuth / Scopes
# ---------------------------------------------------------------------

# Scopes needed for:
# - Calendar (read events)
# - Drive (copy template + read link)
# - Slides (replace placeholders)
SCOPES = [
    "https://www.googleapis.com/auth/calendar.readonly",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/presentations",
]

def _token_path() -> str:
    """Where we cache the OAuth token."""
    os.makedirs(".tokens", exist_ok=True)
    return os.path.join(".tokens", "google_token.json")

def get_google_creds(client_json_path: str) -> Credentials:
    """
    Return authorized Google API Credentials.
    - client_json_path: path to OAuth client JSON (Desktop App)
    Stores/refreshes token at .tokens/google_token.json
    """
    creds: Optional[Credentials] = None
    token_file = _token_path()

    if os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(client_json_path, SCOPES)
            # Opens a local browser for consent; random localhost port is fine for Desktop apps
            creds = flow.run_local_server(port=0, prompt="consent")
        with open(token_file, "w") as f:
            f.write(creds.to_json())

    return creds

# ---------------------------------------------------------------------
# Google APIs: Calendar / Drive / Slides
# ---------------------------------------------------------------------

def list_events(
    creds: Credentials,
    calendar_id: str,
    q: str,
    start_iso: str,
    end_iso: str,
) -> List[Dict]:
    """List events between start_iso and end_iso, filtered by q."""
    service = build("calendar", "v3", credentials=creds)
    events_result = service.events().list(
        calendarId=calendar_id,
        q=q,
        timeMin=start_iso,
        timeMax=end_iso,
        singleEvents=True,
        orderBy="startTime",
    ).execute()
    return events_result.get("items", [])

def copy_template(creds: Credentials, template_id: str, new_name: str) -> str:
    """Copy a Drive file (Slides template) and return the new file ID."""
    drive = build("drive", "v3", credentials=creds)
    copied = drive.files().copy(fileId=template_id, body={"name": new_name}).execute()
    return copied["id"]

def fill_slides_placeholders(creds: Credentials, presentation_id: str, replacements: Dict[str, str]) -> None:
    """
    Replace placeholder tokens in a Google Slides presentation.
    Example replacements:
      {
        "{{CompanyName}}": "Acme",
        "{{POCEmails}}": "a@acme.com, b@acme.com",
        "{{EventTitle}}": "POC Kickoff",
        "{{EventTime}}": "2025-08-12T10:00:00-05:00"
      }
    """
    slides = build("slides", "v1", credentials=creds)
    requests = []
    for token, value in replacements.items():
        requests.append({
            "replaceAllText": {
                "containsText": {"text": token, "matchCase": True},
                "replaceText": value
            }
        })
    slides.presentations().batchUpdate(
        presentationId=presentation_id,
        body={"requests": requests}
    ).execute()

def get_web_view_link(creds: Credentials, file_id: str) -> str:
    """Return Drive webViewLink for a file."""
    drive = build("drive", "v3", credentials=creds)
    meta = drive.files().get(fileId=file_id, fields="webViewLink").execute()
    return meta.get("webViewLink", "")

# ---------------------------------------------------------------------
# Company detection
# ---------------------------------------------------------------------

def attendees_to_company(attendees, event_summary: Optional[str] = None, event_description: Optional[str] = None) -> Dict[str, str]:
    """
    Guess the company from attendee emails and optional event text.
    Returns: {"companyName": str, "companyDomain": str}

    Config via .env (all optional; comma-separated for lists):
      INTERNAL_DOMAINS  - domains to ignore as "personal/internal"
      IGNORE_DOMAINS    - utility domains to ignore (zoom.us, etc.)
      DOMAIN_PRIORITY   - if multiple externals, prefer these
      DOMAIN_NAME_MAP   - explicit mapping "fb.com:Meta,google.com:Google"
    """
    INTERNAL = set([
        d.strip().lower() for d in os.getenv(
            "INTERNAL_DOMAINS",
            "kempfenterprise.com,gmail.com,outlook.com,hotmail.com,yahoo.com,"
            "icloud.com,aol.com,proton.me,protonmail.com,me.com"
        ).split(",")
        if d.strip()
    ])
    IGNORE = set([
        d.strip().lower() for d in os.getenv(
            "IGNORE_DOMAINS",
            "zoom.us,meetup.com,calendar.google.com,teams.microsoft.com"
        ).split(",")
        if d.strip()
    ])
    PRIORITY = [
        d.strip().lower() for d in os.getenv("DOMAIN_PRIORITY", "").split(",") if d.strip()
    ]
    NAME_MAP = {}
    for pair in [p.strip() for p in os.getenv("DOMAIN_NAME_MAP", "").split(",") if ":" in p]:
        dom, name = pair.split(":", 1)
        NAME_MAP[dom.strip().lower()] = name.strip()

    def base_domain(d: str) -> str:
        d = (d or "").lower().strip().replace("@", "")
        if d.startswith("www."):
            d = d[4:]
        parts = d.split(".")
        if len(parts) >= 3:
            # heuristic: keep last two labels (works for most corporate domains)
            d = ".".join(parts[-2:])
        return d

    def pretty_from_domain(d: str) -> str:
        core = base_domain(d).split(".")[0]
        return core.capitalize() if core else "Unknown"

    # 1) Collect candidate domains from attendee emails
    domains: List[str] = []
    for a in attendees or []:
        if isinstance(a, dict):
            email = (a.get("email") or "").lower().strip()
        else:
            email = str(a).lower().strip()
        if "@" in email:
            domains.append(base_domain(email.split("@", 1)[1]))

    # Remove internal and utility domains
    external = [d for d in domains if d and d not in INTERNAL and d not in IGNORE]

    # 2) Priority domain wins if present
    for p in PRIORITY:
        pb = base_domain(p)
        if pb in external:
            return {
                "companyDomain": pb,
                "companyName": NAME_MAP.get(pb, pretty_from_domain(pb))
            }

    # 3) Majority vote among external domains
    if external:
        dom, _ = Counter(external).most_common(1)[0]
        return {
            "companyDomain": dom,
            "companyName": NAME_MAP.get(dom, pretty_from_domain(dom))
        }

    # 4) Fallback: any non-internal non-ignored domain at all
    for d in domains:
        if d and d not in INTERNAL and d not in IGNORE:
            return {"companyDomain": d, "companyName": NAME_MAP.get(d, pretty_from_domain(d))}

    # 5) Title/description fallback (simple heuristics)
    text = " ".join([(event_summary or ""), (event_description or "")])
    m = re.search(r"(?:with|for)\s+([A-Z][A-Za-z0-9&.\- ]{2,})", text, flags=re.IGNORECASE) \
        or re.search(r"POC[:\-]\s*([A-Z][A-Za-z0-9&.\- ]{2,})", text, flags=re.IGNORECASE) \
        or re.search(r"\(([A-Z][A-Za-z0-9&.\- ]{2,})\)", text)
    if m:
        cand = m.group(1).strip()
        if cand.lower() not in {"poc", "call", "meeting", "demo", "discovery"}:
            return {"companyDomain": "", "companyName": cand}

    # 6) Unknown
    return {"companyDomain": "", "companyName": "Unknown"}
