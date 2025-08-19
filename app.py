import os
import json
import datetime as dt
import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI
import requests

from google_helpers import (
    get_google_creds,
    list_events,
    copy_template,
    fill_slides_placeholders,
    get_web_view_link,
    attendees_to_company,
)

# ---------------------------
# Env & constants
# ---------------------------
load_dotenv()

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
SLACK_WEBHOOK_URL = os.getenv("SLACK_WEBHOOK_URL", "")
GOOGLE_CALENDAR_ID = os.getenv("GOOGLE_CALENDAR_ID", "primary")
GOOGLE_OAUTH_CLIENT_JSON = os.getenv("GOOGLE_OAUTH_CLIENT_JSON", "client_secret.json")
GSLIDES_TEMPLATE_ID = os.getenv("GSLIDES_TEMPLATE_ID", "")
EXCLUDED_DOMAIN = os.getenv("EXCLUDED_DOMAIN","kempfenterprise.com").lower().strip()
EXCLUSION_MODE = os.getenv("EXCLUSION_MODE", "any").lower().strip()  # any | all

# Optional: define resources via JSON in .env (RESOURCES_JSON='[{"label":"...","url":"..."},...]')
RESOURCES_JSON = os.getenv("RESOURCES_JSON", "").strip()

DEFAULT_RESOURCES = [
    {"label": "AWS Connection Guide", "url": "https://help.drata.com/en/articles/5048935-aws-amazon-web-services-connection"},
    {"label": "Azure Connection Guide", "url": "https://help.drata.com/en/articles/5032404-azure-connection"},
    {"label": "GCP Connection Guide", "url": "https://help.drata.com/en/articles/4663373-gcp-google-cloud-platform-connection"},
    {"label": "Intune Connection Guide", "url": "https://help.drata.com/en/articles/5604949-intune-windows-connection"},
    {"label": "EntraID/O365 Connection Guide", "url": "https://help.drata.com/en/articles/4797766-microsoft-365-connection"},
    {"label": "Okta Connection Guide", "url": "https://help.drata.com/en/articles/5608136-okta-connection-for-identity-management"},
    {"label": "Jira Connection Guide", "url": "https://help.drata.com/en/articles/4663378-jira-connection"},
]

try:
    RESOURCES = json.loads(RESOURCES_JSON) if RESOURCES_JSON else DEFAULT_RESOURCES
    # Ensure exactly 6 items (trim or pad if needed)
    if len(RESOURCES) < 7:
        RESOURCES = RESOURCES + DEFAULT_RESOURCES[: 7 - len(RESOURCES)]
    RESOURCES = RESOURCES[:7]
except Exception:
    RESOURCES = DEFAULT_RESOURCES

# ---------------------------
# Streamlit config & header
# ---------------------------
st.set_page_config(page_title="POC Briefing App", layout="wide")
st.title("POC Briefing App – by Hack_Street_Boyz")

# Sidebar toggle for Slack auto-send
if "auto_slack" not in st.session_state:
    st.session_state["auto_slack"] = True  # set False if you prefer default OFF

with st.sidebar:
    st.session_state["auto_slack"] = st.checkbox(
        "Send to Slack automatically",
        value=st.session_state["auto_slack"],
        help="If on, each generated brief is posted to Slack via webhook."
    )

# ---------------------------
# Summary cache in session state
# ---------------------------
if "summaries" not in st.session_state:
    st.session_state["summaries"] = {}   # { event_id: summary_text }
if "current_event_id" not in st.session_state:
    st.session_state["current_event_id"] = None

# ---------------------------
# Helpers
# ---------------------------
def post_to_slack_if_enabled(summary, link, company_name, event_title, event_time, poc_emails, resources_text, slack_webhook):
    """
    Returns True if a Slack message was sent, False otherwise.
    Respects the sidebar toggle; never sends if toggle is OFF or webhook missing.
    """
    if not st.session_state.get("auto_slack"):
        return False  # toggle OFF
    if not slack_webhook:
        return False  # no webhook configured

    # Add the resources as a single mrkdwn block (simple bullets)
    resources_block = {"type": "section", "text": {"type": "mrkdwn", "text": resources_text}} if resources_text else None

    blocks = [
        {"type": "header", "text": {"type": "plain_text", "text": f"POC Slides Brief — {company_name}", "emoji": True}},
        {
            "type": "section",
            "fields": [
                {"type": "mrkdwn", "text": f"*Event:*\n{event_title}"},
                {"type": "mrkdwn", "text": f"*When:*\n{event_time}"},
                {"type": "mrkdwn", "text": f"*POC Emails:*\n{poc_emails or '—'}"},
            ]
        },
        {"type": "section", "text": {"type": "mrkdwn", "text": summary or "_(no summary)_" }},
        {
            "type": "actions",
            "elements": [
                {"type": "button", "text": {"type": "plain_text", "text": "Open Slides"}, "url": link, "style": "primary"}
            ]
        }
    ]
    if resources_block:
        # Insert resources above the actions row so the button stays at the bottom
        blocks.insert(3, resources_block)

    try:
        r = requests.post(slack_webhook, json={"blocks": blocks}, timeout=10)
        r.raise_for_status()
        return True
    except Exception:
        return False


def collect_event_emails(e):
    """Collect attendee + organizer + creator emails; return list of lower-cased emails (deduped)."""
    emails = []

    # Attendees (dict or str)
    for a in (e.get("attendees") or []):
        if isinstance(a, dict):
            em = (a.get("email") or "").strip().lower()
            if em:
                emails.append(em)
        elif isinstance(a, str):
            emails.append(a.strip().lower())

    # Organizer / Creator
    for key in ("organizer", "creator"):
        info = e.get(key) or {}
        em = (info.get("email") or "").strip().lower()
        if em:
            emails.append(em)

    # Deduplicate keep order
    out = []
    seen = set()
    for em in emails:
        if em not in seen:
            seen.add(em)
            out.append(em)
    return out


def exclude_event_by_domain(all_emails, excluded_domain, mode="any"):
    """
    Return True if the event should be excluded based on domain.
    mode='any'  -> exclude if ANY email ends with @excluded_domain
    mode='all'  -> exclude if ALL emails (and at least one exists) end with @excluded_domain
    """
    if not excluded_domain or not all_emails:
        return False
    suffix = "@" + excluded_domain
    if mode == "all":
        return all(em.endswith(suffix) for em in all_emails)
    return any(em.endswith(suffix) for em in all_emails)


def format_resources_text(selected_resources):
    """Return a markdown string like:
    • Label: URL
    • Label: URL
    """
    if not selected_resources:
        return ""
    lines = [f"• *{r['label']}*: {r['url']}" for r in selected_resources]
    return "\n".join(lines)


def build_event_summary(event_row: dict, company: dict, resources_text: str = "") -> str:
    """Call OpenAI to summarize the event/company. Returns summary text."""
    if not OPENAI_API_KEY:
        return ""
    client = OpenAI(api_key=OPENAI_API_KEY)

    # Inputs
    event_title = event_row.get("summary", "")
    event_time  = event_row.get("start", "")
    poc_emails  = ", ".join(collect_event_emails(event_row.get("_raw", {})))

    prompt = f"""
You are assisting with **Prospect Discovery** research, using the framework defined in this project:
https://chatgpt.com/g/g-686d26f0ea2481919ca8df79dc949359-prospect-discovery

Generate a short pre-POC briefing that aligns with that methodology:
- Summarize the company (industry, product focus, approximate size if known).
- Include one headline-level piece of recent news if available from 2025.
- Highlight any discovery-relevant context that would inform the sales call.
- Keep the output under 1000 words and concise for a Slack/Slides briefing.

Company: {company.get('companyName','')} ({company.get('companyDomain','')})
POC Emails: {poc_emails}
Meeting: {event_title} on {event_time}
Resources:
{resources_text or '—'}
"""
    res = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "You are a concise sales research assistant."},
            {"role": "user", "content": prompt},
        ],
        temperature=0.4,
    )
    return (res.choices[0].message.content or "").strip()

# ---------------------------
# Inputs
# ---------------------------
col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    start = st.date_input("Start", dt.date.today())
with col2:
    end = st.date_input("End", dt.date.today() + dt.timedelta(days=14))
with col3:
    search = st.text_input("Search term", "POC")

# ---------------------------
# Load events first
# ---------------------------
if st.button("Load events"):
    try:
        creds = get_google_creds(GOOGLE_OAUTH_CLIENT_JSON)
        events = list_events(
            creds,
            calendar_id="primary" if GOOGLE_CALENDAR_ID in ("me", "primary", "") else GOOGLE_CALENDAR_ID,
            q=search,
            start_iso=dt.datetime.combine(start, dt.time.min).isoformat() + "Z",
            end_iso=dt.datetime.combine(end, dt.time.max).isoformat() + "Z",
        )

        rows = []
        for e in events:
            all_emails = collect_event_emails(e)
            # Apply domain exclusion per config
            if exclude_event_by_domain(all_emails, EXCLUDED_DOMAIN, mode=EXCLUSION_MODE):
                continue

            rows.append({
                "id": e.get("id", ""),
                "summary": e.get("summary", ""),
                "start": e.get("start", {}).get("dateTime") or e.get("start", {}).get("date") or "",
                "attendees": ", ".join(all_emails),
                "_raw": e
            })

        st.session_state["rows"] = rows
        st.success(f"Loaded {len(rows)} event(s).")

    except Exception as ex:
        st.error(f"Error loading events: {ex}")

# Show table regardless (empty if not loaded yet)
rows = st.session_state.get("rows", [])
st.subheader("Events")
st.dataframe(rows, use_container_width=True)

# Row select (only matters if rows exist)
idx = st.number_input(
    "Row index",
    min_value=0,
    max_value=max(len(rows) - 1, 0),
    value=0,
    step=1
) if rows else 0

# ---------------------------
# Step 2: show the 6 checkbox URLs AFTER events load
# ---------------------------
selected_resources = []
if rows:
    st.markdown("### Include resources in this brief")
    # store each checkbox in session_state to preserve between reruns
    for i, res in enumerate(RESOURCES):
        key = f"res_{i}"
        if key not in st.session_state:
            st.session_state[key] = False
        st.session_state[key] = st.checkbox(f"{res['label']} — {res['url']}", value=st.session_state[key])
        if st.session_state[key]:
            selected_resources.append(res)

# ---------------------------
# Auto-generate summary when the selected row changes
# ---------------------------
if rows:
    selected_row = rows[idx]
    event_id = selected_row.get("id")

    # Build resources_text from current checkboxes (if any selected)
    current_selected_resources = []
    for i in range(6):
        key = f"res_{i}"
        if key in st.session_state and st.session_state[key]:
            current_selected_resources.append(RESOURCES[i])
    resources_text_for_cache = format_resources_text(current_selected_resources)

    if event_id and event_id != st.session_state["current_event_id"]:
        st.session_state["current_event_id"] = event_id

        # Compute company for this event
        raw_event = selected_row["_raw"]
        company_for_cache = attendees_to_company(
            raw_event.get("attendees", []),
            event_summary=raw_event.get("summary", ""),
            event_description=raw_event.get("description", ""),
        )

        # If not already cached, generate and cache
        if event_id not in st.session_state["summaries"]:
            with st.spinner("Generating AI summary for this event..."):
                try:
                    summary_text = build_event_summary(selected_row, company_for_cache, resources_text_for_cache)
                except Exception as ex:
                    summary_text = f"(Summary error: {ex})"
                st.session_state["summaries"][event_id] = summary_text

# ---------------------------
# Step 3: generate Slides + send Slack with selected resources
# ---------------------------
if rows:
    selected = rows[idx]["_raw"]
    company = attendees_to_company(
        selected.get("attendees", []),
        event_summary=selected.get("summary", ""),
        event_description=selected.get("description", ""),
    )
    label = f"{company['companyName']}" + (f" ({company['companyDomain']})" if company.get("companyDomain") else "")
    st.write(f"**Company guest:** {label}")

    # Always show the cached/most recent summary for the selected row
    current_event_id = st.session_state.get("current_event_id")
    current_summary = st.session_state["summaries"].get(current_event_id, "") if current_event_id else ""
    st.text_area("AI Summary (auto-generated on row change)", current_summary, height=180)

    # Button appears AFTER events are loaded and checkboxes are visible
    if st.button("Generate POC Document"):
        if not GSLIDES_TEMPLATE_ID:
            st.error("Missing GSLIDES_TEMPLATE_ID in .env")
        else:
            try:
                creds = get_google_creds(GOOGLE_OAUTH_CLIENT_JSON)

                name = f"{company['companyName']} - POC Brief ({rows[idx]['start']})"
                file_id = copy_template(creds, GSLIDES_TEMPLATE_ID, name)

                # Build resources text for Slides and Slack
                resources_text = format_resources_text(selected_resources)

                # Replace placeholders in Slides (add {{Resources}} to your template)
                replacements = {
                    "{{CompanyName}}": company['companyName'],
                    "{{POCEmails}}": ", ".join(collect_event_emails(selected)),
                    "{{EventTitle}}": rows[idx]['summary'],
                    "{{EventTime}}": rows[idx]['start'],
                    "{{Resources}}": resources_text or "—",
                }
                fill_slides_placeholders(creds, file_id, replacements)
                link = get_web_view_link(creds, file_id)

                # Use cached summary if available; (optional) regenerate to include link
                event_id = rows[idx]["id"]
                summary = st.session_state["summaries"].get(event_id, "")
                # If you prefer to refresh after the Slides link exists, uncomment:
                # summary = build_event_summary(rows[idx], company, resources_text)
                # st.session_state["summaries"][event_id] = summary

                st.success("Slides brief generated!")
                st.markdown(f"**Slides:** {link}")
                st.text_area("Summary (final)", summary, height=180)

                # Slack (single, centralized call that respects toggle)
                sent = post_to_slack_if_enabled(
                    summary=summary,
                    link=link,
                    company_name=company["companyName"],
                    event_title=rows[idx]["summary"],
                    event_time=rows[idx]["start"],
                    poc_emails=replacements["{{POCEmails}}"],
                    resources_text=resources_text,
                    slack_webhook=SLACK_WEBHOOK_URL
                )
                if sent:
                    st.info("Sent to Slack ✅")
                else:
                    if st.session_state.get("auto_slack") and SLACK_WEBHOOK_URL:
                        st.warning("Could not send to Slack (check webhook URL/permissions).")

            except Exception as ex:
                st.error(f"Error generating slides brief: {ex}")
