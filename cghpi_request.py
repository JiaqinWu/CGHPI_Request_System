import streamlit as st
import re
import pandas as pd
from millify import millify 
from streamlit_extras.metric_cards import style_metric_cards
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import altair as alt
import json
import time
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import io
from mailjet_rest import Client
import openpyxl
from openpyxl.styles import PatternFill
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from PIL import Image as PILImage, ImageDraw, ImageFont
import base64
import urllib.request

st.set_page_config(
        page_title="CGHPI Request System",
        page_icon="https://raw.githubusercontent.com/JiaqinWu/HRSA64_Dash/main/Georgetown_logo_blueRGB.png", 
        layout="centered"
    ) 

scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
#creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
# Use Streamlit's secrets management
creds_dict = st.secrets["gcp_service_account"]
# Extract individual attributes needed for ServiceAccountCredentials
credentials = {
    "type": creds_dict.type,
    "project_id": creds_dict.project_id,
    "private_key_id": creds_dict.private_key_id,
    "private_key": creds_dict.private_key,
    "client_email": creds_dict.client_email,
    "client_id": creds_dict.client_id,
    "auth_uri": creds_dict.auth_uri,
    "token_uri": creds_dict.token_uri,
    "auth_provider_x509_cert_url": creds_dict.auth_provider_x509_cert_url,
    "client_x509_cert_url": creds_dict.client_x509_cert_url,
}

# Create JSON string for credentials
creds_json = json.dumps(credentials)

# Load credentials and authorize gspread
creds = ServiceAccountCredentials.from_json_keyfile_dict(json.loads(creds_json), scope)
client = gspread.authorize(creds)


def send_email_mailjet(to_email, subject, body):
    """Send email via Mailjet and return success status"""
    api_key = st.secrets["mailjet"]["api_key"]
    api_secret = st.secrets["mailjet"]["api_secret"]
    sender = st.secrets["mailjet"]["sender"]

    mailjet = Client(auth=(api_key, api_secret), version='v3.1')

    data = {
        'Messages': [
            {
                "From": {
                    "Email": sender,
                    "Name": "CGHPI Request System"
                },
                "To": [
                    {
                        "Email": to_email,
                        "Name": to_email.split("@")[0]
                    }
                ],
                "Subject": subject,
                "TextPart": body
            }
        ]
    }

    try:
        result = mailjet.send.create(data=data)
        if result.status_code == 200:
            return True
        else:
            st.warning(f"❌ Failed to email {to_email}: Status {result.status_code}")
            return False
    except Exception as e:
        st.error(f"❗ Mailjet error: {e}")
        return False


def upload_file_to_drive(file, filename, folder_id, creds_dict):
    """Upload a Streamlit-uploaded file object to Google Drive and return a shareable link."""
    # Convert Streamlit secret dict into Google Credentials object
    drive_creds = Credentials.from_service_account_info(creds_dict, scopes=[
        "https://www.googleapis.com/auth/drive"
    ])

    # Build the Drive API service
    drive_service = build('drive', 'v3', credentials=drive_creds)

    # Prepare file metadata
    file_metadata = {
        'name': filename,
        'parents': [folder_id]
    }

    # Read uploaded file and prepare media
    media = MediaIoBaseUpload(io.BytesIO(file.read()), mimetype=file.type)

    # Upload file
    uploaded = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id'
    ).execute()

    # Make the file public
    drive_service.permissions().create(
        fileId=uploaded['id'],
        body={'type': 'anyone', 'role': 'reader'}
    ).execute()

    # Return the sharable link
    return f"https://drive.google.com/file/d/{uploaded['id']}/view"


def upload_bytes_to_drive(file_bytes, filename, mimetype, folder_id, creds_dict):
    """Upload raw bytes (e.g., generated PDF) to Google Drive and return a shareable link."""
    drive_creds = Credentials.from_service_account_info(creds_dict, scopes=[
        "https://www.googleapis.com/auth/drive"
    ])
    drive_service = build('drive', 'v3', credentials=drive_creds)

    file_metadata = {
        'name': filename,
        'parents': [folder_id]
    }

    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mimetype)

    uploaded = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id'
    ).execute()

    drive_service.permissions().create(
        fileId=uploaded['id'],
        body={'type': 'anyone', 'role': 'reader'}
    ).execute()

    return f"https://drive.google.com/file/d/{uploaded['id']}/view"


def generate_request_pdf(request_dict):
    """Generate a standardized PDF summary for a communications request and return its bytes."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=0.5 * inch,
        leftMargin=0.5 * inch,
        topMargin=0.5 * inch,
        bottomMargin=0.5 * inch,
    )

    styles = getSampleStyleSheet()
    story = []

    title_style = ParagraphStyle(
        "RequestTitle",
        parent=styles["Heading1"],
        fontSize=16,
        textColor=colors.HexColor("#000000"),
        alignment=1,
        spaceAfter=12,
    )
    section_title_style = ParagraphStyle(
        "SectionTitle",
        parent=styles["Heading2"],
        fontSize=12,
        textColor=colors.HexColor("#000000"),
        spaceBefore=6,
        spaceAfter=4,
    )

    def fmt(value):
        if isinstance(value, list):
            return ", ".join(str(v) for v in value)
        if value is None:
            return ""
        return str(value)

    ticket_id = fmt(request_dict.get("Ticket ID"))
    submit_date = fmt(request_dict.get("Submit Date"))

    # Title
    story.append(Paragraph("CGHPI Communications Request Summary", title_style))
    story.append(Spacer(1, 0.15 * inch))

    # Basic info table
    basic_data = [
        ["Ticket ID", ticket_id, "Submit Date", submit_date],
        ["Requestor Name", fmt(request_dict.get("Name")), "Email Address", fmt(request_dict.get("Email Address"))],
        ["Project/Grant", fmt(request_dict.get("Project/Grant")), "Request Type", fmt(request_dict.get("Request Type"))],
    ]
    basic_table = Table(basic_data, colWidths=[1.7 * inch, 2.3 * inch, 1.7 * inch, 2.3 * inch])
    basic_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f1f1f1")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
            ]
        )
    )
    story.append(basic_table)
    story.append(Spacer(1, 0.15 * inch))

    # Section: Request Details
    story.append(Paragraph("Request Details", section_title_style))
    details_data = [
        ["Type of Support Needed", fmt(request_dict.get("Type of Support Needed"))],
        ["Primary Purpose", fmt(request_dict.get("Primary Purpose"))],
        ["Target Audience", fmt(request_dict.get("Target Audience"))],
        ["Audience Action", fmt(request_dict.get("Audience Action"))],
        ["Key Points to Include", fmt(request_dict.get("Key Points"))],
        ["Subject Matter Expectations", fmt(request_dict.get("Subject Matter"))],
    ]
    details_table = Table(details_data, colWidths=[2.2 * inch, 6.0 * inch])
    details_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
            ]
        )
    )
    story.append(details_table)
    story.append(Spacer(1, 0.15 * inch))

    # Section: Timeline & Priority
    story.append(Paragraph("Timeline & Priority", section_title_style))
    timeline_data = [
        ["Requested Due Date", fmt(request_dict.get("Requested Due Date")), "Priority Level", fmt(request_dict.get("Priority Level"))],
        ["Driver of Deadline", fmt(request_dict.get("Driver Deadline")), "Tied to Grant Deliverable", fmt(request_dict.get("Tie to Grant Deliverable"))],
    ]
    timeline_table = Table(timeline_data, colWidths=[1.9 * inch, 2.1 * inch, 1.9 * inch, 2.1 * inch])
    timeline_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
            ]
        )
    )
    story.append(timeline_table)
    story.append(Spacer(1, 0.15 * inch))

    # Section: Publishing & Compliance
    story.append(Paragraph("Publishing & Compliance", section_title_style))
    pub_data = [
        ["Will this be shared externally?", fmt(request_dict.get("Share Externally"))],
        ["Includes any sensitive information", fmt(request_dict.get("Information Include"))],
        ["Permissions for photos/quotes secured", fmt(request_dict.get("Permission Secure"))],
    ]
    pub_table = Table(pub_data, colWidths=[2.8 * inch, 5.4 * inch])
    pub_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
            ]
        )
    )
    story.append(pub_table)
    story.append(Spacer(1, 0.15 * inch))

    # Section: Scope & Format
    story.append(Paragraph("Scope & Format", section_title_style))
    scope_data = [
        ["Estimated Length/Size", fmt(request_dict.get("Estimated Length"))],
        ["Level of Design Support Needed", fmt(request_dict.get("Level of Design Support"))],
        ["Where it will live", fmt(request_dict.get("Live"))],
    ]
    scope_table = Table(scope_data, colWidths=[2.8 * inch, 5.4 * inch])
    scope_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
            ]
        )
    )
    story.append(scope_table)
    story.append(Spacer(1, 0.2 * inch))

    # Footer
    generated_on = datetime.today().strftime("%Y-%m-%d %H:%M")
    footer_text = f"Generated on {generated_on} | CGHPI Communications Request System"
    story.append(Paragraph(footer_text, styles["Normal"]))

    doc.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes


def _get_records_with_retry(spreadsheet_name, worksheet_name, retries=3, base_delay=0.5):
    """Fetch worksheet records with simple exponential backoff to mitigate 429s."""
    attempt = 0
    last_exc = None
    while attempt < retries:
        try:
            spreadsheet = client.open(spreadsheet_name)
            worksheet = spreadsheet.worksheet(worksheet_name)
            return worksheet.get_all_records()
        except Exception as exc:
            last_exc = exc
            delay = base_delay * (2 ** attempt)
            time.sleep(delay)
            attempt += 1
    # If all retries failed, re-raise last exception
    raise last_exc

@st.cache_data(ttl=600)
def load_communication_sheet():
    df = pd.DataFrame(_get_records_with_retry('CGHPI_Request_System', 'Communication'))

    # Normalize/construct a 'Submit Date' column robustly, since legacy sheets may use different headers.
    submit_date_col = None
    for candidate in ["Submit Date", "Submit date", "submit_date", "SubmitDate"]:
        if candidate in df.columns:
            submit_date_col = candidate
            break

    if submit_date_col:
        df["Submit Date"] = pd.to_datetime(df[submit_date_col], errors="coerce")
    else:
        # Fallback: create an empty/NaT submit date column so downstream code works
        df["Submit Date"] = pd.NaT

    return df

df = load_communication_sheet()


# --- Demo user database
USERS = {
    'jw2104@georgetown.edu':{
        "Coordinator": {"password": "Jiaqin123!", "name": "Jiaqin Wu"}
    },
    'ew898@georgetown.edu':{
        "Coordinator": {"password": "Eric123!", "name": "Eric Wagner"}
    }
}

# --- Initialize session state
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "role" not in st.session_state:
    st.session_state.role = None
if "user_email" not in st.session_state:
    st.session_state.user_email = ""

# --- Role selection
if st.session_state.role is None:
    #st.image("Georgetown_logo_blueRGB.png",width=200)
    #st.title("Welcome to the GU Technical Assistance Provider System")
    st.markdown(
        """
        <div style='
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            background: #f8f9fa;
            padding: 2em 0 1em 0;
            border-radius: 18px;
            box-shadow: 0 4px 24px rgba(0,0,0,0.07);
            margin-bottom: 2em;
        '>
            <img src='https://raw.githubusercontent.com/JiaqinWu/HRSA64_Dash/main/Georgetown_logo_blueRGB.png' width='200' style='margin-bottom: 1em;'/>
            <h1 style='
                color: #1a237e;
                font-family: "Segoe UI", "Arial", sans-serif;
                font-weight: 700;
                margin: 0;
                font-size: 2.2em;
                text-align: center;
            '>Welcome to the CGHPI Request System</h1>
        </div>
        """,
        unsafe_allow_html=True
    )

    role = st.selectbox(
        "Select your role",
        ["Requester", "Coordinator"],
        index=None,
        placeholder="Select option..."
    )

    if role:
        st.session_state.role = role
        st.rerun()

# --- Show view based on role
else:
    st.sidebar.markdown(
        f"""
        <div style='
            background: #f8f9fa;
            border-radius: 12px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
            padding: 1.2em 1em 1em 1em;
            margin-bottom: 1.5em;
            text-align: center;
            font-family: Arial, "Segoe UI", sans-serif;
        '>
            <span style='
                font-size: 1.15em;
                font-weight: 700;
                color: #1a237e;
                letter-spacing: 0.5px;
            '>
                Role: {st.session_state.role}
            </span>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.sidebar.button("🔄 Switch Role", on_click=lambda: st.session_state.update({
        "authenticated": False,
        "role": None,
        "user_email": ""
    }))

    # Sidebar: refresh cached datasets
    if st.sidebar.button("🔁 Refresh Data"):
        st.cache_data.clear()
        st.rerun()

    st.markdown("""
        <style>
        .stButton > button {
            width: 100%;
            background-color: #cdb4db;
            color: black;
            font-family: Arial, "Segoe UI", sans-serif;
            font-weight: 600;
            border-radius: 8px;
            padding: 0.6em;
            margin-top: 1em;
            transition: background 0.2s;
        }
        .stButton > button:hover {
            background-color: #b197fc;
            color: #222;
        }
        </style>
    """, unsafe_allow_html=True)

    # Requester: No login needed
    if st.session_state.role == "Requester":
        st.markdown(
            """
            <div style='
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
                background: #f8f9fa;
                padding: 2em 0 1em 0;
                border-radius: 18px;
                box-shadow: 0 4px 24px rgba(0,0,0,0.07);
                margin-bottom: 2em;
            '>
                <img src='https://raw.githubusercontent.com/JiaqinWu/HRSA64_Dash/main/Georgetown_logo_blueRGB.png' width='200' style='margin-bottom: 1em;'/>
                <h1 style='
                    color: #1a237e;
                    font-family: "Segoe UI", "Arial", sans-serif;
                    font-weight: 700;
                    margin: 0;
                    font-size: 2.2em;
                    text-align: center;
                '>📥 Communications Request Form</h1>
            </div>
            """,
            unsafe_allow_html=True
        )
        # Extract last Ticket ID from the existing sheet (robust to empty/no-column cases)
        if "Ticket ID" in df.columns and not df["Ticket ID"].dropna().empty:
            last_ticket = (
                df["Ticket ID"]
                .dropna()
                .astype(str)
                .str.extract(r"GU(\d+)", expand=False)
                .astype(float)
                .max()
            )
            next_ticket_number = 1 if pd.isna(last_ticket) else int(last_ticket) + 1
        else:
            next_ticket_number = 1
        new_ticket_id = f"GU{next_ticket_number:04d}"
        st.markdown("FORM DESCRIPTION")
        st.write("Please complete this form for all communications-related requests. Providing detailed information helps ensure your request is prioritized appropriately and completed on time. We will review your request and will be in touch within 1-2 business days.")
        # Requester form
        st.markdown("SECTION 1: About the Request")
        col1, col2 = st.columns(2)
        with col1:
            name = st.text_input("Requestor Name *",placeholder="Enter text", key="requester_name")
        with col2:
            email = st.text_input("Email Address *",placeholder='Enter text', key="email_address")
        col3, col4 = st.columns(2)
        with col3:
            project = st.selectbox(
                "Project/Grant *",
                ['HRSA (EHE/GU-TAP)','Country Program','Cross-Center','Other'],
                index=None,
                placeholder="Select option...",
                key="project_grant"
            )
            # If "Other" is selected, show a text input for custom value
            if project == "Other":
                project_other = st.text_input("Please specify the project/grant *", key="project_grant_other")
                if project_other:
                    project = project_other 
        with col4:
            request_type = st.selectbox(
                "Request Type *",
                ['New Product','Update/Revision','Event Support','Dissemination Support','Thought Leadership','Other'],
                index=None,
                placeholder="Select option...",
                key="request_type"
            )
            if request_type == "Other":
                request_type_other = st.text_input("Please specify the request type *", key="request_type_other")
                if request_type_other:
                    request_type = request_type_other 
        col5, col6 = st.columns(2)
        with col5:
            type_support = st.multiselect(
                "Type of Support Needed (Check all that apply) *",
                [
                    'Copyediting',
                    'Web – landing page updates',
                    'Social media (site visit recap, case study, spotlight, update)',
                    'One-pager / Factsheet',
                    'Visuals (infographic, data visual, photos)',
                    'Video (explainer, toolkit module, interview)',
                    'Presentation support (slides)',
                    'Event support (webinar/workshop photography, recap)',
                    'Dissemination templates (slide deck, checklist, toolkit template)',
                    'Other',
                ],
                placeholder="Select option...",
                key="type_support"
            )
            if "Other" in type_support:
                type_support_other = st.text_input("Please specify the type of support needed *", key="type_support_other")
                if type_support_other:
                    # Replace "Other" with custom value but keep any other selected options
                    cleaned = [opt for opt in type_support if opt != "Other"]
                    cleaned.append(type_support_other)
                    type_support = cleaned
        with col6:
            primary_purpose = st.multiselect(
                "Primary Purpose (Check all that apply) *",
                ['Inform','Share learning','Promote an event/resource','Support TA delivery','Support reporting','Thought leadership','Other'],
                placeholder="Select option...",
                key="primary_purpose"
            )
            if "Other" in primary_purpose:
                primary_purpose_other = st.text_input("Please specify the type of support needed *", key="primary_purpose_other")
                if primary_purpose_other:
                    cleaned = [opt for opt in primary_purpose if opt != "Other"]
                    cleaned.append(primary_purpose_other)
                    primary_purpose = cleaned
        
        target_audience = st.multiselect(
            "Target Audience (Check all that apply) *",
            ['HRSA','Jurisdictions','National HIV audience','Global HIV audience','Internal team','Conference attendees','General public','Other'],
            placeholder="Select option...",
            key="target_audience"
        )
        if "Other" in target_audience:
            target_audience_other = st.text_input("Please specify the target audience *", key="target_audience_other")
            if target_audience_other:
                cleaned = [opt for opt in target_audience if opt != "Other"]
                cleaned.append(target_audience_other)
                target_audience = cleaned

        audience_action = st.text_area("What should the audience do after seeing this? *", placeholder="Example: Download toolkit, register for webinar, apply a practice, contact team, etc.", height=150, key="audience_action")

        st.markdown("SECTION 2: Timeline & Priority")
        due_date = st.date_input(
            "Requested Due Date *",
            value=None,
            key="requested_due_date"
        )
        driver_deadline = st.text_input("What is driving this deadline? *", placeholder="Event, grant deliverable, conference, leadership request, etc.", key="driver_deadline")
        col7, col8 = st.columns(2)
        with col7:
            tie_grant_deliverable = st.selectbox(
                "Is this tied to a grant deliverable? *",
                ['Yes','No','Unsure'],
                index=None,
                placeholder="Select option...",
                key="tie_grant_deliverable"
            )

        with col8:
            priority_level = st.selectbox(
                "Priority Level *",
                ["High (external deadline confirmed)","Medium (target date but flexible)", "Flexible"],
                index=None,
                placeholder="Select option...",
                key="priority_level"
            )
        st.markdown("SECTION 3: Content Inputs")
        background_share = st.file_uploader(
            "Upload any background materials if available.",accept_multiple_files=True, key="background_share"
        )
        draft_copy = st.file_uploader(
            "Upload draft copy if available.",accept_multiple_files=True, key="draft_copy"
        )

        key_points = st.text_area("Key Points to Include *", placeholder="Please provide bullet points, messages, or required language.", height=150, key="key_points")
        subject_matter = st.text_area("Subject Matter Expectations (if different from requestor) *", placeholder="Enter text here", height=150, key="subject_matter")

        st.markdown("SECTION 4: Publishing & Compliance")
        share_external = st.selectbox("Will this be shared externally? *", ["Yes","Internal only","Unsure"], index=None, placeholder="Select option...", key="share_external")
        information_include = st.multiselect(
            "Does this include any of the following? (Check all that apply) *",
            ['Identifiable jurisdiction information','Performance data or targets','Photos of people','Direct quotes','None of the above'],
            placeholder="Select option...",
            key="information_include"
        )
        permission_secure = st.selectbox("If photos or quotes are included, have permissions been secured? *", ["Yes","No","Not applicable"], index=None, placeholder="Select option...", key="permission_secure")
        st.markdown("SECTION 5: Scope & Format")
        estimated_length = st.selectbox(
                "Estimated Length/Size *",
                ['Short post (<300 words)','1-page product','2-5 pages','10+ pages','Multi-part series','Not surelive'],
                index=None,
                placeholder="Select option...",
                key="estimated_length"
            )

        level_of_design_support = st.selectbox(
                "Level of Design Support Needed *",
                ['Minimal formatting','Template-based design','Custom design','Video production'],
                index=None,
                placeholder="Select option...",
                key="level_of_design_support"
            )
        live = st.multiselect(
            "Where will this live? (Check all that apply) *",
            ['Website','LinkedIn','Email distribution','Webinar','Conference','Shared Drive (Box)','Other'],
            placeholder="Select option...",
            key="live"
        )
        live_other = None
        if "Other" in live:
            live_other = st.text_input("Please specify where it will live *", key="live_other")
            if live_other:
                # Append custom location but keep selected options
                if live_other not in live:
                    live = live + [live_other]
        
        submit_date = datetime.today().strftime("%Y-%m-%d")
        
        # Submit button
        st.markdown("""
            <style>
            .stButton > button {
                width: 100%;
                background-color: #cdb4db;
                color: black;
                font-family: Arial, "Segoe UI", sans-serif;
                font-weight: 600;
                border-radius: 8px;
                padding: 0.6em;
                margin-top: 1em;
            }
            </style>
        """, unsafe_allow_html=True)

        # Submit logic
        if st.button("Submit", key="requester_submit"):
            email_pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'

            errors = []


            drive_links = ""
            # Required field checks
            if not name: errors.append("Name is required.")
            if not email: errors.append("Email Address is required.")
            if not project: errors.append("Project/Grant is required.")
            if not request_type: errors.append("Request Type is required.")
            if not type_support: errors.append("Type of Support Needed is required.")
            if not primary_purpose: errors.append("Primary Purpose is required.")
            if not target_audience: errors.append("Target Audience is required.")
            if not audience_action: errors.append("Audience Action is required.")
            if not due_date: errors.append("Requested Due Date is required.")
            if not driver_deadline: errors.append("Driver Deadline is required.")
            if not tie_grant_deliverable: errors.append("Tie to Grant Deliverable is required.")
            if not priority_level: errors.append("Priority Level is required.")
            if not background_share: errors.append("Background Share is required.")
            if not draft_copy: errors.append("Draft Copy is required.")
            if not key_points: errors.append("Key Points is required.")
            if not subject_matter: errors.append("Subject Matter is required.")
            if not share_external: errors.append("Share Externally is required.")
            if not information_include: errors.append("Information Include is required.")
            if not permission_secure: errors.append("Permission Secure is required.")
            if not estimated_length: errors.append("Estimated Length is required.")
            if not live: errors.append("Live is required.")
            if "Other" in live and not live_other:
                errors.append("Please specify where it will live when 'Other' is selected.")

            # Show warnings or success
            if errors:
                for error in errors:
                    st.warning(error)
            else:
                # Only upload files if all validation passes
                if background_share:
                    try:
                        folder_id = "1BEDk2QLWKgfyG14VTqsx5w8EpmSo0tPP" 
                        links = []
                        upload_count = 0
                        for file in background_share:
                            # Rename file as: GU0001_filename.pdf
                            renamed_filename = f"{draft_copy}_{file.name}_{submit_date}"
                            link = upload_file_to_drive(
                                file=file,
                                filename=renamed_filename,
                                folder_id=folder_id,
                                creds_dict=st.secrets["gcp_service_account"]
                            )
                            links.append(link)
                            upload_count += 1
                            st.success(f"✅ Successfully uploaded: {file.name}")
                        drive_links = ", ".join(links)
                        if upload_count > 0:
                            st.success(f"✅ All {upload_count} file(s) uploaded successfully to Google Drive!")    
                    except Exception as e:
                        st.error(f"❌ Error uploading file(s) to Google Drive: {str(e)}")
                if draft_copy:
                    try:
                        folder_id = "18SdlxcJ9aPHCTy0N6lI-jjImGd6VlOKP" 
                        links = []
                        upload_count = 0
                        for file in draft_copy:
                            # Rename file as: GU0001_filename.pdf
                            renamed_filename = f"{name}_{file.name}_{submit_date}"
                            link = upload_file_to_drive(
                                file=file,
                                filename=renamed_filename,
                                folder_id=folder_id,
                                creds_dict=st.secrets["gcp_service_account"]
                            )
                            links.append(link)
                            upload_count += 1
                            st.success(f"✅ Successfully uploaded: {file.name}")
                        drive_links = ", ".join(links)
                        if upload_count > 0:
                            st.success(f"✅ All {upload_count} file(s) uploaded successfully to Google Drive!")    
                    except Exception as e:
                        st.error(f"❌ Error uploading file(s) to Google Drive: {str(e)}")

                # Generate a standardized PDF summary for this request and upload to Drive
                request_dict_for_pdf = {
                    'Ticket ID': new_ticket_id,
                    'Project/Grant': project,
                    'Name': name,
                    'Email Address': email,
                    'Request Type': request_type,
                    'Type of Support Needed': type_support,
                    'Primary Purpose': primary_purpose,
                    'Target Audience': target_audience,
                    'Audience Action': audience_action,
                    'Requested Due Date': due_date,
                    'Driver Deadline': driver_deadline,
                    'Tie to Grant Deliverable': tie_grant_deliverable,
                    'Priority Level': priority_level,
                    'Key Points': key_points,
                    'Subject Matter': subject_matter,
                    'Share Externally': share_external,
                    'Information Include': information_include,
                    'Permission Secure': permission_secure,
                    'Estimated Length': estimated_length,
                    'Level of Design Support': level_of_design_support,
                    'Live': live,
                    'Submit Date': submit_date,
                }
                pdf_link = ""
                try:
                    pdf_bytes = generate_request_pdf(request_dict_for_pdf)
                    pdf_filename = f"{new_ticket_id}_communications_request_summary.pdf"
                    pdf_folder_id = "1capjg8_dM314TuVp0bCpPG2ydb5U7jdf"
                    pdf_link = upload_bytes_to_drive(
                        file_bytes=pdf_bytes,
                        filename=pdf_filename,
                        mimetype="application/pdf",
                        folder_id=pdf_folder_id,
                        creds_dict=st.secrets["gcp_service_account"],
                    )
                except Exception as e:
                    st.error(f"❌ Error generating or uploading PDF summary: {str(e)}")

                new_row = {
                    'Ticket ID': new_ticket_id,
                    'Project/Grant': project,
                    'Name': name,
                    'Email Address': email,
                    'Request Type': request_type,
                    'Type of Support Needed': type_support,
                    'Primary Purpose': primary_purpose,
                    'Target Audience': target_audience,
                    'Audience Action': audience_action,
                    'Requested Due Date': due_date,
                    'Driver Deadline': driver_deadline,
                    'Tie to Grant Deliverable': tie_grant_deliverable,
                    'Priority Level': priority_level,
                    'Background Share': background_share,
                    'Draft Copy': draft_copy,
                    'Key Points': key_points,
                    'Subject Matter': subject_matter,
                    'Share Externally': share_external,
                    'Information Include': information_include,
                    'Permission Secure': permission_secure,
                    'Estimated Length': estimated_length,
                    'Level of Design Support': level_of_design_support,
                    'Live': live,
                    'Submit Date': submit_date,
                    'Request PDF Link': pdf_link,
                    # Workflow status fields
                    'Status': 'Submitted',
                    'Status Message': '',
                    'Output Links': '',
                    'Closed Date': pd.NA,
                }
                new_data = pd.DataFrame([new_row])

                try:
                    # Append new data to Google Sheet
                    updated_sheet = pd.concat([df, new_data], ignore_index=True)
                    updated_sheet = updated_sheet.applymap(
                        lambda x: x.strftime("%Y-%m-%d") if isinstance(x, (datetime, pd.Timestamp)) else x
                    )
                    # Replace NaN with empty strings to ensure JSON compatibility
                    updated_sheet = updated_sheet.fillna("")
                    # Keep Communications data in the same spreadsheet/worksheet used for reading
                    spreadsheet1 = client.open('CGHPI_Request_System')
                    worksheet1 = spreadsheet1.worksheet('Communication')
                    worksheet1.update([updated_sheet.columns.values.tolist()] + updated_sheet.values.tolist())
                    
                    # Clear cache to refresh data
                    st.cache_data.clear()
                    
                    # Send email notifications to coordinators (currently routed to Jiaqin Wu)
                    coordinator_emails = ["jw2104@georgetown.edu"]

                    subject = f"New Communications Request Submitted: {new_ticket_id}"
                    for coord_email in coordinator_emails:
                        coordinator_name = USERS.get(coord_email, {}).get("Coordinator", {}).get("name", "Coordinator")
                        personalized_body = f"""
                        Hi {coordinator_name},

                        A new communications request has been submitted in the CGHPI Request System.

                        Ticket ID: {new_ticket_id}
                        Requestor Name: {name}
                        Requestor Email: {email}
                        Project/Grant: {project}
                        Request Type: {request_type}
                        Type of Support Needed: {", ".join(type_support) if isinstance(type_support, list) else type_support}
                        Primary Purpose: {", ".join(primary_purpose) if isinstance(primary_purpose, list) else primary_purpose}
                        Target Audience: {", ".join(target_audience) if isinstance(target_audience, list) else target_audience}
                        Requested Due Date: {due_date.strftime("%Y-%m-%d")}
                        Priority Level: {priority_level}
                        Key Points: {key_points}
                        Share Externally: {share_external}
                        Estimated Length/Size: {estimated_length}
                        Where it will live: {", ".join(live) if isinstance(live, list) else live}
                        Attachments (Drive links): {drive_links or 'None'}
                        Request PDF: {pdf_link or 'Not available'}

                        Please log into the CGHPI Request System (https://cghpirequest.streamlit.app/) to review and manage this request.

                        Best,
                        CGHPI Request System
                        """
                        try:
                            send_email_mailjet(
                                to_email=coord_email,
                                subject=subject,
                                body=personalized_body,
                            )
                        except Exception as e:
                            st.warning(f"⚠️ Failed to send email to coordinator {coord_email}: {e}")


                    # Send confirmation email to requester
                    confirmation_subject = f"Your Communications Request ({new_ticket_id}) has been received"
                    confirmation_body = f"""
                    Hi {name},

                    Thank you for submitting your communications request to the CGHPI team.

                    Here is a summary of your submission:
                    - Ticket ID: {new_ticket_id}
                    - Project/Grant: {project}
                    - Request Type: {request_type}
                    - Type of Support Needed: {", ".join(type_support) if isinstance(type_support, list) else type_support}
                    - Primary Purpose: {", ".join(primary_purpose) if isinstance(primary_purpose, list) else primary_purpose}
                    - Target Audience: {", ".join(target_audience) if isinstance(target_audience, list) else target_audience}
                    - Requested Due Date: {due_date.strftime("%Y-%m-%d")}
                    - Priority Level: {priority_level}
                    - Key Points: {key_points}
                    - Share Externally: {share_external}
                    - Estimated Length/Size: {estimated_length}
                    - Where it will live: {", ".join(live) if isinstance(live, list) else live}

                    You can view a PDF summary of your request here: {pdf_link or 'Not available at this time'}.

                    A coordinator will review your request and follow up within 1–2 business days to confirm scope and timeline.
                    If you have any questions, please contact ew898@georgetown.edu.

                    Best,
                    CGHPI Request System
                    """

                    try:
                        send_email_mailjet(
                            to_email=email,
                            subject=confirmation_subject,
                            body=confirmation_body,
                        )
                    except Exception as e:
                        st.warning(f"⚠️ Failed to send confirmation email to requester: {e}")


                    st.success("✅ Thank you for your request. We will review and follow up within 1–2 business days to confirm scope and timeline.")
                    
                    # Clear cache to refresh data
                    st.cache_data.clear()
                    
                    # Clear form fields by using st.rerun() which will reset all form widgets
                    st.session_state.clear()
                    
                    # Wait a moment then redirect to main page
                    time.sleep(3)
                    st.rerun()

                except Exception as e:
                    st.error(f"Error updating Google Sheets: {str(e)}")




    # --- Coordinator or Staff: Require login
    elif st.session_state.role in ["Coordinator"]:
        if not st.session_state.authenticated:
            st.subheader("🔐 Login Required")

            email = st.text_input("Email")
            password = st.text_input("Password", type="password")
            login = st.button("Login")

            if login:
                # Normalize email to lowercase for case-insensitive lookup
                email_normalized = email.strip().lower() if email else ""
                user_roles = USERS.get(email_normalized)
                if user_roles and st.session_state.role in user_roles:
                    user = user_roles[st.session_state.role]
                    if user["password"] == password:
                        st.session_state.authenticated = True
                        st.session_state.user_email = email_normalized
                        st.success("Login successful!")
                        st.rerun()
                    else:
                        st.error("Invalid credentials or role mismatch.")
                else:
                    st.error("Invalid credentials or role mismatch.")

        else:
            if st.session_state.role == "Coordinator":
                user_info = USERS.get(st.session_state.user_email)
                coordinator_name = user_info["Coordinator"]["name"]
                # Check if current coordinator is Mabintou (only sees Travel Authorization Review Center)
                is_mabintou_coordinator = st.session_state.user_email == "mo887@georgetown.edu"
                st.markdown(
                    """
                    <div style='
                        display: flex;
                        flex-direction: column;
                        align-items: center;
                        justify-content: center;
                        background: #f8f9fa;
                        padding: 2em 0 1em 0;
                        border-radius: 18px;
                        box-shadow: 0 4px 24px rgba(0,0,0,0.07);
                        margin-bottom: 2em;
                    '>
                        <img src='https://raw.githubusercontent.com/JiaqinWu/HRSA64_Dash/main/Georgetown_logo_blueRGB.png' width='200' style='margin-bottom: 1em;'/>
                        <h1 style='
                            color: #1a237e;
                            font-family: "Segoe UI", "Arial", sans-serif;
                            font-weight: 700;
                            margin: 0;
                            font-size: 2.2em;
                            text-align: center;
                        '>📬 Coordinator Dashboard</h1>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
                #st.header("📬 Coordinator Dashboard")
                # Personalized greeting
                if user_info and "Coordinator" in user_info:
                    st.markdown(f"""
                    <div style='                      
                    background: #f8f9fa;                        
                    border-radius: 12px;                        
                    box-shadow: 0 2px 8px rgba(0,0,0,0.04);                        
                    padding: 1.2em 1em 1em 1em;                        
                    margin-bottom: 1.5em;                        
                    text-align: center;                        
                    font-family: Arial, "Segoe UI", sans-serif;                    
                    '>
                        <span style='                           
                        font-size: 1.15em;
                        font-weight: 700;
                        color: #1a237e;
                        letter-spacing: 0.5px;'>
                            👋 Welcome, {coordinator_name}!
                        </span>
                    </div>
                    """, unsafe_allow_html=True)
                
                # Dashboard Overview Metrics - visible to all coordinators (no header, just show metrics directly)
                col1, col2, col3 = st.columns(3)
                total_request = df['Ticket ID'].nunique()
                submitted_request = df[df['Status'] == 'Submitted']['Ticket ID'].nunique()
                inprogress_request = df[df['Status'] == 'In Progress']['Ticket ID'].nunique()

                col1.metric(label="# of Total Requests", value=millify(total_request, precision=2))
                col2.metric(label="# of Submitted Requests", value=millify(submitted_request, precision=2))
                col3.metric(label="# of In-Progress Requests", value=millify(inprogress_request, precision=2))
                style_metric_cards(border_left_color="#DBF227")

                col1, col2, col3 = st.columns(3)
                today = datetime.today()
                last_month = today - timedelta(days=30)
                declined_request = df[df['Status'] == 'Declined']['Ticket ID'].nunique()
                completed_request = df[df['Status'] == 'Completed']['Ticket ID'].nunique()
                recent_request = df[df['Submit Date'] >= last_month]['Ticket ID'].nunique()

                col1.metric(label="# of Declined Requests", value=millify(declined_request, precision=2))
                col2.metric(label="# of Completed Requests", value=millify(completed_request, precision=2))
                col3.metric(label="# of Requests (last 30 days)", value=millify(recent_request, precision=2))
                style_metric_cards(border_left_color="#DBF227")

                # ------------------------------
                # Request management and status workflow
                # ------------------------------

                # Ensure workflow-related columns exist
                management_df = df.copy()
                workflow_cols = [
                    'Status Message',
                    'Output Links',
                    'Closed Date',
                ]
                for col_name in workflow_cols:
                    if col_name not in management_df.columns:
                        management_df[col_name] = ""

                st.subheader("Request Management")

                status_options = ["All", "Submitted", "In Progress", "Declined", "Completed"]
                status_filter = st.selectbox(
                    "Filter by status",
                    status_options,
                    index=0,
                    key="status_filter"
                )

                if status_filter == "All":
                    filtered_df = management_df
                else:
                    filtered_df = management_df[management_df["Status"] == status_filter]

                if filtered_df.empty:
                    st.info("No requests found for the selected filter.")
                else:
                    # Sort by most recent submit date first (falling back to Ticket ID)
                    if "Submit Date" in filtered_df.columns:
                        filtered_df = filtered_df.sort_values("Submit Date", ascending=False)

                    for idx, row in filtered_df.iterrows():
                        ticket_id = row.get("Ticket ID", "") or "N/A"
                        requester_name = row.get("Name", "") or ""
                        current_status = row.get("Status", "Submitted") or "Submitted"

                        expander_label = f"{ticket_id} – {requester_name} ({current_status})"
                        with st.expander(expander_label, expanded=(current_status == "Submitted")):
                            # Basic request summary
                            st.markdown(f"**Project/Grant:** {row.get('Project/Grant', '')}")
                            st.markdown(f"**Request Type:** {row.get('Request Type', '')}")
                            st.markdown(f"**Type of Support Needed:** {row.get('Type of Support Needed', '')}")
                            st.markdown(f"**Primary Purpose:** {row.get('Primary Purpose', '')}")
                            st.markdown(f"**Target Audience:** {row.get('Target Audience', '')}")
                            st.markdown(f"**Requested Due Date:** {row.get('Requested Due Date', '')}")
                            st.markdown(f"**Priority Level:** {row.get('Priority Level', '')}")
                            st.markdown(f"**Key Points:** {row.get('Key Points', '')}")
                            st.markdown(f"**Share Externally:** {row.get('Share Externally', '')}")
                            st.markdown(f"**Where it will live:** {row.get('Live', '')}")
                            st.markdown(f"**Request PDF:** {row.get('Request PDF Link', '')}")

                            st.markdown("---")

                            # Status controls
                            status_choices = ["Submitted", "In Progress", "Declined", "Completed"]
                            current_status_index = status_choices.index(current_status) if current_status in status_choices else 0
                            new_status = st.selectbox(
                                "Status",
                                status_choices,
                                index=current_status_index,
                                key=f"status_{ticket_id}"
                            )

                            message_default = row.get("Status Message", "") or ""
                            status_message = st.text_area(
                                "Message to requester",
                                value=message_default,
                                height=120,
                                key=f"status_msg_{ticket_id}"
                            )

                            output_files = None
                            if new_status == "Completed":
                                output_files = st.file_uploader(
                                    "Attach output files (optional)",
                                    accept_multiple_files=True,
                                    key=f"outputs_{ticket_id}"
                                )

                            if st.button("Save updates", key=f"save_{ticket_id}"):
                                validation_errors = []
                                if new_status in ["In Progress", "Declined", "Completed"] and not status_message.strip():
                                    validation_errors.append("Please add a message for this status update.")

                                if validation_errors:
                                    for err in validation_errors:
                                        st.warning(err)
                                else:
                                    # Update the management_df row
                                    management_df.loc[idx, "Status"] = new_status
                                    management_df.loc[idx, "Status Message"] = status_message

                                    # Closed date logic
                                    if new_status == "Completed":
                                        management_df.loc[idx, "Closed Date"] = datetime.today().strftime("%Y-%m-%d")
                                    elif new_status in ["Submitted", "In Progress", "Declined"]:
                                        # Keep existing closed date if already set, otherwise blank
                                        if pd.isna(management_df.loc[idx, "Closed Date"]):
                                            management_df.loc[idx, "Closed Date"] = ""

                                    # Handle output file uploads when marking as completed
                                    existing_output_links = row.get("Output Links", "") or ""
                                    if new_status == "Completed" and output_files:
                                        try:
                                            folder_id = "188y2H9OkCaqTFSMShTFe6uK0oJihAup1"
                                            links = []
                                            upload_count = 0
                                            today_str = datetime.today().strftime("%Y-%m-%d")
                                            for file in output_files:
                                                renamed_filename = f"{ticket_id}_output_{file.name}_{today_str}"
                                                link = upload_file_to_drive(
                                                    file=file,
                                                    filename=renamed_filename,
                                                    folder_id=folder_id,
                                                    creds_dict=st.secrets["gcp_service_account"]
                                                )
                                                links.append(link)
                                                upload_count += 1
                                                st.success(f"✅ Successfully uploaded output file: {file.name}")
                                            all_links = []
                                            if existing_output_links:
                                                all_links.extend([l.strip() for l in existing_output_links.split(",") if l.strip()])
                                            all_links.extend(links)
                                            management_df.loc[idx, "Output Links"] = ", ".join(all_links)
                                            if upload_count > 0:
                                                st.success(f"✅ All {upload_count} output file(s) uploaded successfully to Google Drive!")
                                        except Exception as e:
                                            st.error(f"❌ Error uploading output file(s) to Google Drive: {str(e)}")
                                            # Preserve existing links if upload fails
                                            management_df.loc[idx, "Output Links"] = existing_output_links
                                    else:
                                        management_df.loc[idx, "Output Links"] = existing_output_links

                                    # Write updated management_df back to Google Sheet
                                    updated_sheet = management_df.applymap(
                                        lambda x: x.strftime("%Y-%m-%d") if isinstance(x, (datetime, pd.Timestamp)) else x
                                    )
                                    updated_sheet = updated_sheet.fillna("")

                                    try:
                                        spreadsheet2 = client.open('CGHPI_Request_System')
                                        worksheet2 = spreadsheet2.worksheet('Communication')
                                        worksheet2.update([updated_sheet.columns.values.tolist()] + updated_sheet.values.tolist())
                                    except Exception as e:
                                        st.error(f"Error updating Google Sheets with status change: {str(e)}")
                                    else:
                                        # Send email notification to requester about the status change
                                        requester_email = row.get("Email Address", "")
                                        requester_name = requester_name or "there"

                                        if requester_email:
                                            if new_status == "In Progress":
                                                subject = f"Your Communications Request ({ticket_id}) is now In Progress"
                                                body = f"""
                                                Hi {requester_name},

                                                Your communications request ({ticket_id}) has been reviewed and is now **In Progress**.

                                                Message from coordinator:
                                                {status_message}

                                                Best,
                                                CGHPI Request System
                                                """
                                            elif new_status == "Declined":
                                                subject = f"Your Communications Request ({ticket_id}) has been Declined"
                                                body = f"""
                                                Hi {requester_name},

                                                Your communications request ({ticket_id}) has been reviewed and **declined**.

                                                Message from coordinator:
                                                {status_message}

                                                If you have questions, please contact ew898@georgetown.edu.

                                                Best,
                                                CGHPI Request System
                                                """
                                            elif new_status == "Completed":
                                                subject = f"Your Communications Request ({ticket_id}) has been Completed"
                                                body = f"""
                                                Hi {requester_name},

                                                Your communications request ({ticket_id}) has been marked as **Completed**.

                                                Message from coordinator:
                                                {status_message}

                                                Output files (if any): {management_df.loc[idx, 'Output Links'] or 'None'}

                                                Best,
                                                CGHPI Request System
                                                """
                                            else:
                                                # For Submitted (rare to send) we skip sending an additional email
                                                subject = None
                                                body = None

                                            if subject and body:
                                                try:
                                                    send_email_mailjet(
                                                        to_email=requester_email,
                                                        subject=subject,
                                                        body=body,
                                                    )
                                                except Exception as e:
                                                    st.warning(f"⚠️ Failed to send status email to requester: {e}")

                                        st.success("Request updated successfully.")
                                        st.cache_data.clear()
                                        time.sleep(1)
                                        st.rerun()
