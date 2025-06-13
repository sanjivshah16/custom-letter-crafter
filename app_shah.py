import streamlit as st
import hashlib
from datetime import date
import io
from openai import OpenAI
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
import pandas as pd
import fitz
from twilio.rest import Client

DEFAULT_TEMPLATE = "Shah_LOS_template.docx"

def notify_usage():
    try:
        twilio_sid = st.secrets["twilio"]["account_sid"]
        twilio_token = st.secrets["twilio"]["auth_token"]
        client = Client(twilio_sid, twilio_token)
        client.messages.create(
            body="📄 Letter Crafter was just used to generate a new letter.",
            from_=st.secrets["twilio"]["from_number"],
            to=st.secrets["twilio"]["to_number"]
        )
    except Exception as e:
        st.warning(f"(SMS failed: {e})")

st.set_page_config(page_title="Custom Letter Crafter", layout="wide")
st.title("📄 Custom Letter Crafter for Sanjiv J. Shah, MD")

def verify_password(pw: str) -> bool:
    return hashlib.sha256(pw.encode()).hexdigest() == st.secrets.get("password_hash", "")

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    pw = st.text_input("Enter password", type="password")
    if pw and verify_password(pw):
        st.session_state.authenticated = True
        st.rerun()
    elif pw:
        st.error("Incorrect password.")
    st.stop()

client = OpenAI(api_key=st.secrets["openai_api_key"])

# --- File extractors ---
def extract_text_from_pdf(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    return "\n".join([page.get_text() for page in doc])

def extract_text_from_docx(file):
    doc = Document(file)
    return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])

def extract_text_from_xlsx(file):
    dfs = pd.read_excel(file, sheet_name=None)
    output = []
    for name, df in dfs.items():
        output.append(df.head(20).to_string(index=False))
    return "\n".join(output)

def prepare_file_context(files):
    previews = []
    for f in files:
        filename = f.name
        if filename.endswith(".pdf"):
            text = extract_text_from_pdf(f)
        elif filename.endswith(".docx"):
            text = extract_text_from_docx(f)
        elif filename.endswith(".xlsx"):
            text = extract_text_from_xlsx(f)
        else:
            text = f.read().decode(errors="ignore")
        previews.append(f"{filename}:\n{text[:2000]}\n")
    return "\n".join(previews)

# Inputs
uploaded_files = st.file_uploader("Upload CVs, drafts, personal statements, etc.", accept_multiple_files=True)
relationship_text = st.text_area("How do you know the applicant? (1–2 sentences)", height=120)

addressee = st.text_input("Addressee", "")
salutation = st.text_input("Salutation", "To Whom It May Concern:")
letter_date = date.today().strftime("%B %d, %Y")
filename = st.text_input("Output filename (no extension)", value="recommendation_letter")

font_name = st.selectbox("Font", ["Arial", "Times New Roman", "Calibri", "Aptos"], index=0)
font_size = st.selectbox("Font size", [9, 10, 10.5, 11, 11.5, 12], index=3)

# Generate letter
def generate_letter():
    file_context = prepare_file_context(uploaded_files)
    prompt = f"Relationship: {relationship_text}\nFiles:\n{file_context}\nWrite the letter body only."
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.7,
        max_tokens=1000
    )
    return response.choices[0].message.content.strip()

if st.button("✍️ Generate Letter"):
    if not uploaded_files or not relationship_text.strip():
        st.warning("Please upload files and describe your relationship.")
        st.stop()
    letter_body = generate_letter()
    st.session_state.letter_text = letter_body
    notify_usage()
    st.success("Letter body generated.")

# Template insertion
def replace_placeholders(doc, replacements):
    for p in doc.paragraphs:
        for placeholder, replacement in replacements.items():
            if placeholder in p.text:
                p.text = p.text.replace(placeholder, replacement)

if "letter_text" in st.session_state:
    doc = Document(DEFAULT_TEMPLATE)
    replacements = {
        "<<Date>>": letter_date,
        "<<Addressee>>": addressee,
        "<<Salutation>>": salutation,
        "<<Enter text here>>": st.session_state.letter_text
    }
    replace_placeholders(doc, replacements)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    st.download_button(
        label="Download Letter (DOCX)",
        data=buffer,
        file_name=f"{filename}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
