import streamlit as st
from docx import Document
from urllib.parse import unquote
from datetime import date
import io
import hashlib
import os

# Load password hash from secrets
PASSWORD_HASH = st.secrets.get("password_hash", "")

def verify_password(pw: str) -> bool:
    return hashlib.sha256(pw.encode()).hexdigest() == PASSWORD_HASH

# Prompt for password
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    pw = st.text_input("Enter password", type="password")
    if pw and verify_password(pw):
        st.session_state.authenticated = True
        st.experimental_rerun()
    elif pw:
        st.error("Incorrect password.")
    st.stop()

# Authenticated section
st.title("📄 Format Your Recommendation Letter")

# Get query parameters
params = st.query_params
letter_text = unquote(params.get("text", ""))
addressee = unquote(params.get("addressee", ""))
salutation = unquote(params.get("salutation", ""))
letter_date = unquote(params.get("date", date.today().strftime("%B %d, %Y")))

# Let user optionally rename file
filename = st.text_input("Enter filename (without extension)", value="recommendation_letter")

# Use local template instead of upload
template_path = os.path.join(os.path.dirname(__file__), "Shah_LOS_template.docx")
template = Document(template_path)

# Replace placeholders
def replace(doc, replacements):
    for p in doc.paragraphs:
        for key, val in replacements.items():
            if key in p.text:
                for run in p.runs:
                    run.text = run.text.replace(key, val)
    return doc

if letter_text and addressee and salutation:
    replacements = {
        "<<Date>>": letter_date,
        "<<Addressee>>": addressee,
        "<<Salutation>>": salutation,
        "<<Enter text here>>": letter_text
    }

    updated_doc = replace(template, replacements)

    # Save DOCX
    docx_buffer = io.BytesIO()
    updated_doc.save(docx_buffer)
    docx_buffer.seek(0)

    st.download_button(
        label="📥 Download DOCX",
        data=docx_buffer,
        file_name=f"{filename}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    st.info("To convert the DOCX file to PDF, please use an external tool or service.")
else:
    st.info("Awaiting letter text, addressee, and salutation.")
