import streamlit as st
from docx import Document
from urllib.parse import unquote
from datetime import date
import hashlib
import requests
import io
import os

# --- Password protection ---
def verify_password(pw: str) -> bool:
    return hashlib.sha256(pw.encode()).hexdigest() == st.secrets.get("password_hash", "")

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

# Cache query parameters before rerun
if "query_params" not in st.session_state:
    st.session_state.query_params = dict(st.query_params)

if not st.session_state.authenticated:
    pw = st.text_input("Enter password", type="password")
    if pw and verify_password(pw):
        st.session_state.authenticated = True
        st.rerun()
    elif pw:
        st.error("Incorrect password.")
    st.stop()

# --- Authenticated section ---
st.title("📄 Format Your Recommendation Letter")

params = st.session_state.query_params
letter_text = unquote(params.get("text", ""))
paste_id = params.get("paste_id", "")

# If no direct text provided, try to load from pastebin
if paste_id and not letter_text:
    try:
        paste_url = f"https://pastebin.com/raw/{paste_id}"
        resp = requests.get(paste_url)
        if resp.status_code == 200:
            letter_text = resp.text
        else:
            st.warning("Failed to load letter text from Pastebin.")
    except Exception as e:
        st.error("Error loading letter from Pastebin.")

addressee = unquote(params.get("addressee", ""))
salutation = unquote(params.get("salutation", ""))
letter_date = unquote(params.get("date", date.today().strftime("%B %d, %Y")))

filename = st.text_input("Enter filename (without extension)", value="recommendation_letter")

# Use local template from folder
template_path = os.path.join(os.path.dirname(__file__), "Shah_LOS_template.docx")

if os.path.exists(template_path) and letter_text and addressee and salutation:
    template = Document(template_path)

    def replace(doc, replacements):
        for p in doc.paragraphs:
            for key, val in replacements.items():
                if key in p.text:
                    for run in p.runs:
                        run.text = run.text.replace(key, val)
        return doc

    replacements = {
        "<<Date>>": letter_date,
        "<<Addressee>>": addressee,
        "<<Salutation>>": salutation,
        "<<Enter text here>>": letter_text
    }

    updated_doc = replace(template, replacements)

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
