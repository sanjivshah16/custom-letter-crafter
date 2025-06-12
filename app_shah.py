import streamlit as st
import openai
import hashlib
from datetime import date
import io
import os
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

# --- Config ---
st.set_page_config(page_title="Letter Crafter", layout="wide")
st.title("🧠 Letter Crafter: Recommendation Letter Generator")

# --- Authentication ---
def verify_password(pw):
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

openai.api_key = st.secrets["openai_api_key"]

# --- Inputs ---
st.subheader("📁 Upload Documents")
uploaded_files = st.file_uploader("Upload CVs, drafts, reference guidelines, etc.", accept_multiple_files=True)

st.subheader("👥 Relationship Context")
relationship_text = st.text_area("Describe your relationship with the applicant (1-4 sentences)", height=100)

addressee = st.text_input("Addressee (e.g., Admissions Committee)", "")
salutation = st.text_input("Salutation (e.g., Dear Committee)", "")
if not salutation.strip():
    salutation = "To Whom It May Concern"

letter_date = date.today().strftime("%B %d, %Y")
filename = st.text_input("Filename for output", value="recommendation_letter")

if st.button("✍️ Generate Letter"):
    if not uploaded_files or not relationship_text:
        st.warning("Please upload at least one file and enter relationship details.")
        st.stop()

    try:
        # Create OpenAI tool messages
        system_prompt = (
            "You are Letter Crafter, a professional recommendation letter writer. "
            "The user will upload documents including CVs, draft letters, and reference guidelines. "
            "They also describe their relationship with the applicant. "
            "Your job is to write the body of a polished recommendation letter. "
            "DO NOT include the date, opening, or closing in the letter text. Just return the main body."
        )

        # Build file objects
        file_objs = []
        for f in uploaded_files:
            file_objs.append(("file", (f.name, f, f.type)))

        # Prepare message
        user_message = {
            "role": "user",
            "content": f"My relationship to the applicant: {relationship_text}\n\nPlease generate the letter body."
        }

        # Use OpenAI API with files
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": system_prompt},
                user_message
            ],
            files=uploaded_files,
            temperature=0.7,
            max_tokens=1000
        )

        letter_text = response["choices"][0]["message"]["content"].strip()
        st.session_state.letter_text = letter_text
        st.session_state.salutation = salutation
        st.session_state.addressee = addressee
        st.session_state.date = letter_date
        st.success("✅ Letter generated successfully.")
    except Exception as e:
        st.error(f"Error generating letter: {e}")
        st.stop()

# --- Template ---
if "letter_text" in st.session_state:
    template_path = os.path.join(os.path.dirname(__file__), "Shah_LOS_template.docx")
    if not os.path.exists(template_path):
        st.error("📁 Template file `Shah_LOS_template.docx` not found.")
    else:
        font_name = st.selectbox("Font", ["Arial", "Times New Roman", "Calibri", "Aptos"], index=0)
        font_size = st.selectbox("Font size", [9, 10, 10.5, 11, 11.5, 12], index=3)

        if st.button("📄 Format and Download Letter"):
            try:
                doc = Document(template_path)

                def replace_placeholders(doc, replacements):
                    for paragraph in doc.paragraphs:
                        for placeholder, replacement in replacements.items():
                            if placeholder in paragraph.text:
                                paragraph.text = paragraph.text.replace(placeholder, replacement)
                                for run in paragraph.runs:
                                    run.font.name = font_name
                                    run.font.size = Pt(font_size)
                                    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

                replacements = {
                    "<<Date>>": st.session_state.date,
                    "<<Addressee>>": st.session_state.addressee,
                    "<<Salutation>>": st.session_state.salutation,
                    "<<Enter text here>>": st.session_state.letter_text
                }

                replace_placeholders(doc, replacements)

                buffer = io.BytesIO()
                doc.save(buffer)
                buffer.seek(0)

                st.download_button(
                    label="📥 Download Letter (DOCX)",
                    data=buffer,
                    file_name=f"{filename}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Error formatting letter: {e}")
