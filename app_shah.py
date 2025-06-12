import streamlit as st
from docx import Document
from urllib.parse import unquote
from datetime import date
from docx.shared import Pt
from docx.oxml.ns import qn
import io
import requests
import hashlib
import json
import os

# --- Page config ---
st.set_page_config(page_title="Letter Formatter", layout="wide")
st.title("📄 Format Your Recommendation Letter")

# --- Password protection ---
def verify_password(pw: str) -> bool:
    return hashlib.sha256(pw.encode()).hexdigest() == st.secrets.get("password_hash", "")

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if "saved_query_params" not in st.session_state:
    st.session_state.saved_query_params = dict(st.query_params)

if not st.session_state.authenticated:
    pw = st.text_input("Enter password", type="password")
    if pw and verify_password(pw):
        st.session_state.authenticated = True
        st.rerun()
    elif pw:
        st.error("Incorrect password.")
    st.stop()

# --- Load query params after password ---
params = st.session_state.saved_query_params

# Load from pastebin or override
letter_text = ""
addressee = ""
salutation = ""
letter_date = date.today().strftime("%B %d, %Y")

if "paste_id" in params:
    try:
        paste_url = f"https://pastebin.com/raw/{params['paste_id']}"
        resp = requests.get(paste_url)
        if resp.status_code == 200:
            data = json.loads(resp.text)
            letter_text = data.get("text", "")
            addressee = data.get("addressee", "")
            salutation = data.get("salutation", "")
            letter_date = data.get("date", letter_date)
            st.success("✅ Letter data loaded from Pastebin.")
        else:
            st.error("Failed to load letter data from Pastebin.")
    except Exception as e:
        st.error(f"Error loading Pastebin: {e}")

# Override with individual query fields
letter_text = unquote(params.get("text", letter_text))
addressee = unquote(params.get("addressee", addressee))
salutation = unquote(params.get("salutation", salutation))
letter_date = unquote(params.get("date", letter_date))

# Show content preview
col1, col2 = st.columns([2, 1])
with col1:
    st.subheader("Parsed Content:")
    st.write(f"**Date:** {letter_date}")
    st.write(f"**Addressee:** {addressee or '(None provided)'}")
    st.write(f"**Salutation:** {salutation}")
    st.write(f"**Letter Text Length:** {len(letter_text)} characters")
    if letter_text:
        with st.expander("📖 Preview Letter Text"):
            st.write(letter_text)

with col2:
    st.subheader("🧾 Template Info")
    st.info("Using preloaded Word template: `Shah_LOS_template.docx`.\n\nIt must contain these placeholders:\n- `<<Date>>`\n- `<<Addressee>>`\n- `<<Salutation>>`\n- `<<Enter text here>>`")

# Font settings
filename = st.text_input("Enter filename (without extension)", value="recommendation_letter")
font_name = st.selectbox("Font", ["Arial", "Times New Roman", "Calibri", "Aptos"], index=0)
font_size = st.selectbox("Font size", [9, 10, 10.5, 11, 11.5, 12], index=3)

# Cache font settings
font_changed = (
    'last_font_name' in st.session_state and st.session_state.last_font_name != font_name
) or (
    'last_font_size' in st.session_state and st.session_state.last_font_size != font_size
)

if font_changed and 'processed_doc' in st.session_state:
    if st.button("🔄 Regenerate Letter with Updated Font Formatting"):
        del st.session_state.processed_doc
        del st.session_state.cache_key
        st.rerun()

# Main template path
template_path = os.path.join(os.path.dirname(__file__), "Shah_LOS_template.docx")

# Processing logic
if os.path.exists(template_path) and letter_text and salutation:
    try:
        cache_key = f"{hash(letter_text)}_{hash(addressee)}_{hash(salutation)}_{font_name}_{font_size}"
        if 'processed_doc' not in st.session_state or st.session_state.get('cache_key') != cache_key:
            template = Document(template_path)

            def replace_text_in_document(doc, replacements):
                replacements_made = {}
                paragraphs_to_remove = []
                letter_content_paragraph_index = None

                for i, paragraph in enumerate(doc.paragraphs):
                    original_text = paragraph.text

                    if not addressee and "<<Addressee>>" in original_text:
                        paragraphs_to_remove.append(i)
                        continue
                    if not addressee and i > 0 and "<<Addressee>>" in doc.paragraphs[i-1].text and original_text.strip() == "":
                        paragraphs_to_remove.append(i)
                        continue

                    for placeholder, replacement in replacements.items():
                        if placeholder in original_text:
                            paragraph.clear()
                            run = paragraph.add_run(original_text.replace(placeholder, replacement))
                            run.font.name = font_name
                            run.font.size = Pt(font_size)
                            run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
                            replacements_made[placeholder] = True

                            if placeholder == "<<Enter text here>>":
                                letter_content_paragraph_index = i
                            break

                if letter_content_paragraph_index is not None:
                    for i in range(letter_content_paragraph_index + 1, len(doc.paragraphs)):
                        paragraph = doc.paragraphs[i]
                        if paragraph.text.strip():
                            for run in paragraph.runs:
                                run.font.name = font_name
                                run.font.size = Pt(font_size)
                                run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

                for idx in sorted(paragraphs_to_remove, reverse=True):
                    p = doc.paragraphs[idx]
                    p._element.getparent().remove(p._element)

                return doc, replacements_made

            replacements = {
                "<<Date>>": letter_date,
                "<<Addressee>>": addressee,
                "<<Salutation>>": salutation,
                "<<Enter text here>>": letter_text
            }

            updated_doc, replacements_made = replace_text_in_document(template, replacements)
            st.session_state.processed_doc = updated_doc
            st.session_state.cache_key = cache_key
            st.session_state.last_font_name = font_name
            st.session_state.last_font_size = font_size

        st.success("🎉 Letter formatted successfully.")
        docx_buffer = io.BytesIO()
        st.session_state.processed_doc.save(docx_buffer)
        docx_buffer.seek(0)

        st.download_button(
            label="📥 Download Letter (DOCX)",
            data=docx_buffer.getvalue(),
            file_name=f"{filename}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"❌ Error: {e}")
else:
    if not os.path.exists(template_path):
        st.error("📁 Local template file `Shah_LOS_template.docx` not found.")
    elif not letter_text:
        st.info("📝 No letter text found.")
    elif not salutation:
        st.info("👋 Missing salutation.")
