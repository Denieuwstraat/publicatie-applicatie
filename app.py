import streamlit as st
import fitz  # PyMuPDF
from docx import Document
import re
import io

st.set_page_config(page_title="Kennisgeving generator", layout="centered")

st.title("📄 Kennisgeving automatisch invullen")
st.write("Sleep hieronder een PDF-bestand van de e-mail in, en ontvang het ingevulde Word-document.")

# === PDF uploaden ===
uploaded_pdf = st.file_uploader("📥 Sleep hier de PDF van de e-mail in", type=["pdf"])

if uploaded_pdf is not None:
    # Lees de tekst uit PDF met PyMuPDF
    with fitz.open(stream=uploaded_pdf.read(), filetype="pdf") as doc:
        pdf_text = "\n".join(page.get_text() for page in doc)

    # Extractie via regex
    def extract_field(patroon, tekst, fallback=""):
        match = re.search(patroon, tekst, re.IGNORECASE | re.DOTALL)
        return match.group(1).strip() if match else fallback

    data = {
        "Omgevingsloket-nummer": extract_field(r"Omgevingsloket-nummer[:\s]*([A-Z0-9_]+)", pdf_text),
        "Dossiernummer": extract_field(r"Dossiernummer[:\s]*([A-Z0-9\-\/]+)", pdf_text),
        "Gegevens aanvrager": extract_field(r"Gegevens aanvrager[:\s]*(.+?)\n", pdf_text),
        "Gegevens van de exploitant": extract_field(r"Gegevens van de exploitant[:\s]*(.+?)\n", pdf_text),
        "Ligging van het project": extract_field(r"Ligging van het project[:\s]*(.+?)\n", pdf_text),
        "Kadastrale gegevens": extract_field(r"Kadastrale gegevens[:\s]*(.+?)\n", pdf_text),
        "Onderwerp van het verzoek": extract_field(r"Onderwerp van het verzoek[:\s]*(.+?)\n", pdf_text),
    }

    # Toon ter controle
    with st.expander("🔍 Geëxtraheerde gegevens (controleer hier)", expanded=False):
        for k, v in data.items():
            st.markdown(f"**{k}:** {v}")

    # Laad het sjabloon vanuit bestand
    sjabloon_path = "Sjabloon helsinki.docx"
    doc = Document(sjabloon_path)

    # Vervang tekst op basis van exact voorkomen in paragrafen
    for p in doc.paragraphs:
        for key, value in data.items():
            if key in p.text:
                p.text = p.text.replace(key, f"{key}\n{value}")

    # Genereer bestand in geheugen
    output_stream = io.BytesIO()
    doc.save(output_stream)
    output_stream.seek(0)

    # Downloadknop
    st.success("✅ Document ingevuld. Download het hieronder.")
    st.download_button(
        label="📥 Download ingevuld Word-bestand",
        data=output_stream,
        file_name="Kennisgeving ingevuld.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
