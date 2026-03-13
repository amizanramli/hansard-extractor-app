import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import io
import re

# ---------------------------
# Page Configuration
# ---------------------------
st.set_page_config(
    page_title="Hansard Curator",
    page_icon="📜",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------------------
# Custom Styling
# ---------------------------
st.markdown("""
<style>
.main {
    background: #0f172a;
    color: #f8fafc;
}
.stApp {
    background-color: #0f172a;
}
h1 {
    color: #f8fafc;
    text-align: center;
    margin-bottom: 0.5rem;
}
h1 span {
    background: linear-gradient(135deg, #60a5fa, #c026d3);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
}
.stAlert {
    background-color: rgba(30, 41, 59, 0.6) !important;
    border: 1px solid rgba(255, 255, 255, 0.1) !important;
    backdrop-filter: blur(12px) !important;
    color: #f8fafc !important;
}
.stButton>button {
    background: linear-gradient(135deg, #6366f1, #c026d3);
    color: white;
    border: none;
    border-radius: 8px;
    transition: transform 0.2s ease-in-out;
}
.stButton>button:hover {
    transform: translateY(-2px);
    box-shadow: 0 10px 15px -3px rgba(99,102,241,0.5);
}
</style>
""", unsafe_allow_html=True)

# ---------------------------
# PDF Processing
# ---------------------------
def process_hansard_pdf(file_bytes: bytes, start_page: int, end_page: int):

    transcript = []

    with fitz.open(stream=file_bytes, filetype="pdf") as doc:

        start_idx = max(0, start_page - 1)
        end_idx = min(len(doc), end_page)

        speaker_pattern = re.compile(
            r"^([A-Z][a-zA-Z\s\(\)@\-\’\'\.]+?)(?:\s*\[(.*?)\])?\s*:\s*(.*)",
            re.MULTILINE | re.DOTALL
        )

        current_speaker = None
        current_constituency = None
        current_speech = []

        for i in range(start_idx, end_idx):

            page = doc[i]
            blocks = page.get_text("blocks")

            for block in blocks:
                text = block[4].strip()

                if not text:
                    continue

                # Skip page numbers or artifacts
                if re.match(r"^\d+$", text) or text.startswith(("DR.", "■")):
                    continue

                match = speaker_pattern.match(text)

                if match:

                    if current_speaker:
                        transcript.append({
                            "Speaker": current_speaker.strip(),
                            "Constituency": current_constituency.strip() if current_constituency else "",
                            "Speech": "\n".join(current_speech).strip(),
                            "Page": i + 1
                        })

                    current_speaker = match.group(1).replace("\n", " ")
                    current_constituency = match.group(2).replace("\n", " ") if match.group(2) else ""
                    speech_text = match.group(3).strip()

                    current_speech = [speech_text] if speech_text else []

                else:
                    if current_speaker:
                        current_speech.append(text.replace("\n", " "))

        if current_speaker:
            transcript.append({
                "Speaker": current_speaker.strip(),
                "Constituency": current_constituency.strip() if current_constituency else "",
                "Speech": "\n".join(current_speech).strip(),
                "Page": end_page
            })

    return transcript


# ---------------------------
# Normalize Speaker Names
# ---------------------------
def normalize_speakers(transcript):

    titles_to_remove = [
        "Yang Berhormat ", "Yang Amat Berhormat ",
        "Dato' Seri ", "Dato' Sri ", "Datuk Seri ",
        "Datuk ", "Dato' ", "Tan Sri ", "Tun ",
        "Dr. ", "Tuan ", "Puan ", "Haji ", "Hajjah ",
        "Panglima ", "Ir. ", "Ts. "
    ]

    for entry in transcript:

        clean_name = entry["Speaker"]

        for title in titles_to_remove:
            clean_name = re.compile(re.escape(title), re.IGNORECASE).sub("", clean_name)

        entry["Normalized_Speaker"] = clean_name.strip()

    return transcript


# ---------------------------
# Excel Export
# ---------------------------
def dataframe_to_xlsx(df):

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Transcript")

    return output.getvalue()


# ---------------------------
# UI
# ---------------------------
st.markdown("<h1>Hansard <span>Extractor </span></h1>", unsafe_allow_html=True)
st.markdown(
    "<p style='text-align:center;color:#94a3b8;margin-bottom:2rem;'>"
    "V1"
    "</p>",
    unsafe_allow_html=True
)

uploaded_files = st.file_uploader(
    "Drag & Drop up to 3 PDFs here",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:

    if len(uploaded_files) > 3:
        st.error("Maximum 3 Hansard PDFs allowed.")

    else:

        st.subheader("Configure Page Ranges")

        cols = st.columns(len(uploaded_files))
        page_ranges = []

        for idx, file in enumerate(uploaded_files):

            with cols[idx]:
                st.write(f"**{file.name}**")

                rng = st.text_input(
                    "Page Range (e.g. 15-50)",
                    key=f"range_{idx}",
                    placeholder="15-50"
                )

                page_ranges.append(rng)

        if st.button("🚀 Process Documents", use_container_width=True):

            with st.spinner("Extracting text and detecting speakers..."):

                combined_transcript = []
                has_error = False

                for idx, file in enumerate(uploaded_files):

                    try:
                        start_str, end_str = page_ranges[idx].split("-")
                        start_p = int(start_str.strip())
                        end_p = int(end_str.strip())

                    except Exception:
                        st.error(f"Invalid page range for {file.name}. Use 'start-end'.")
                        has_error = True
                        break

                    try:

                        file_bytes = file.read()

                        transcript = process_hansard_pdf(
                            file_bytes,
                            start_p,
                            end_p
                        )

                        for item in transcript:
                            item["Document_Name"] = file.name

                        combined_transcript.extend(transcript)

                    except Exception as e:
                        st.error(f"Failed processing {file.name}: {e}")
                        has_error = True
                        break

                if not has_error:

                    combined_transcript = normalize_speakers(combined_transcript)

                    st.success(
                        f"Processing Complete! Extracted {len(combined_transcript)} speech segments."
                    )

                    df = pd.DataFrame(combined_transcript)

                    st.subheader("Data Preview")
                    st.dataframe(df, use_container_width=True, height=400)

                    st.subheader("Export Options")

                    col1, col2 = st.columns(2)

                    with col1:

                        xlsx_data = dataframe_to_xlsx(df)

                        st.download_button(
                            "📥 Download Excel (XLSX)",
                            data=xlsx_data,
                            file_name="hansard_export.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

                    with col2:

                        json_str = df.to_json(
                            orient="records",
                            indent=4
                        )

                        st.download_button(
                            "📥 Download JSON",
                            data=json_str,
                            file_name="hansard_export.json",
                            mime="application/json",
                            use_container_width=True
                        )