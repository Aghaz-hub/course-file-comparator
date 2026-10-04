import streamlit as st
import pdfplumber
from docx import Document
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity
import numpy as np
import re
from datetime import datetime
from pathlib import Path
import traceback

st.set_page_config(
    page_title="Course File Comparator",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ====================== PATHS ======================
STANDARDS_DIR = Path("standards")
STANDARDS_DIR.mkdir(exist_ok=True)
STANDARD_PATH = STANDARDS_DIR / "dr_sachin_standard"

# ====================== MODEL ======================
@st.cache_resource(show_spinner="Loading AI model (first time may take 30-60 seconds)...")
def load_model():
    try:
        return SentenceTransformer("all-MiniLM-L6-v2")
    except Exception as e:
        st.error(f"Failed to load AI model: {e}")
        return None

# ====================== HELPERS ======================
def save_uploaded_file(uploaded_file, save_path_without_ext):
    if uploaded_file is None:
        return None
    ext = Path(uploaded_file.name).suffix.lower()
    full_path = Path(str(save_path_without_ext) + ext)
    for old_ext in [".pdf", ".docx", ".txt"]:
        old = Path(str(save_path_without_ext) + old_ext)
        if old.exists() and old != full_path:
            try:
                old.unlink()
            except:
                pass
    with open(full_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return full_path

def load_standard_file():
    for ext in [".pdf", ".docx", ".txt"]:
        path = Path(str(STANDARD_PATH) + ext)
        if path.exists():
            return path
    return None

def extract_text_from_path(path):
    if path is None or not path.exists():
        return ""
    name = path.name.lower()
    try:
        if name.endswith(".pdf"):
            text = ""
            with pdfplumber.open(path) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"
            return text
        elif name.endswith(".docx"):
            doc = Document(path)
            parts = []
            for p in doc.paragraphs:
                if p.text.strip():
                    parts.append(p.text.strip())
            # Also extract text from tables
            for table in doc.tables:
                for row in table.rows:
                    row_text = " | ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
                    if row_text:
                        parts.append(row_text)
            return "\n".join(parts)
        elif name.endswith(".txt"):
            return path.read_text(encoding="utf-8", errors="ignore")
    except Exception as e:
        st.warning(f"Error extracting text: {e}")
    return ""

def extract_text_from_upload(uploaded_file):
    if uploaded_file is None:
        return ""
    name = uploaded_file.name.lower()
    try:
        if name.endswith(".pdf"):
            text = ""
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"
            return text
        elif name.endswith(".docx"):
            doc = Document(uploaded_file)
            parts = []
            for p in doc.paragraphs:
                if p.text.strip():
                    parts.append(p.text.strip())
            for table in doc.tables:
                for row in table.rows:
                    row_text = " | ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
                    if row_text:
                        parts.append(row_text)
            return "\n".join(parts)
        elif name.endswith(".txt"):
            return uploaded_file.read().decode("utf-8", errors="ignore")
    except Exception as e:
        st.warning(f"Error extracting text from upload: {e}")
    return ""

def clean_text(text):
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def extract_chunks(text, max_chars=350):
    text = clean_text(text)
    if not text:
        return []
    
    # Prefer splitting on numbered sections / headings
    paragraphs = [p.strip() for p in re.split(r'\n+', text) if len(p.strip()) > 30]
    chunks = []
    current = ""
    
    for para in paragraphs:
        if len(current) + len(para) < max_chars:
            current += " " + para
        else:
            if current:
                chunks.append(current.strip())
            current = para
    if current:
        chunks.append(current.strip())
    
    # Limit total chunks to avoid memory issues on free Streamlit
    if len(chunks) > 120:
        chunks = chunks[:120]
    
    return [c for c in chunks if len(c) > 40]

def compare_files(std_chunks, your_chunks, threshold=0.58):
    if not std_chunks and not your_chunks:
        return [], [], 0.0
    
    model = load_model()
    if model is None:
        return [], [], 0.0
    
    try:
        std_emb = model.encode(std_chunks, show_progress_bar=False, batch_size=16) if std_chunks else np.array([])
        your_emb = model.encode(your_chunks, show_progress_bar=False, batch_size=16) if your_chunks else np.array([])
    except Exception as e:
        st.error(f"Embedding failed: {e}")
        return [], [], 0.0
    
    missing = []
    covered = 0
    
    for i, chunk in enumerate(std_chunks):
        if len(your_emb) == 0:
            missing.append({"text": chunk, "similarity": 0.0})
            continue
        sims = cosine_similarity([std_emb[i]], your_emb)[0]
        max_sim = float(np.max(sims))
        if max_sim < threshold:
            missing.append({"text": chunk, "similarity": max_sim})
        else:
            covered += 1
    
    extra = []
    if len(std_emb) > 0:
        for i, chunk in enumerate(your_chunks):
            sims = cosine_similarity([your_emb[i]], std_emb)[0]
            max_sim = float(np.max(sims))
            if max_sim < threshold:
                extra.append({"text": chunk, "similarity": max_sim})
    
    coverage = (covered / len(std_chunks) * 100) if std_chunks else 0.0
    return missing, extra, round(coverage, 1)

def generate_report(missing, extra, coverage, your_name, std_name, threshold):
    report = f"""# Course File Comparison Report
**Generated:** {datetime.now().strftime("%Y-%m-%d %H:%M")}
**Your File:** {your_name}
**Standard File (Dr. Sachin):** {std_name}
**Similarity Threshold:** {threshold}
**Coverage Score:** {coverage}%

---

## Missing Content ({len(missing)} items)
Content present in **Dr. Sachin’s file** but missing or weakly covered in **your file**:

"""
    if not missing:
        report += "✅ No significant missing content found.\n"
    else:
        for i, item in enumerate(missing[:40], 1):  # limit report size
            report += f"\n### Missing #{i} (Similarity: {item['similarity']:.2f})\n"
            report += item["text"][:700] + ("..." if len(item["text"]) > 700 else "") + "\n"
        if len(missing) > 40:
            report += f"\n... and {len(missing)-40} more items.\n"
    
    report += f"""
---

## Extra Content ({len(extra)} items)
Content present in **your file** but not found in **Dr. Sachin’s file**:

"""
    if not extra:
        report += "✅ No significant extra content found.\n"
    else:
        for i, item in enumerate(extra[:40], 1):
            report += f"\n### Extra #{i} (Similarity: {item['similarity']:.2f})\n"
            report += item["text"][:700] + ("..." if len(item["text"]) > 700 else "") + "\n"
        if len(extra) > 40:
            report += f"\n... and {len(extra)-40} more items.\n"
    
    return report

# ====================== UI ======================
st.title("📚 Course File Comparator")
st.markdown("Compare any faculty course file with **Dr. Sachin’s standard course file**.")

# ---------- SIDEBAR ----------
with st.sidebar:
    st.header("⚙️ Settings")
    threshold = st.slider(
        "Similarity Threshold",
        0.40, 0.80, 0.58, 0.01,
        help="Lower = more items marked as Missing/Extra"
    )
    
    st.markdown("---")
    st.header("📌 Dr. Sachin Standard File")
    st.caption("Upload once. (On free Streamlit Cloud you may need to re-upload after app restart)")
    
    existing_std = load_standard_file()
    if existing_std:
        st.success(f"✅ Loaded: `{existing_std.name}`")
        if st.button("Remove Standard File"):
            try:
                existing_std.unlink()
            except:
                pass
            st.rerun()
    else:
        st.warning("No standard file uploaded yet")
    
    std_upload = st.file_uploader(
        "Upload Dr. Sachin Course File",
        type=["pdf", "docx", "txt"],
        key="standard"
    )
    if std_upload:
        with st.spinner("Saving standard file..."):
            saved = save_uploaded_file(std_upload, STANDARD_PATH)
            st.success(f"Saved: {saved.name}")
            st.rerun()

# ---------- MAIN AREA ----------
st.subheader("📤 Upload Faculty Course File")

your_file = st.file_uploader(
    "Upload your course file (PDF / DOCX / TXT)",
    type=["pdf", "docx", "txt"],
    key="your_file"
)

if st.button("🚀 Generate Comparison Report", type="primary", use_container_width=True):
    
    std_path = load_standard_file()
    
    if not std_path:
        st.error("❌ Please upload Dr. Sachin’s standard course file from the **sidebar** first.")
    elif not your_file:
        st.error("❌ Please upload your course file.")
    else:
        try:
            progress = st.progress(0)
            status = st.empty()
            
            status.info("Step 1/4 : Extracting text from Dr. Sachin standard file...")
            progress.progress(10)
            std_text = extract_text_from_path(std_path)
            
            status.info("Step 2/4 : Extracting text from your course file...")
            progress.progress(30)
            your_text = extract_text_from_upload(your_file)
            
            if not std_text.strip():
                st.error("❌ Could not extract any text from Dr. Sachin’s standard file. Please try a different format.")
            elif not your_text.strip():
                st.error("❌ Could not extract any text from your course file. Please try a different format.")
            else:
                status.info(f"Step 3/4 : Creating chunks (Standard: {len(std_text)} chars, Your: {len(your_text)} chars)...")
                progress.progress(50)
                
                std_chunks = extract_chunks(std_text)
                your_chunks = extract_chunks(your_text)
                
                st.write(f"Standard chunks: **{len(std_chunks)}** | Your chunks: **{len(your_chunks)}**")
                
                if len(std_chunks) == 0 or len(your_chunks) == 0:
                    st.error("Not enough text content found to compare.")
                else:
                    status.info("Step 4/4 : Running AI comparison (this may take 20-60 seconds)...")
                    progress.progress(70)
                    
                    missing, extra, coverage = compare_files(std_chunks, your_chunks, threshold)
                    
                    progress.progress(100)
                    status.success("✅ Comparison complete!")
                    
                    # ===== RESULTS =====
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("Your Chunks", len(your_chunks))
                    m2.metric("Standard Chunks", len(std_chunks))
                    m3.metric("Missing Items", len(missing))
                    m4.metric("Coverage Score", f"{coverage}%")
                    
                    st.progress(min(coverage / 100, 1.0))
                    st.caption(f"Your file covers approximately **{coverage}%** of Dr. Sachin’s content.")
                    
                    # Missing
                    st.subheader(f"🔍 Missing Content ({len(missing)} items)")
                    if not missing:
                        st.success("🎉 No significant missing content found.")
                    else:
                        for i, item in enumerate(missing[:30], 1):
                            with st.expander(f"Missing #{i}  •  Similarity: {item['similarity']:.2f}"):
                                st.write(item["text"])
                        if len(missing) > 30:
                            st.info(f"... and {len(missing)-30} more missing items (see full report download)")
                    
                    # Extra
                    st.subheader(f"➕ Extra Content ({len(extra)} items)")
                    if not extra:
                        st.success("No significant extra content found.")
                    else:
                        for i, item in enumerate(extra[:30], 1):
                            with st.expander(f"Extra #{i}  •  Similarity: {item['similarity']:.2f}"):
                                st.write(item["text"])
                        if len(extra) > 30:
                            st.info(f"... and {len(extra)-30} more extra items (see full report download)")
                    
                    # Download
                    report = generate_report(
                        missing, extra, coverage,
                        your_file.name, std_path.name, threshold
                    )
                    
                    st.markdown("---")
                    st.download_button(
                        label="📥 Download Full Report (Markdown)",
                        data=report,
                        file_name=f"course_comparison_{datetime.now().strftime('%Y%m%d_%H%M')}.md",
                        mime="text/markdown",
                        use_container_width=True
                    )
                    
        except Exception as e:
            st.error("❌ An error occurred while generating the report.")
            st.code(str(e))
            with st.expander("Technical details"):
                st.code(traceback.format_exc())

else:
    st.info("1️⃣ First upload **Dr. Sachin’s standard course file** from the sidebar (only once).\n\n"
            "2️⃣ Then upload **your course file** above.\n\n"
            "3️⃣ Click **Generate Comparison Report**.")

st.markdown("---")
st.caption("Note: Comparing two completely different courses (e.g. C Programming vs Computer Networks) will show almost everything as Missing/Extra. Use the same course for meaningful results.")
