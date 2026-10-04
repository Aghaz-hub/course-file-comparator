import streamlit as st
import pdfplumber
from docx import Document
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity
import numpy as np
import re
from datetime import datetime
from pathlib import Path

st.set_page_config(
    page_title="Course File Comparator – Theory & Lab",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ====================== PATHS ======================
STANDARDS_DIR = Path("standards")
STANDARDS_DIR.mkdir(exist_ok=True)

THEORY_STD_PATH = STANDARDS_DIR / "theory_standard"
LAB_STD_PATH = STANDARDS_DIR / "lab_standard"

# ====================== MODEL ======================
@st.cache_resource
def load_model():
    return SentenceTransformer("all-MiniLM-L6-v2")

model = load_model()

# ====================== HELPERS ======================
def save_uploaded_file(uploaded_file, save_path_without_ext):
    """Save uploaded file and return the full path."""
    if uploaded_file is None:
        return None
    ext = Path(uploaded_file.name).suffix.lower()
    full_path = Path(str(save_path_without_ext) + ext)
    # Remove any previous version with different extension
    for old_ext in [".pdf", ".docx", ".txt"]:
        old = Path(str(save_path_without_ext) + old_ext)
        if old.exists() and old != full_path:
            old.unlink()
    with open(full_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return full_path

def load_standard_file(base_path):
    """Load the first existing standard file (pdf/docx/txt)."""
    for ext in [".pdf", ".docx", ".txt"]:
        path = Path(str(base_path) + ext)
        if path.exists():
            return path
    return None

def extract_text_from_path(path):
    if path is None or not path.exists():
        return ""
    name = path.name.lower()
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
        return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    elif name.endswith(".txt"):
        return path.read_text(encoding="utf-8", errors="ignore")
    return ""

def extract_text_from_upload(uploaded_file):
    if uploaded_file is None:
        return ""
    name = uploaded_file.name.lower()
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
        return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    elif name.endswith(".txt"):
        return uploaded_file.read().decode("utf-8", errors="ignore")
    return ""

def clean_text(text):
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def extract_chunks(text, max_chars=400):
    text = clean_text(text)
    if not text:
        return []
    
    paragraphs = [p.strip() for p in text.split("\n") if len(p.strip()) > 40]
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
    
    if len(chunks) < 3:
        sentences = re.split(r'(?<=[.!?])\s+', text)
        chunks = []
        current = ""
        for s in sentences:
            if len(current) + len(s) < max_chars:
                current += " " + s
            else:
                if current:
                    chunks.append(current.strip())
                current = s
        if current:
            chunks.append(current.strip())
    
    return [c for c in chunks if len(c) > 50]

def compare_files(std_chunks, your_chunks, threshold=0.62):
    """Returns missing (in std but not in yours) and extra (in yours but not in std)."""
    if not std_chunks and not your_chunks:
        return [], [], 0.0
    
    std_emb = model.encode(std_chunks, show_progress_bar=False) if std_chunks else np.array([])
    your_emb = model.encode(your_chunks, show_progress_bar=False) if your_chunks else np.array([])
    
    missing = []
    covered = 0
    
    # Missing = in standard but not well covered in yours
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
    
    # Extra = in yours but not well covered in standard
    extra = []
    if len(std_emb) > 0:
        for i, chunk in enumerate(your_chunks):
            sims = cosine_similarity([your_emb[i]], std_emb)[0]
            max_sim = float(np.max(sims))
            if max_sim < threshold:
                extra.append({"text": chunk, "similarity": max_sim})
    
    coverage = (covered / len(std_chunks) * 100) if std_chunks else 0.0
    return missing, extra, round(coverage, 1)

def generate_section_report(title, missing, extra, coverage, your_name, std_name):
    report = f"""## {title} Comparison Report
**Your File:** {your_name}  
**Standard File:** {std_name}  
**Coverage Score:** {coverage}%

### Missing Content ({len(missing)} items)
Content present in the **standard** but missing or weakly covered in **your** file:
"""
    if not missing:
        report += "\n✅ No significant missing content found.\n"
    else:
        for i, item in enumerate(missing, 1):
            report += f"\n**Missing #{i}** (Similarity: {item['similarity']:.2f})\n"
            report += item["text"][:850] + ("..." if len(item["text"]) > 850 else "") + "\n"
    
    report += f"\n### Extra Content ({len(extra)} items)\n"
    report += "Content present in **your** file but not found in the **standard**:\n"
    
    if not extra:
        report += "\n✅ No significant extra content found.\n"
    else:
        for i, item in enumerate(extra, 1):
            report += f"\n**Extra #{i}** (Similarity: {item['similarity']:.2f})\n"
            report += item["text"][:850] + ("..." if len(item["text"]) > 850 else "") + "\n"
    
    return report

# ====================== UI ======================
st.title("📚 Course File Comparator – Theory & Lab")
st.markdown("Pre-load **Dr. Sachin’s Theory** and **Lab** standard files once. Then upload your Theory + Lab files to get separate Missing & Extra reports.")

# ---------- SIDEBAR: SETTINGS + STANDARD FILES ----------
with st.sidebar:
    st.header("⚙️ Settings")
    threshold = st.slider(
        "Similarity Threshold",
        0.45, 0.85, 0.62, 0.01,
        help="Lower value = more items marked as Missing/Extra"
    )
    
    st.markdown("---")
    st.header("📌 Standard Files (Pre-upload once)")
    
    # Theory Standard
    st.subheader("Theory Standard")
    existing_theory = load_standard_file(THEORY_STD_PATH)
    if existing_theory:
        st.success(f"✅ Loaded: `{existing_theory.name}`")
        if st.button("Remove Theory Standard", key="rm_theory"):
            existing_theory.unlink()
            st.rerun()
    else:
        st.warning("No Theory standard uploaded yet")
    
    theory_std_upload = st.file_uploader(
        "Upload Dr. Sachin Theory Standard",
        type=["pdf", "docx", "txt"],
        key="theory_std"
    )
    if theory_std_upload:
        saved = save_uploaded_file(theory_std_upload, THEORY_STD_PATH)
        st.success(f"Saved Theory standard: {saved.name}")
        st.rerun()
    
    st.markdown("---")
    
    # Lab Standard
    st.subheader("Lab / Practical Standard")
    existing_lab = load_standard_file(LAB_STD_PATH)
    if existing_lab:
        st.success(f"✅ Loaded: `{existing_lab.name}`")
        if st.button("Remove Lab Standard", key="rm_lab"):
            existing_lab.unlink()
            st.rerun()
    else:
        st.warning("No Lab standard uploaded yet")
    
    lab_std_upload = st.file_uploader(
        "Upload Dr. Sachin Lab Standard",
        type=["pdf", "docx", "txt"],
        key="lab_std"
    )
    if lab_std_upload:
        saved = save_uploaded_file(lab_std_upload, LAB_STD_PATH)
        st.success(f"Saved Lab standard: {saved.name}")
        st.rerun()

# ---------- MAIN AREA: YOUR FILES ----------
st.subheader("📤 Upload Your Files")

col1, col2 = st.columns(2)

with col1:
    st.markdown("#### Your Theory File")
    your_theory = st.file_uploader("Upload your Theory course file", type=["pdf", "docx", "txt"], key="your_theory")

with col2:
    st.markdown("#### Your Lab / Practical File")
    your_lab = st.file_uploader("Upload your Lab course file", type=["pdf", "docx", "txt"], key="your_lab")

# ---------- COMPARISON ----------
if st.button("🚀 Generate Comparison Report", type="primary", use_container_width=True):
    
    theory_std_path = load_standard_file(THEORY_STD_PATH)
    lab_std_path = load_standard_file(LAB_STD_PATH)
    
    if not theory_std_path and not lab_std_path:
        st.error("Please upload at least one Standard file (Theory or Lab) from the sidebar first.")
    elif not your_theory and not your_lab:
        st.error("Please upload at least one of your files (Theory or Lab).")
    else:
        with st.spinner("Analyzing files... This may take 15–40 seconds"):
            
            full_report = f"""# Course Comparison Report
**Generated:** {datetime.now().strftime("%Y-%m-%d %H:%M")}
**Threshold used:** {threshold}

"""
            
            # ===== THEORY COMPARISON =====
            if your_theory and theory_std_path:
                st.markdown("---")
                st.header("📖 Theory Comparison")
                
                std_text = extract_text_from_path(theory_std_path)
                your_text = extract_text_from_upload(your_theory)
                
                std_chunks = extract_chunks(std_text)
                your_chunks = extract_chunks(your_text)
                
                missing, extra, coverage = compare_files(std_chunks, your_chunks, threshold)
                
                # Metrics
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Your Chunks", len(your_chunks))
                c2.metric("Standard Chunks", len(std_chunks))
                c3.metric("Missing", len(missing))
                c4.metric("Coverage", f"{coverage}%")
                
                st.progress(min(coverage / 100, 1.0))
                
                # Missing
                st.subheader(f"🔍 Missing in your Theory ({len(missing)})")
                if not missing:
                    st.success("No significant missing content.")
                else:
                    for i, item in enumerate(missing, 1):
                        with st.expander(f"Missing #{i} • Sim: {item['similarity']:.2f}"):
                            st.write(item["text"])
                
                # Extra
                st.subheader(f"➕ Extra in your Theory ({len(extra)})")
                if not extra:
                    st.success("No significant extra content.")
                else:
                    for i, item in enumerate(extra, 1):
                        with st.expander(f"Extra #{i} • Sim: {item['similarity']:.2f}"):
                            st.write(item["text"])
                
                full_report += generate_section_report(
                    "Theory", missing, extra, coverage,
                    your_theory.name, theory_std_path.name
                )
            
            elif your_theory and not theory_std_path:
                st.warning("You uploaded a Theory file but no Theory Standard is set.")
            
            # ===== LAB COMPARISON =====
            if your_lab and lab_std_path:
                st.markdown("---")
                st.header("🧪 Lab / Practical Comparison")
                
                std_text = extract_text_from_path(lab_std_path)
                your_text = extract_text_from_upload(your_lab)
                
                std_chunks = extract_chunks(std_text)
                your_chunks = extract_chunks(your_text)
                
                missing, extra, coverage = compare_files(std_chunks, your_chunks, threshold)
                
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Your Chunks", len(your_chunks))
                c2.metric("Standard Chunks", len(std_chunks))
                c3.metric("Missing", len(missing))
                c4.metric("Coverage", f"{coverage}%")
                
                st.progress(min(coverage / 100, 1.0))
                
                st.subheader(f"🔍 Missing in your Lab ({len(missing)})")
                if not missing:
                    st.success("No significant missing content.")
                else:
                    for i, item in enumerate(missing, 1):
                        with st.expander(f"Missing #{i} • Sim: {item['similarity']:.2f}"):
                            st.write(item["text"])
                
                st.subheader(f"➕ Extra in your Lab ({len(extra)})")
                if not extra:
                    st.success("No significant extra content.")
                else:
                    for i, item in enumerate(extra, 1):
                        with st.expander(f"Extra #{i} • Sim: {item['similarity']:.2f}"):
                            st.write(item["text"])
                
                full_report += "\n\n" + generate_section_report(
                    "Lab / Practical", missing, extra, coverage,
                    your_lab.name, lab_std_path.name
                )
            
            elif your_lab and not lab_std_path:
                st.warning("You uploaded a Lab file but no Lab Standard is set.")
            
            # Download full report
            st.markdown("---")
            st.download_button(
                label="📥 Download Full Report (Theory + Lab)",
                data=full_report,
                file_name=f"course_comparison_report_{datetime.now().strftime('%Y%m%d_%H%M')}.md",
                mime="text/markdown",
                use_container_width=True
            )

else:
    st.info("1️⃣ First upload the Standard files from the **sidebar** (only once).  \n2️⃣ Then upload your Theory and/or Lab files above.  \n3️⃣ Click **Generate Comparison Report**.")

st.markdown("---")
st.caption("Theory and Lab are compared separately against their own standard files.")
