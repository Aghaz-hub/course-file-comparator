import streamlit as st
import pdfplumber
from docx import Document
import re
from datetime import datetime
from pathlib import Path
import traceback
from difflib import SequenceMatcher

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

def normalize_heading(text):
    """Clean heading for comparison"""
    text = text.lower().strip()
    text = re.sub(r'^\d+[\.\)]\s*', '', text)          # remove leading numbers
    text = re.sub(r'^[a-z][\.\)]\s*', '', text)         # remove a. b. etc
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'[^\w\s]', '', text)                 # remove special chars
    return text.strip()

def is_likely_heading(line):
    """Heuristic to detect headings"""
    line = line.strip()
    if not line or len(line) < 4 or len(line) > 120:
        return False
    
    # Common patterns in course files
    heading_patterns = [
        r'^\d+\.',                          # 1. 2. 3.
        r'^\d+\.\d+',                       # 1.1 2.3
        r'^[A-Z][A-Z\s]{3,}$',             # ALL CAPS
        r'^SECTION',                       # SECTION A
        r'^PART',
        r'^UNIT',
        r'^CHAPTER',
        r'^OUTCOME',
        r'^COURSE',
        r'^VISION',
        r'^MISSION',
        r'^PROGRAM',
        r'^PEO',
        r'^PO\d',
        r'^PSO',
        r'^GA\d',
        r'^WK\d',
        r'^SK\d',
        r'^AK\d',
        r'^CO\d',
        r'^SDG',
        r'^LESSON PLAN',
        r'^ASSESSMENT',
        r'^RUBRIC',
        r'^ASSIGNMENT',
        r'^QUESTION',
        r'^END.?TERM',
        r'^MID.?SEM',
        r'^LAB',
        r'^THEORY',
        r'^ATTENDANCE',
        r'^SLOW LEARNER',
        r'^ADVANCED LEARNER',
        r'^IMPACT ANALYSIS',
        r'^ALTERNATIVE PLAN',
        r'^APPENDICES',
        r'^TEXTBOOK',
        r'^REFERENCE',
        r'^MOOC',
        r'^ONLINE RESOURCE',
    ]
    
    for pat in heading_patterns:
        if re.search(pat, line, re.IGNORECASE):
            return True
    
    # Short lines that look like titles
    if len(line) < 70 and line[0].isupper() and not line.endswith('.'):
        words = line.split()
        if 1 <= len(words) <= 10:
            return True
    
    return False

def extract_headings(text):
    """Extract list of headings from full text"""
    lines = text.split('\n')
    headings = []
    seen = set()
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if is_likely_heading(line):
            norm = normalize_heading(line)
            if norm and norm not in seen and len(norm) > 3:
                seen.add(norm)
                headings.append(line.strip())
    
    return headings

def similarity(a, b):
    return SequenceMatcher(None, a, b).ratio()

def compare_headings(std_headings, your_headings, threshold=0.72):
    """Find missing and extra headings"""
    missing = []
    matched_your = set()
    
    your_norm = [normalize_heading(h) for h in your_headings]
    
    for std_h in std_headings:
        std_n = normalize_heading(std_h)
        best_score = 0
        best_idx = -1
        for i, y_n in enumerate(your_norm):
            score = similarity(std_n, y_n)
            if score > best_score:
                best_score = score
                best_idx = i
        
        if best_score < threshold:
            missing.append({"heading": std_h, "score": best_score})
        else:
            matched_your.add(best_idx)
    
    extra = []
    for i, your_h in enumerate(your_headings):
        if i not in matched_your:
            extra.append({"heading": your_h, "score": 0.0})
    
    covered = len(std_headings) - len(missing)
    coverage = (covered / len(std_headings) * 100) if std_headings else 0.0
    return missing, extra, round(coverage, 1)

def generate_report(missing, extra, coverage, your_name, std_name, threshold):
    report = f"""# Course File Heading Comparison Report
**Generated:** {datetime.now().strftime("%Y-%m-%d %H:%M")}
**Your File:** {your_name}
**Standard File (Dr. Sachin):** {std_name}
**Matching Threshold:** {threshold}
**Coverage Score:** {coverage}%

---

## Missing Headings ({len(missing)} items)
These headings are present in **Dr. Sachin’s standard file** but missing (or weakly matched) in **your file**:

"""
    if not missing:
        report += "✅ No significant missing headings found.\n"
    else:
        for i, item in enumerate(missing, 1):
            report += f"{i}. **{item['heading']}**  (match score: {item['score']:.2f})\n"
    
    report += f"""
---

## Extra Headings ({len(extra)} items)
These headings are present in **your file** but not found in **Dr. Sachin’s standard file**:

"""
    if not extra:
        report += "✅ No significant extra headings found.\n"
    else:
        for i, item in enumerate(extra, 1):
            report += f"{i}. **{item['heading']}**\n"
    
    return report

# ====================== UI ======================
st.title("📚 Course File Comparator (Headings Only)")
st.markdown("Compares **headings & sub-headings** structure against Dr. Sachin’s standard course file.")

# ---------- SIDEBAR ----------
with st.sidebar:
    st.header("⚙️ Settings")
    threshold = st.slider(
        "Heading Match Threshold",
        0.50, 0.95, 0.72, 0.01,
        help="Higher = stricter matching of headings"
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

if st.button("🚀 Generate Heading Comparison Report", type="primary", use_container_width=True):
    
    std_path = load_standard_file()
    
    if not std_path:
        st.error("❌ Please upload Dr. Sachin’s standard course file from the **sidebar** first.")
    elif not your_file:
        st.error("❌ Please upload your course file.")
    else:
        try:
            progress = st.progress(0)
            status = st.empty()
            
            status.info("Step 1/3 : Extracting text from Dr. Sachin standard file...")
            progress.progress(20)
            std_text = extract_text_from_path(std_path)
            
            status.info("Step 2/3 : Extracting text from your course file...")
            progress.progress(50)
            your_text = extract_text_from_upload(your_file)
            
            if not std_text.strip():
                st.error("❌ Could not extract any text from Dr. Sachin’s standard file.")
            elif not your_text.strip():
                st.error("❌ Could not extract any text from your course file.")
            else:
                status.info("Step 3/3 : Extracting and comparing headings...")
                progress.progress(80)
                
                std_headings = extract_headings(std_text)
                your_headings = extract_headings(your_text)
                
                missing, extra, coverage = compare_headings(std_headings, your_headings, threshold)
                
                progress.progress(100)
                status.success("✅ Comparison complete!")
                
                # ===== RESULTS =====
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Your Headings", len(your_headings))
                m2.metric("Standard Headings", len(std_headings))
                m3.metric("Missing Headings", len(missing))
                m4.metric("Coverage Score", f"{coverage}%")
                
                st.progress(min(coverage / 100, 1.0))
                st.caption(f"Your file covers approximately **{coverage}%** of the required headings from Dr. Sachin’s file.")
                
                # Show extracted headings for transparency
                with st.expander("📋 View all extracted headings from Standard file"):
                    for i, h in enumerate(std_headings, 1):
                        st.write(f"{i}. {h}")
                
                with st.expander("📋 View all extracted headings from Your file"):
                    for i, h in enumerate(your_headings, 1):
                        st.write(f"{i}. {h}")
                
                # Missing
                st.subheader(f"🔍 Missing Headings ({len(missing)})")
                if not missing:
                    st.success("🎉 Excellent! All important headings from the standard file are present.")
                else:
                    for i, item in enumerate(missing, 1):
                        st.markdown(f"**{i}. {item['heading']}**  \nMatch score: `{item['score']:.2f}`")
                
                # Extra
                st.subheader(f"➕ Extra Headings in your file ({len(extra)})")
                if not extra:
                    st.success("No significant extra headings found.")
                else:
                    for i, item in enumerate(extra, 1):
                        st.markdown(f"**{i}. {item['heading']}**")
                
                # Download
                report = generate_report(
                    missing, extra, coverage,
                    your_file.name, std_path.name, threshold
                )
                
                st.markdown("---")
                st.download_button(
                    label="📥 Download Full Report (Markdown)",
                    data=report,
                    file_name=f"heading_comparison_{datetime.now().strftime('%Y%m%d_%H%M')}.md",
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
            "3️⃣ Click **Generate Heading Comparison Report**.")

st.markdown("---")
st.caption("This version focuses only on headings & sub-headings structure (as required for course file completeness).")
