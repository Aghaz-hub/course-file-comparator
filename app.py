import streamlit as st
import pdfplumber
from docx import Document
import re
from datetime import datetime
from pathlib import Path
import traceback
from difflib import SequenceMatcher

st.set_page_config(
    page_title="Course File Structure Checker",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ====================== PATHS ======================
STANDARDS_DIR = Path("standards")
STANDARDS_DIR.mkdir(exist_ok=True)
STANDARD_PATH = STANDARDS_DIR / "dr_sachin_standard"

# ====================== STANDARD REQUIRED HEADINGS ======================
# These are the structural headings that must be present in every course file
# (taken from the official COURSE FILE CONTENTS checklist)

REQUIRED_HEADINGS = [
    "Vision and Mission of the University",
    "Vision and Mission of the School",
    "Vision and Mission of the Department",
    "GA's, PEOs, POs and PSOs of the Program",
    "Graduate Attributes",
    "Program Educational Objectives",
    "Program Outcomes",
    "Program Specific Outcomes",
    "List of KSA Profiles of the Program",
    "Knowledge, Skill and Attitude Profiles",
    "Academic Calendar",
    "Course Design of Theory component",
    "KSA Profile Elements of the Theory component",
    "List of SDGs/Ethics/Skill Levels",
    "Course Objective and List of Course Outcomes",
    "Mapping to Skill/Employment/Entrepreneurship",
    "Relevance of the Theory component",
    "Syllabus with Text books, Reference books, Online Resources, MOOCs",
    "Lesson Plan",
    "CO-PO Mapping",
    "CO-KSA Mapping",
    "Course Assessment Plan",
    "Course Design of Lab component",
    "List of Experiments",
    "List of Miniprojects",
    "Setting CO Targets",
    "List of faculty members teaching the course",
    "Time table of concerned faculty members",
    "Lecture Notes and PowerPoint handouts",
    "Tutorials/Assignments",
    "Solution of Tutorials/Assignments",
    "Mid-Sem Question Paper",
    "End-Term Question Paper",
    "Solution of Mid-Sem and End-Term Question Papers",
    "Question Papers / Continuous Assessment",
    "Rubrics for Lab assessments",
    "Note on how the course can contribute to the ePortfolio",
    "Attendance Record",
    "Continuous Evaluation for all Assessments",
    "Assessment-wise Marks Uploaded on ERP",
    "Records of Rubrics based Evaluation",
    "Minutes of Course Coordinator Level Meetings",
    "CO attainment analysis and action taken",
    "List of slow Learners and Advanced Learners",
    "The action taken to support slow and advanced learners",
    "Impact Analysis Report and Action Taken",
    "Alternative Plan of Action",
    "Sample hardcopies of Tutorials",
    "Sample hardcopies of answer sheets",
]

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

def extract_full_text(source, is_upload=False):
    """Extract all text from PDF / DOCX / TXT"""
    try:
        if is_upload:
            name = source.name.lower()
            file_obj = source
        else:
            name = source.name.lower()
            file_obj = source

        if name.endswith(".pdf"):
            text = ""
            with pdfplumber.open(file_obj) as pdf:
                for page in pdf.pages:
                    t = page.extract_text()
                    if t:
                        text += t + "\n"
            return text

        elif name.endswith(".docx"):
            doc = Document(file_obj)
            parts = []
            for p in doc.paragraphs:
                if p.text.strip():
                    parts.append(p.text.strip())
            for table in doc.tables:
                for row in table.rows:
                    cells = [c.text.strip() for c in row.cells if c.text.strip()]
                    if cells:
                        parts.append(" | ".join(cells))
            return "\n".join(parts)

        elif name.endswith(".txt"):
            if is_upload:
                return file_obj.read().decode("utf-8", errors="ignore")
            else:
                return Path(file_obj).read_text(encoding="utf-8", errors="ignore")
    except Exception as e:
        st.warning(f"Text extraction error: {e}")
    return ""

def normalize(text):
    text = text.lower()
    text = re.sub(r"[^\w\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text

def similarity(a, b):
    return SequenceMatcher(None, normalize(a), normalize(b)).ratio()

def find_heading_in_text(required_heading, full_text, threshold=0.65):
    """
    Search whether a required heading exists in the document text.
    Also try to estimate if there is content after it.
    """
    lines = [l.strip() for l in full_text.split("\n") if l.strip()]
    best_score = 0.0
    best_line = ""
    best_idx = -1

    for i, line in enumerate(lines):
        # Skip very long lines (unlikely to be headings)
        if len(line) > 180:
            continue
        score = similarity(required_heading, line)
        if score > best_score:
            best_score = score
            best_line = line
            best_idx = i

    found = best_score >= threshold

    # Estimate content length after the heading
    content_length = 0
    if found and best_idx >= 0:
        # Take next 8 non-empty lines as rough content sample
        content_lines = []
        for j in range(best_idx + 1, min(best_idx + 12, len(lines))):
            # Stop if we hit another strong heading-like line
            if len(lines[j]) < 100 and lines[j][0].isupper() and similarity(required_heading, lines[j]) < 0.4:
                # possible next heading
                if any(kw in lines[j].lower() for kw in ["vision", "mission", "course", "lesson", "assessment", "outcome", "syllabus", "lab", "theory", "attendance", "slow", "impact"]):
                    break
            content_lines.append(lines[j])
        content_length = sum(len(c) for c in content_lines)

    return {
        "found": found,
        "matched_line": best_line,
        "score": best_score,
        "content_length": content_length
    }

def check_required_headings(full_text, threshold=0.65, min_content=30):
    results = []
    missing_count = 0

    for req in REQUIRED_HEADINGS:
        res = find_heading_in_text(req, full_text, threshold)
        status = "Present"
        reason = ""

        if not res["found"]:
            status = "Missing"
            reason = "Heading not found"
            missing_count += 1
        elif res["content_length"] < min_content:
            status = "Blank / Incomplete"
            reason = f"Heading found but content is too short ({res['content_length']} chars)"
            missing_count += 1
        else:
            reason = "OK"

        results.append({
            "required": req,
            "status": status,
            "reason": reason,
            "matched_line": res["matched_line"],
            "score": res["score"],
            "content_length": res["content_length"]
        })

    total = len(REQUIRED_HEADINGS)
    coverage = round(((total - missing_count) / total) * 100, 1)
    return results, coverage, missing_count

def generate_report(results, coverage, missing_count, your_name, std_name, threshold):
    report = f"""# Course File Structure Checklist Report
**Generated:** {datetime.now().strftime("%Y-%m-%d %H:%M")}
**Faculty File:** {your_name}
**Standard Template:** {std_name}
**Heading Match Threshold:** {threshold}
**Coverage Score:** {coverage}%
**Missing / Blank sections:** {missing_count}

---

## Detailed Checklist

| # | Required Heading | Status | Remarks |
|---|------------------|--------|---------|
"""
    for i, r in enumerate(results, 1):
        report += f"| {i} | {r['required']} | {r['status']} | {r['reason']} |\n"

    report += "\n---\n\n## Missing / Incomplete Sections\n\n"
    miss = [r for r in results if r["status"] != "Present"]
    if not miss:
        report += "✅ All required structural headings are present and have content.\n"
    else:
        for i, r in enumerate(miss, 1):
            report += f"{i}. **{r['required']}**\n"
            report += f"   - Status: {r['status']}\n"
            report += f"   - Reason: {r['reason']}\n"
            if r["matched_line"]:
                report += f"   - Closest match found: `{r['matched_line']}` (score {r['score']:.2f})\n"
            report += "\n"

    return report

# ====================== UI ======================
st.title("📚 Course File Structure Checker")
st.markdown("""
This tool checks whether a faculty course file contains **all the required structural headings** 
from the official course-file template (as present in Dr. Sachin’s file).

- Course name, code, syllabus content, lecture notes etc. can be different.
- Only the **common headings** (Vision & Mission, POs, KSA, Lesson Plan, Assessment Plan, etc.) are checked.
- A heading is marked incomplete if it is present but has almost no content under it.
""")

# ---------- SIDEBAR ----------
with st.sidebar:
    st.markdown("### Developed by **aghaZ**")
    st.markdown("---")
    st.header("⚙️ Settings")
    threshold = st.slider("Heading Match Sensitivity", 0.50, 0.90, 0.65, 0.01,
                          help="Lower = more flexible matching")
    min_content = st.number_input("Minimum content under heading (chars)", 10, 200, 30, 5)

    st.markdown("---")
    st.header("📌 Dr. Sachin Standard File")
    st.caption("Upload the reference course file that contains the complete checklist of headings.")

    existing = load_standard_file()
    if existing:
        st.success(f"✅ Loaded: `{existing.name}`")
        if st.button("Remove Standard"):
            try:
                existing.unlink()
            except:
                pass
            st.rerun()
    else:
        st.warning("No standard file uploaded yet")

    std_up = st.file_uploader("Upload Dr. Sachin Course File", type=["pdf", "docx", "txt"], key="std")
    if std_up:
        save_uploaded_file(std_up, STANDARD_PATH)
        st.success("Standard file saved")
        st.rerun()

# ---------- MAIN ----------
st.subheader("📤 Upload Faculty Course File")
your_file = st.file_uploader("Upload faculty course file", type=["pdf", "docx", "txt"], key="faculty")

if st.button("🚀 Check Structure", type="primary", use_container_width=True):
    std_path = load_standard_file()

    if not std_path:
        st.error("Please upload Dr. Sachin’s standard file from the sidebar first.")
    elif not your_file:
        st.error("Please upload a faculty course file.")
    else:
        try:
            with st.spinner("Extracting text and checking headings..."):
                # We still extract the standard mainly for display; the real checklist is hard-coded
                std_text = extract_full_text(std_path, is_upload=False)
                your_text = extract_full_text(your_file, is_upload=True)

                if not your_text.strip():
                    st.error("Could not extract text from the faculty file.")
                else:
                    results, coverage, missing_count = check_required_headings(
                        your_text, threshold=threshold, min_content=min_content
                    )

                    # Metrics
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Total Required Headings", len(REQUIRED_HEADINGS))
                    c2.metric("Missing / Blank", missing_count)
                    c3.metric("Coverage", f"{coverage}%")

                    st.progress(min(coverage / 100, 1.0))

                    # Table of results
                    st.subheader("📋 Checklist Result")

                    present = [r for r in results if r["status"] == "Present"]
                    problems = [r for r in results if r["status"] != "Present"]

                    if problems:
                        st.error(f"{len(problems)} headings are Missing or have blank content")
                        for r in problems:
                            with st.expander(f"❌ {r['required']}  —  {r['status']}"):
                                st.write(f"**Reason:** {r['reason']}")
                                if r["matched_line"]:
                                    st.write(f"**Closest match in file:** `{r['matched_line']}` (score {r['score']:.2f})")
                                else:
                                    st.write("No similar heading found in the document.")
                    else:
                        st.success("🎉 All required structural headings are present and contain content!")

                    with st.expander("✅ Present Headings"):
                        for r in present:
                            st.markdown(f"- **{r['required']}**  (matched: `{r['matched_line'][:80]}...`)")

                    # Download
                    report = generate_report(results, coverage, missing_count,
                                             your_file.name, std_path.name, threshold)
                    st.download_button(
                        "📥 Download Full Report",
                        data=report,
                        file_name=f"structure_check_{datetime.now().strftime('%Y%m%d_%H%M')}.md",
                        mime="text/markdown",
                        use_container_width=True
                    )

        except Exception as e:
            st.error("An error occurred")
            st.code(str(e))
            with st.expander("Details"):
                st.code(traceback.format_exc())

else:
    st.info("1. Upload Dr. Sachin’s standard file in the sidebar (once)\n"
            "2. Upload any faculty course file\n"
            "3. Click **Check Structure**")

st.markdown("---")
st.caption("Only structural headings of the official course-file template are checked. Subject-specific content is ignored.")
