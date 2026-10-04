# 📚 Course File Comparator

Compare any faculty course file with **Dr. Sachin’s standard course file** (which already contains both Theory + Lab content).

### Features
- Upload Dr. Sachin’s standard file **once** (from the sidebar)
- Faculty only need to upload **their own course file**
- Shows **Missing** content and **Extra** content
- Coverage score (%)
- Downloadable Markdown report
- Supports PDF, DOCX and TXT

### How to use
1. Open the **sidebar**
2. Upload **Dr. Sachin’s standard course file** (only once)
3. Upload **your course file**
4. Click **Generate Comparison Report**

### Run locally

```bash
git clone https://github.com/YOUR_USERNAME/course-file-comparator.git
cd course-file-comparator
python -m venv venv
source venv/bin/activate          # Windows: venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

### Deploy on Streamlit Cloud (Free)
1. Push this repository to GitHub
2. Go to https://share.streamlit.io
3. Connect your GitHub account and select this repository
4. Main file path: `app.py`
5. Deploy

### Notes
- The first load may take 20–40 seconds (AI model download).
- Standard file is saved in the `standards/` folder.
- Adjust the **Similarity Threshold** in the sidebar if needed.
