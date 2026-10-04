# 📚 Course File Comparator – Theory & Lab

Compare your **Theory** and **Lab** course files against Dr. Sachin’s standard files.

### Features
- Pre-upload Dr. Sachin’s Theory + Lab standard files once (from the sidebar)
- Upload your Theory + Lab files
- Separate reports for Theory and Lab
- Shows **Missing** content and **Extra** content
- Coverage score (%)
- Downloadable Markdown report
- Supports PDF, DOCX and TXT

### How to use
1. Open the **sidebar**
2. Upload Dr. Sachin **Theory Standard** and **Lab Standard** (only once)
3. Upload your Theory and Lab files in the main area
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
4. Deploy

### Folder Structure
```
course-file-comparator/
├── app.py
├── requirements.txt
├── README.md
├── .gitignore
├── .streamlit/
│   └── config.toml
└── standards/               # Standard files are saved here
    └── .gitkeep
```

### Notes
- Standard files are saved in the `standards/` folder.
- On Streamlit Cloud the files persist only during the session (or until the app restarts). For permanent storage you can later add cloud storage.
- Adjust the **Similarity Threshold** in the sidebar if too many/too few items are marked as missing.
