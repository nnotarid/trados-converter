# Trados XML Analysis to Excel Converter

A Streamlit web application that converts Trados Studio XML analysis files into professionally formatted Excel workbooks. Supports multilingual projects with each language in a separate tab.

## Features

✅ **Grouped Match Bands**: No Match, 85-94, 95-99, 100%, Reps, 100% Reps
✅ **Weighted Word Count**: Automatic calculation using industry-standard formulas
✅ **Multi-Language Support**: Each target language gets its own Excel tab
✅ **Professional Formatting**: Tables, borders, colors matching your VBA macro
✅ **Batch Processing**: Upload multiple XML files at once

## How It Works

1. Upload one or more Trados XML analysis files (one per target language)
2. Click "Convert to Excel"
3. Download the formatted Excel workbook

The app automatically:
- Parses XML files to extract match band data
- Groups fuzzy matches according to your specifications
- Calculates weighted word counts (No Match: 100%, 85-94: 60%, 95-99: 40%, 100%+Reps: 33%)
- Creates separate tabs for each language
- Applies professional Excel formatting

## Deployment to Streamlit Community Cloud (FREE)

### Step 1: Create a GitHub Account
If you don't have one already, go to https://github.com and sign up.

### Step 2: Create a New Repository

1. Go to https://github.com/new
2. Repository name: `trados-converter` (or any name you prefer)
3. Description: "Trados XML to Excel converter"
4. Select **Public**
5. Click "Create repository"

### Step 3: Upload Files to GitHub

1. On your new repository page, click "uploading an existing file"
2. Upload these files:
   - `trados_converter_app.py`
   - `requirements.txt`
   - `README.md` (this file)
3. Commit the files

### Step 4: Deploy to Streamlit Community Cloud

1. Go to https://share.streamlit.io
2. Click "New app"
3. Sign in with your GitHub account
4. Grant Streamlit access to your repository
5. Fill in the deployment form:
   - **Repository**: Select your `trados-converter` repository
   - **Branch**: main (or master)
   - **Main file path**: `trados_converter_app.py`
6. Click "Deploy!"

The app will deploy in 2-3 minutes. You'll get a permanent URL like:
`https://yourname-trados-converter-app.streamlit.app`

### Step 5: Share with Your Team

Share the URL with your team. Anyone can access it and use it immediately!

## Local Testing (Optional)

If you want to run it locally first:

```bash
pip install -r requirements.txt
streamlit run trados_converter_app.py
```

The app will open in your browser at `http://localhost:8501`

## Match Band Calculations

The app replicates your VBA macro's logic exactly:

| Band | Components | Weighting |
|------|-----------|-----------|
| **No Match** | New + Fuzzy 50-74 + Fuzzy 75-84 + Internal 50-84 | 100% |
| **85-94** | Fuzzy 85-94 + Internal 85-94 | 60% |
| **95-99** | Fuzzy 95-99 + Internal 95-99 | 40% |
| **100%** | Exact + Perfect + Context | 0% (but included in 100% Reps) |
| **Reps** | Repetitions + Cross-file Repetitions | 0% (but included in 100% Reps) |
| **100% Reps** | 100% + Reps combined | 33% |

**Weighted Formula:**
```
Weighted = No Match 
         + (85-94 × 0.6) 
         + (95-99 × 0.4) 
         + (100% Reps × 0.33)
```

## File Structure

```
trados-converter/
├── trados_converter_app.py    # Main Streamlit application
├── requirements.txt            # Python dependencies
└── README.md                   # This file
```

## Troubleshooting

**Q: The app says "No valid analysis files could be parsed"**
A: Make sure you're uploading XML files exported from Trados Studio as analysis reports, not project files or other XML types.

**Q: The language tabs have weird names**
A: Excel limits sheet names to 31 characters. Long language names are automatically truncated.

**Q: Can I customize the weighted calculation?**
A: Yes! Edit line 93-96 in `trados_converter_app.py` to change the weighting percentages.

**Q: Does this work with Trados GroupShare?**
A: Yes! The XML format is the same. Just export the analysis from GroupShare.

## Support

If you encounter issues:
1. Check that your XML files are valid Trados analysis exports
2. Try uploading one file at a time to identify problematic files
3. Check the Streamlit Cloud logs if deployed

## License

Free to use and modify for your team's needs.

---

**Built with:**
- Python 3.9+
- Streamlit (web framework)
- openpyxl (Excel generation)
- xml.etree.ElementTree (XML parsing)
