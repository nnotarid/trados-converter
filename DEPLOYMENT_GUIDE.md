# Quick Deployment Guide - Trados Converter

## 🚀 Get Your App Online in 5 Minutes

### Step 1: Prepare Your Files (Already Done! ✅)

You have three files:
1. `trados_converter_app.py` - The main application
2. `requirements.txt` - Python dependencies
3. `README.md` - Documentation

### Step 2: Upload to GitHub

**Option A: Use GitHub Web Interface (Easiest)**

1. Go to https://github.com/new
2. Create repository:
   - Name: `trados-converter`
   - Visibility: Public
   - Click "Create repository"
3. Click "uploading an existing file"
4. Drag and drop all three files
5. Click "Commit changes"

**Option B: Use GitHub Desktop (If you prefer GUI)**

1. Download GitHub Desktop from https://desktop.github.com
2. Create new repository
3. Copy the three files to the repository folder
4. Commit and push

### Step 3: Deploy to Streamlit Cloud

1. Go to https://share.streamlit.io/signup
2. Sign in with GitHub
3. Click "New app"
4. Select:
   - Repository: `your-username/trados-converter`
   - Branch: `main`
   - Main file: `trados_converter_app.py`
5. Click "Deploy!"

Wait 2-3 minutes and your app will be live!

### Step 4: Share with Team

Your app URL will be something like:
```
https://yourname-trados-converter.streamlit.app
```

Share this URL with your team - they can use it immediately!

---

## 💡 Tips

**Free Tier Limits:**
- Unlimited public apps
- 1GB storage per app
- Generous bandwidth
- Automatic HTTPS
- No credit card required

**Privacy Note:**
- The code is public (on GitHub)
- The uploaded XML files are NOT stored
- Files are processed in memory and discarded after download
- Each user's session is isolated

**Updating the App:**
Just commit changes to GitHub - Streamlit auto-updates the app!

---

## 🔧 Customization Options

### Change Weighted Calculation
Edit `trados_converter_app.py`, line 93-96:
```python
weighted = round(
    imf_new +                          # No Match: 100%
    round(imf_low_fuzzy * 0.6, 0) +    # 85-94: 60%
    round(imf_high_fuzzy * 0.4, 0) +   # 95-99: 40%
    round(imf_100_reps * 0.33, 0),     # 100% Reps: 33%
    0
)
```

### Change Column Headers
Edit `trados_converter_app.py`, line 207:
```python
headers = ['ID', 'File Name', 'No Match', '85-94', '95-99', '100%', 'Reps', '100% Reps', 'Total (Gross)', 'Weighted (Net)']
```

### Change Color Scheme
Edit `trados_converter_app.py`, lines 281-289 for colors:
```python
header_fill = PatternFill(start_color="4472C4", ...)  # Blue header
gray_fill = PatternFill(start_color="D9D9D9", ...)    # Gray highlight
```

---

## 📊 How It Compares to Your VBA Macro

| Feature | VBA Macro | Streamlit App |
|---------|-----------|---------------|
| **Installation** | Needs Excel, macro-enabled | Just a URL |
| **Platform** | Windows only | Any device, any OS |
| **Updates** | Manual redistribution | Auto-updates from GitHub |
| **Team Access** | Send file to everyone | Share one link |
| **Multi-language** | One at a time | All at once |
| **Security** | Email attachments | HTTPS web app |

---

## ❓ Common Questions

**Q: Will this cost money?**
A: No! Streamlit Community Cloud is completely free for public apps.

**Q: What if I need IT approval?**
A: Show them:
- It's hosted on Streamlit's official cloud (trusted by Fortune 500)
- No data is stored (files processed in memory only)
- Source code is visible on GitHub (full transparency)
- Uses HTTPS encryption
- No access to corporate networks

**Q: Can I make it private?**
A: Yes, but requires Streamlit Teams (paid). For most teams, public deployment is fine since:
- No sensitive data is in the code
- User data isn't stored
- XML files are never saved

**Q: What if GitHub or Streamlit go down?**
A: Keep the Python files - you can run locally with `streamlit run trados_converter_app.py` or deploy to any Python hosting service.

---

## 🎯 Next Steps

1. ✅ Deploy to Streamlit Cloud (follow Step 3 above)
2. 📧 Email the app URL to your team
3. 🧪 Test with a few XML files
4. 📝 Bookmark for easy access
5. 🔄 Update as needed (just push to GitHub)

**Need help?** The Streamlit community forum is very responsive:
https://discuss.streamlit.io

---

Enjoy your new web-based Trados converter! 🎉
