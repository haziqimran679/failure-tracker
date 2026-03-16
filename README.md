# Failure Tracking System — Supabase Version
### Always Online | Permanent Storage | Free

---

## What's Different from Local Version
- Data saves to **Supabase cloud database** (not local Excel)
- App runs on **Streamlit Cloud** (always online, even when your PC is off)
- **Download Excel** button still works — exports all data to Excel anytime
- Permanent shareable link for all colleagues

---

## Deployment Steps

### Step 1 — Upload to GitHub
1. Go to github.com → create new repository named `failure-tracker`
2. Upload these files:
   - `app.py`
   - `requirements.txt`
   - README.md (this file)
   ⚠️ Do NOT upload the `.streamlit/secrets.toml` file — keep it private!

### Step 2 — Deploy on Streamlit Cloud
1. Go to share.streamlit.io
2. Sign in with GitHub
3. Click "New app"
4. Select your repo → main file: `app.py`
5. Click "Advanced settings" → "Secrets"
6. Paste this into the secrets box:
   ```
   SUPABASE_URL = "https://pvsgqnmsmdioawuehwmr.supabase.co"
   SUPABASE_KEY = "your_supabase_key_here"
   ```
7. Click "Deploy"

### Step 3 — Share the link
Copy the URL and share with your colleagues. Done! ✅

---

## Running Locally (optional)
```bash
pip install -r requirements.txt
streamlit run app.py
```
Make sure `.streamlit/secrets.toml` exists with your credentials.

---

## Files
```
failure_tracker_supabase/
├── app.py                      # Main app
├── requirements.txt            # Dependencies
├── README.md                   # This file
└── .streamlit/
    └── secrets.toml            # Credentials (DO NOT upload to GitHub)
```
