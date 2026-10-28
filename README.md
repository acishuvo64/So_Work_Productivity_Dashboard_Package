# SO Productivity Dashboard

## Folder structure (GitHub repo)
```
dashboard.py
requirements.txt
.streamlit/config.toml      <- hidden folder, must be uploaded (it hides Edit/Deploy buttons)
data/<your excel file>.xlsx <- 3 sheets: Summary Report, SO Activity, Attendance
assets/badge.png, logo.png
```

## Deploy
1. Upload everything above to a GitHub repo (use a **Private** repo so nobody can read the code).
2. share.streamlit.io -> New app -> select repo -> main file `dashboard.py` -> Deploy.
3. Send the app link to the Director.

## Daily / monthly update
Put the new Excel in `data/` (delete the old one, or keep it - the newest file is used) and push to GitHub.
The dashboard refreshes automatically in 1-2 minutes.

## Settings (top of dashboard.py)
Daily targets (20 outlets / 8 memos), 11 working hours, 8:45 late time, score weights.
