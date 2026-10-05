# Construction Projects Dashboard
Upload an .xlsx file -> a button is created for every sheet -> click to open a dashboard.

Run locally:
    pip install -r requirements.txt
    streamlit run app.py

Deploy: push to GitHub, then share.streamlit.io -> New app -> select repo, branch main, main file `app.py`.
Edit COLUMN_MAP at the top of app.py to match your Excel column names.
