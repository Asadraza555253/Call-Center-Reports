# CX AI Call Center Reporting Agent — Web App

A Streamlit web app for the CX daily call-center report you provided.

## Features
- Upload Queue Report CSV/XLS/XLSX
- Upload Agent Summary CSV/XLS/XLSX
- Daily agent dashboard
- Morning / Evening / Night shift analysis
- AHT and handling-time buckets
- IVR waiting-time analysis
- IVR hangups separated from real abandoned calls
- Within-1-minute and within-2-minute KPIs
- Agents above AHT target
- Interactive charts
- AI management summary (optional OpenAI API)
- Built-in report Q&A for common questions
- Excel report export

## Run locally

```bash
python -m venv .venv
# Windows
.venv\\Scripts\\activate
# macOS/Linux
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Open the URL shown by Streamlit, normally http://localhost:8501.

## Optional AI API

Set `OPENAI_API_KEY` in your environment. You can also set `OPENAI_MODEL` to the model available in your OpenAI account. Without a key, the app still generates deterministic management insights.

## Deploy

This folder can be deployed to Streamlit Community Cloud or another service that runs Streamlit. The entry point is `app.py`.
