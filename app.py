import io, os, re, json, shutil, subprocess, tempfile
from datetime import datetime
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
from openai import OpenAI
except Exception:
OpenAI = None

st.set_page_config(page_title="CX AI Reporting Agent", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

DEFAULTS = {"aht":90, "within1":98, "within2":98, "abandon":2, "exclude":["Counter Staff"]}

# ---------- Helpers ----------
def clean_name(x):
return re.sub(r"\s+", " ", str(x).strip()) if pd.notna(x) else ""

def parse_hms(x):
if pd.isna(x): return 0.0
s=str(x).strip()
try:
p=s.split(":")
if len(p)==3: return int(p[0])*3600+int(p[1])*60+float(p[2])
if len(p)==2: return int(p[0])*60+float(p[1])
return float(s)
except: return 0.0

def hms(sec):
sec=max(0,int(round(float(sec or 0))))
return f"{sec//3600:02d}:{(sec%3600)//60:02d}:{sec%60:02d}"

def read_table(upload):
"""Read CSV/XLSX/XLS exports.

Some Intellicon/legacy BIFF .xls exports contain harmless BIFF records that
newer xlrd versions flag as workbook corruption (e.g. "seen[2] == 4").
We first ask xlrd to tolerate those records, then fall back to LibreOffice
conversion when available.
"""
b=upload.getvalue(); n=upload.name.lower()
if n.endswith('.csv'):
return pd.read_csv(io.BytesIO(b), low_memory=False)
if n.endswith('.xlsx'):
return pd.read_excel(io.BytesIO(b), engine='openpyxl')
if n.endswith('.xls'):
# Primary path: xlrd supports legacy .xls and can ignore benign
# workbook-corruption markers produced by some call-center exports.
