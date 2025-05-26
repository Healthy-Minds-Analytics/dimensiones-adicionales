import os
import pandas as pd
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import glob
from datetime import datetime
import json
import random
from copy import deepcopy
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def load_csv(source) -> pd.DataFrame:
    """
    Lee un CSV desde un path o desde un file-like (Streamlit u otros).
    """
    if hasattr(source, "read"):
        return pd.read_csv(source)
    return pd.read_csv(source)

def load_json(source) -> dict:
    """
    Lee un JSON desde un path o desde un file-like.
    """
    if hasattr(source, "read"):
        return json.load(source)
    with open(source, encoding="utf-8") as f:
        return json.load(f)
    
def df_a_reemplazos(df_stats: pd.DataFrame) -> dict:
    reemplazos = {}
    for dim, row in df_stats.iterrows():
        reemplazos[f"MEDIA_{dim}"] = row['mean']
        reemplazos[f"STD_{dim}"] = row['std']
    return reemplazos