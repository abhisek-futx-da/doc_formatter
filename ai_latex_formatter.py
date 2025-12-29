import re
import requests
import os
import sys
from pathlib import Path

# --- USER CONFIGURATION ---
# 1. Please update these paths before running the script.

# Path to the input LaTeX file that needs table correction.
INPUT_LATEX_FILE = 'output/corrected.tex'

# Path where the corrected LaTeX file will be saved.
OUTPUT_LATEX_FILE = 'output/final_corrected.tex'

# 2. Get a free API key from Google AI Studio (https://aistudio.google.com/)
#    and paste it here.
GEMINI_API_KEY = "AIzaSyD_AvxiVPlNccIo1dYtEqmid83H0uswFl8"
# -----------------------------


def get_corrected_table_from_api(raw_table_code: str) -> str:
    """
    Sends raw LaTeX table code to the Google Gemini Free API and returns the corrected version.

    Args:
        raw_table_code: A string containing the messy LaTeX table code.

    Returns:
        A string containing the corrected and formatted LaTeX table code.
    """
    if not GEMINI_API_KEY or GEMINI_API_KEY == "<YOUR_GEMINI_API_KEY>":
        print(" ERROR: Gemini API key is not set. Skipping table correction.")
        return raw_table_code

    API_ENDPOINT = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key={GEMINI_API_KEY}"
    
    prompt = f"""
    You are an expert LaTeX assistant. Your task is to correct and reformat the following raw LaTeX table code.
    
    **Instructions:**
    1.  The final output must be a clean, syntactically correct LaTeX table.
    2.  Format the table using the 'booktabs' package style, which means using `\\toprule`, `\\midrule`, and `\\bottomrule`. Do not use vertical lines.
    3.  Remove any extraneous or problematic LaTeX commands, such as `\\begin{{longtable}}`, `\\begin{{minipage}}`, `\\endhead`, etc. The final table should be in a `tabular` environment.
    4.  Ensure all rows have a consistent number of columns, using '&' as a separator.
    5.  Do NOT include a `\\caption{{...}}` or `\\label{{...}}`. You must only output the table structure itself.
    6.  The final output must ONLY be the LaTeX code from `\\begin{{tabular}}` to `\\end{{tabular}}`. Do not add any explanations, surrounding text, or markdown code fences like ```latex.

    **Raw LaTeX table to correct:**
    {raw_table_code}
    """
    
    headers = {"Content-Type": "application/json"}
    
    
    
    