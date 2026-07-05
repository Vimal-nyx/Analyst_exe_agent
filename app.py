import streamlit as st
import pandas as pd
import google.generativeai as genai
import os
import io
import re
from openpyxl.styles import Font, PatternFill

# Set page config
st.set_page_config(
    page_title="ANALYST.EXE",
    page_icon="📟",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Custom CSS to force terminal style with small fonts and a distinct upload border
st.markdown("""
<style>
/* Hide the Streamlit header, footer, decoration, and hamburger menu */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
.stAppHeader {display: none !important;}
.stAppFooter {display: none !important;}
[data-testid="stHeader"] {display: none !important;}
[data-testid="stToolbar"] {display: none !important;}
[data-testid="stDecoration"] {display: none !important;}

/* Force Background to Black and Text to Terminal Green Monospace with compact/small font sizes */
html, body, [data-testid="stAppViewContainer"], [data-testid="stMain"], .stApp {
    background-color: #000000 !important;
    color: #00ff00 !important;
    font-family: 'Courier New', Courier, monospace !important;
    font-size: 14px !important;
}

/* Ensure all text matches terminal green and compact size */
h1, h2, h3, h4, h5, h6, p, li, span, label, div {
    color: #00ff00 !important;
    font-family: 'Courier New', Courier, monospace !important;
    font-size: 14px !important;
}

/* Headers style */
h1, h2, h3 {
    font-size: 16px !important;
    font-weight: bold !important;
}

/* Set block container padding */
.block-container {
    padding-top: 1.5rem !important;
    padding-bottom: 1.5rem !important;
    max-width: 800px !important;
}

/* Style TextInput component */
[data-testid="stTextInput"] input {
    background-color: #000000 !important;
    color: #00ff00 !important;
    border: none !important;
    border-bottom: 2px solid #00ff00 !important;
    border-radius: 0px !important;
    font-family: 'Courier New', Courier, monospace !important;
    font-size: 14px !important;
    padding-left: 0px !important;
    caret-color: #00ff00 !important;
}

[data-testid="stTextInput"] input:focus {
    outline: none !important;
    box-shadow: none !important;
    border-bottom: 2px solid #00ff00 !important;
}

/* Style File Uploader component */
[data-testid="stFileUploader"] {
    background-color: #000000 !important;
    border: none !important;
    padding: 0px !important;
}

[data-testid="stFileUploader"] section {
    background-color: #000000 !important;
    border: none !important;
}

/* Hide standard file uploader texts/button */
[data-testid="stFileUploadDropzone"] > div {
    display: none !important;
}

/* Inject minimal bracketed text instead */
[data-testid="stFileUploadDropzone"]::before {
    content: "[ + ADD EXCEL ]" !important;
    color: #00ff00 !important;
    font-family: 'Courier New', Courier, monospace !important;
    font-size: 14px !important;
    display: block !important;
    padding: 20px 0 !important;
    cursor: pointer !important;
}

/* Distinct, visible dashed terminal border */
[data-testid="stFileUploadDropzone"] {
    background-color: #000000 !important;
    border: 2px dashed #00ff00 !important;
    border-radius: 0px !important;
    padding: 15px !important;
    text-align: center !important;
    cursor: pointer !important;
}

[data-testid="stFileUploadDropzone"]:hover {
    border-color: #ffffff !important;
}

/* Style the uploaded file details card */
[data-testid="stUploadedFile"] {
    display: flex !important;
    align-items: center !important;
    justify-content: space-between !important;
    background-color: #000000 !important;
    border: 1px solid #00ff00 !important;
    color: #00ff00 !important;
    border-radius: 0px !important;
    font-family: 'Courier New', Courier, monospace !important;
    padding: 8px 12px !important;
    margin-top: 10px !important;
}

/* Ensure inner elements and text wrap cleanly */
[data-testid="stUploadedFile"] > div {
    display: flex !important;
    flex-direction: column !important;
    gap: 4px !important;
    height: auto !important;
    overflow: hidden !important;
}

[data-testid="stUploadedFile"] * {
    color: #00ff00 !important;
    font-family: 'Courier New', Courier, monospace !important;
    font-size: 13px !important;
    white-space: normal !important;
    word-break: break-all !important;
    line-height: 1.3 !important;
}

[data-testid="stUploadedFile"] svg {
    fill: #00ff00 !important;
    color: #00ff00 !important;
    flex-shrink: 0 !important;
}

[data-testid="stUploadedFile"] button {
    background-color: transparent !important;
    border: none !important;
    color: #00ff00 !important;
    cursor: pointer !important;
    flex-shrink: 0 !important;
    font-family: 'Courier New', Courier, monospace !important;
}

/* Style Execute and Download Buttons to look like terminal options */
.stButton button, [data-testid="stDownloadButton"] button {
    background-color: transparent !important;
    color: #00ff00 !important;
    border: 1px solid #00ff00 !important;
    border-radius: 0px !important;
    font-family: 'Courier New', Courier, monospace !important;
    font-size: 13px !important;
    padding: 6px 16px !important;
    cursor: pointer !important;
    text-transform: uppercase !important;
}

.stButton button:hover, [data-testid="stDownloadButton"] button:hover {
    background-color: #00ff00 !important;
    color: #000000 !important;
    border: 1px solid #00ff00 !important;
}

.stButton button:active, [data-testid="stDownloadButton"] button:active {
    background-color: #008800 !important;
    color: #000000 !important;
}

/* Code and preformatted blocks style */
code, pre {
    background-color: #111111 !important;
    color: #00ff00 !important;
    border: 1px solid #00ff00 !important;
    border-radius: 0px !important;
    font-family: 'Courier New', Courier, monospace !important;
    font-size: 13px !important;
}
</style>
""", unsafe_allow_html=True)

# Helper to print terminal style error logs
def term_print(text, type="info"):
    if type == "error":
        color = "#ff3333"
        prefix = "[-]"
    else:
        color = "#00ff00"
        prefix = "[+]"
    st.markdown(f"<pre style='color:{color}; background:black; border:none; padding:0; margin:0;'>{prefix} {text}</pre>", unsafe_allow_html=True)

# Helper to sanitize instructions into a clean filename matching User specification
def sanitize_filename(instruction, columns):
    # Lowercase columns for matching
    col_lower = [c.lower() for c in columns]
    col_map = {c.lower(): c for c in columns} # maps lowercase to original casing
    
    # Translate operators
    s = instruction
    s = s.replace(">=", " gte ")
    s = s.replace("<=", " lte ")
    s = s.replace(">", " gt ")
    s = s.replace("<", " lt ")
    s = s.replace("==", " eq ")
    s = s.replace("=", " eq ")
    s = s.replace("!=", " neq ")
    
    # Remove all other non-alphanumeric characters
    s = re.sub(r'[^a-zA-Z0-9]', ' ', s)
    words = s.split()
    
    # Filter stopwords
    stopwords = {"filter", "select", "get", "show", "with", "a", "an", "the", "for", "to", "in", "on", "at", "by", "of", "and", "where", "find", "list"}
    filtered = []
    for w in words:
        if w.lower() not in stopwords:
            # Restore original column casing if matching
            if w.lower() in col_map:
                filtered.append(col_map[w.lower()])
            else:
                filtered.append(w)
                
    # Reordering logic: [operator] [value] [column] -> [column] [operator] [value]
    # Example: ['students', 'gt', '9', 'CGPA'] -> ['students', 'CGPA', 'gt', '9']
    i = 0
    while i < len(filtered) - 2:
        op = filtered[i].lower()
        val = filtered[i+1]
        col = filtered[i+2]
        if op in {"gt", "lt", "gte", "lte", "eq", "neq"} and val.isdigit() and col.lower() in col_lower:
            filtered[i], filtered[i+1], filtered[i+2] = col, op, val
            i += 3
        else:
            i += 1
            
    # Capitalize the first word
    if filtered:
        filtered[0] = filtered[0].capitalize()
        
    filename = "_".join(filtered)
    
    if not filename:
        filename = "result"
        
    return f"{filename}.xlsx"

# Terminal Banner
st.markdown("""
<pre style="color:#00ff00; background:black; border:none; padding:0; line-height:1.2; font-weight:bold;">
============================================================
              A N A L Y S T . E X E  (v1.0)
   STATELESS NATURAL LANGUAGE DATA MANIPULATION UNIT
============================================================
</pre>
""", unsafe_allow_html=True)

# Securely fetch API key
api_key = os.environ.get("GEMINI_API_KEY")

if not api_key:
    term_print("SYSTEM ERROR: ENVIRONMENT VARIABLE 'GEMINI_API_KEY' NOT DETECTED.", type="error")
    term_print("OPERATION TERMINATED. CONFIGURE ENVIRONMENT AND RESTART APPLICATION.", type="error")
    st.stop()

# Initialize Session State variables for results caching to prevent UI reset on download click
if 'result_df' not in st.session_state:
    st.session_state.result_df = None
if 'ai_code' not in st.session_state:
    st.session_state.ai_code = None
if 'out_filename' not in st.session_state:
    st.session_state.out_filename = None
if 'error_msg' not in st.session_state:
    st.session_state.error_msg = None
if 'last_file' not in st.session_state:
    st.session_state.last_file = None

# 1. Upload Section (xlsx only)
uploaded_file = st.file_uploader("SOURCE_FILE", type=["xlsx"], label_visibility="collapsed")

if uploaded_file is not None:
    # Reset session state if a new file is uploaded
    if st.session_state.last_file != uploaded_file.name:
        st.session_state.last_file = uploaded_file.name
        st.session_state.result_df = None
        st.session_state.ai_code = None
        st.session_state.out_filename = None
        st.session_state.error_msg = None

    try:
        df = pd.read_excel(uploaded_file)
        columns = list(df.columns)
        
        # 2. Metadata Block (placed before instruction input and execute button)
        st.markdown(f"<pre style='color:#00ff00; background:black; border:none; padding:0; margin:0;'>[DATA_LOADED: {df.shape[0]} rows * {df.shape[1]} columns]</pre>", unsafe_allow_html=True)
        st.markdown(f"**COLUMNS DETECTED:** {', '.join(columns)}")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # 3. Input Instruction Section
        instruction = st.text_input("INSTRUCTION_INPUT", label_visibility="collapsed", placeholder="ENTER COMMAND (e.g. Filter students > 9 CGPA)")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # 4. Execute Action Button
        if st.button("[ EXECUTE ]"):
            if not instruction:
                st.session_state.error_msg = "ERROR: Instruction command is empty."
                st.session_state.result_df = None
                st.session_state.ai_code = None
                st.session_state.out_filename = None
            else:
                st.session_state.error_msg = None
                st.session_state.result_df = None
                st.session_state.ai_code = None
                st.session_state.out_filename = None
                
                # Configure Generative AI
                genai.configure(api_key=api_key)
                
                # Construct strict system prompt
                system_prompt = (
                    "You are an expert pandas compiler. Your only output is raw, execution-ready Python pandas code.\n"
                    f"The input DataFrame is named 'df' and contains the following columns: {columns}.\n"
                    f"The user wants to perform this action: '{instruction}'.\n"
                    "Rules:\n"
                    "1. Refer to the input DataFrame strictly as 'df'.\n"
                    "2. Make sure you apply the filters properly using Boolean indexing. For example, if user asks to 'Filter CGPA > 9', your code MUST perform the filter like df[df['CGPA'] > 9] or df.query('CGPA > 9'). Do NOT just return columns without applying filters.\n"
                    "3. You MUST save the final resulting output to a variable named 'result_df'.\n"
                    "4. CRITICAL: 'result_df' MUST be a pandas DataFrame, not a Series. If the resulting operation produces a Series (e.g. selecting a single column like df['Age']), convert it to a DataFrame using `df[['Age']]` or `.to_frame()` to preserve the column names and headers.\n"
                    "5. Do NOT write any markdown blocks (such as ```python ... ```), explanations, HTML, comments, or extra text.\n"
                    "6. Your output MUST be 100% executable Python code."
                )
                
                try:
                    # Call gemini-2.5-flash model
                    model = genai.GenerativeModel(
                        model_name="gemini-2.5-flash",
                        system_instruction="You are a pandas code compiler. You output ONLY raw executable python code. No comments, no explanations, no markdown tags. Assign output to result_df and ensure it is a DataFrame."
                    )
                    
                    response = model.generate_content(system_prompt)
                    ai_code = response.text.strip()
                    
                    # Strip markdown block wrappers if present
                    if ai_code.startswith("```python"):
                        ai_code = ai_code[len("```python"):].strip()
                    elif ai_code.startswith("```"):
                        ai_code = ai_code[len("```"):].strip()
                    if ai_code.endswith("```"):
                        ai_code = ai_code[:-3].strip()
                    
                    # Isolated execution
                    local_vars = {
                        'df': df.copy(),
                        'pd': pd
                    }
                    
                    exec(ai_code, {}, local_vars)
                    
                    if 'result_df' in local_vars:
                        result_df = local_vars['result_df']
                        
                        # Fallback convert Series to DataFrame
                        if isinstance(result_df, pd.Series):
                            col_name = result_df.name if result_df.name else "result"
                            result_df = result_df.to_frame(name=col_name)
                        
                        if isinstance(result_df, pd.DataFrame):
                            st.session_state.result_df = result_df
                            st.session_state.ai_code = ai_code
                            st.session_state.out_filename = sanitize_filename(instruction, columns)
                        else:
                            st.session_state.error_msg = f"RUNTIME ERROR: 'result_df' is of type {type(result_df)}, expected pandas.DataFrame"
                    else:
                        st.session_state.error_msg = "RUNTIME ERROR: Variable 'result_df' was not populated by the compiled code."
                        st.session_state.ai_code = ai_code
                except Exception as e:
                    st.session_state.error_msg = f"COMPILER/RUNTIME EXCEPTION: {str(e)}"

        # 5. Output Result Area (Clean terminal layout: no status/compilation log outputs)
        if st.session_state.error_msg:
            term_print(st.session_state.error_msg, type="error")
            if st.session_state.ai_code:
                st.markdown("<pre style='color:#ff3333; background:black; border:none; padding:0; margin:0;'>[COMPILED CODE FOR DEBUGGING]</pre>", unsafe_allow_html=True)
                st.code(st.session_state.ai_code, language="python")
                
        elif st.session_state.result_df is not None:
            st.markdown("<pre style='color:#00ff00; background:black; border:none; padding:0; margin:0;'>[CODE_GENERATED]</pre>", unsafe_allow_html=True)
            st.code(st.session_state.ai_code, language="python")
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("<pre style='color:#00ff00; background:black; border:none; padding:0; margin:0;'>[PREVIEW_RESULT]</pre>", unsafe_allow_html=True)
            st.code(st.session_state.result_df.head().to_string(index=False), language="text")
            
            # Format Excel Output with explicitly Bold Headers and accent Background Fill
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                st.session_state.result_df.to_excel(writer, index=False)
                
                workbook = writer.book
                worksheet = writer.sheets[list(writer.sheets.keys())[0]]
                
                # Bold headers + Subtle light-gray background to distinguish it from values
                header_font = Font(name='Courier New', size=11, bold=True, color='000000')
                header_fill = PatternFill(start_color='E0E0E0', end_color='E0E0E0', fill_type='solid')
                
                for col_num in range(1, len(st.session_state.result_df.columns) + 1):
                    cell = worksheet.cell(row=1, column=col_num)
                    cell.font = header_font
                    cell.fill = header_fill
                    
            excel_data = buffer.getvalue()
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            # Dynamic download button
            st.download_button(
                label=f"[ DOWNLOAD: {st.session_state.out_filename.upper()} ]",
                data=excel_data,
                file_name=st.session_state.out_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        term_print(f"INPUT READING ERROR: {str(e)}", type="error")