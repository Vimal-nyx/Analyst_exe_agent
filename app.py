import streamlit as st
import pandas as pd
import google.generativeai as genai
import os
import io
import re

# Set page config to collapsed/minimal sidebar and terminal page title
st.set_page_config(
    page_title="ANALYST.EXE",
    page_icon="📟",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Custom CSS to force terminal green/black style and hide Streamlit elements completely
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

/* Force Background to Black and Text to Terminal Green Monospace */
html, body, [data-testid="stAppViewContainer"], [data-testid="stMain"], .stApp {
    background-color: #000000 !important;
    color: #00ff00 !important;
    font-family: 'Courier New', Courier, monospace !important;
}

/* Override all headers, paragraphs, lists, spans, and divs to be green monospace */
h1, h2, h3, h4, h5, h6, p, li, span, label, div {
    color: #00ff00 !important;
    font-family: 'Courier New', Courier, monospace !important;
}

/* Set block container padding and width */
.block-container {
    padding-top: 2rem !important;
    padding-bottom: 2rem !important;
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
    font-size: 16px !important;
    padding-left: 0px !important;
    caret-color: #00ff00 !important;
}

[data-testid="stTextInput"] input:focus {
    outline: none !important;
    box-shadow: none !important;
    border-bottom: 2px solid #00ff00 !important;
}

[data-testid="stTextInput"] label {
    color: #00ff00 !important;
    font-family: 'Courier New', Courier, monospace !important;
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
    font-size: 16px !important;
    display: block !important;
    padding: 20px 0 !important;
    cursor: pointer !important;
}

[data-testid="stFileUploadDropzone"] {
    background: transparent !important;
    border: 1px dashed #00ff00 !important;
    border-radius: 0px !important;
    padding: 10px !important;
    text-align: center !important;
    cursor: pointer !important;
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
    padding: 12px 16px !important;
    margin-top: 15px !important;
}

/* Ensure inner elements and text wrap cleanly without height constraints */
[data-testid="stUploadedFile"] > div {
    display: flex !important;
    flex-direction: column !important;
    gap: 4px !important;
    height: auto !important;
    overflow: hidden !important;
}

/* Style filenames and other texts inside the file details card */
[data-testid="stUploadedFile"] * {
    color: #00ff00 !important;
    font-family: 'Courier New', Courier, monospace !important;
    white-space: normal !important;
    word-break: break-all !important;
    line-height: 1.4 !important;
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
    padding: 8px 20px !important;
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
}
</style>
""", unsafe_allow_html=True)

# Helper to print terminal style logs
def term_print(text, type="info"):
    if type == "error":
        color = "#ff3333"
        prefix = "[-]"
    elif type == "success":
        color = "#00ff00"
        prefix = "[+]"
    else:
        color = "#00ff00"
        prefix = "[+]"
    st.markdown(f"<pre style='color:{color}; background:black; border:none; padding:0; margin:0;'>{prefix} {text}</pre>", unsafe_allow_html=True)

# Helper to sanitize instructions into a clean filename
def sanitize_filename(instruction):
    # Convert operators to words
    s = instruction.lower()
    s = s.replace(">=", " gte ")
    s = s.replace("<=", " lte ")
    s = s.replace(">", " gt ")
    s = s.replace("<", " lt ")
    s = s.replace("==", " eq ")
    s = s.replace("=", " eq ")
    s = s.replace("!=", " neq ")
    
    # Remove all other non-alphanumeric characters, replacing them with spaces
    s = re.sub(r'[^a-zA-Z0-9]', ' ', s)
    
    # Split into words
    words = s.split()
    
    # Filter stopwords
    stopwords = {"filter", "select", "get", "show", "with", "a", "an", "the", "for", "to", "in", "on", "at", "by", "of", "and", "where", "find", "list"}
    filtered_words = [w for w in words if w not in stopwords]
    
    # Fallback to original words if all were stopwords
    if not filtered_words:
        filtered_words = words
        
    # Join with underscores
    filename = "_".join(filtered_words)
    
    # If empty, default to "result"
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

# Initialize Session State variables for result caching to avoid UI loss on download click
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

# File Uploader - Placement at the very top of application inputs
uploaded_file = st.file_uploader("SOURCE_FILE", type=["xlsx"], label_visibility="collapsed")

if uploaded_file is not None:
    # Clear session state if a new file is uploaded
    if st.session_state.last_file != uploaded_file.name:
        st.session_state.last_file = uploaded_file.name
        st.session_state.result_df = None
        st.session_state.ai_code = None
        st.session_state.out_filename = None
        st.session_state.error_msg = None

    try:
        # Load excel sheet into memory
        df = pd.read_excel(uploaded_file)
        columns = list(df.columns)
        
        # User Instruction Input (Follows uploader)
        instruction = st.text_input("INSTRUCTION_INPUT", label_visibility="collapsed", placeholder="ENTER COMMAND (e.g. Filter CGPA > 9)")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Data Metadata Summary (Dimensions and columns list, placed before Execute button)
        st.markdown(f"<pre style='color:#00ff00; background:black; border:none; padding:0; margin:0;'>[DATA_LOADED: {df.shape[0]} rows * {df.shape[1]} columns]</pre>", unsafe_allow_html=True)
        st.markdown(f"<pre style='color:#00ff00; background:black; border:none; padding:0; margin:0;'>Columns: {', '.join(columns)}</pre>", unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Execute button
        if st.button("[ EXECUTE ]"):
            if not instruction:
                st.session_state.error_msg = "ERROR: Instruction command is empty."
                st.session_state.result_df = None
                st.session_state.ai_code = None
                st.session_state.out_filename = None
            else:
                # Clear previous outputs
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
                    
                    # Strip markdown if model ignored instructions
                    if ai_code.startswith("```python"):
                        ai_code = ai_code[len("```python"):].strip()
                    elif ai_code.startswith("```"):
                        ai_code = ai_code[len("```"):].strip()
                    if ai_code.endswith("```"):
                        ai_code = ai_code[:-3].strip()
                    
                    # Setup isolated environment
                    local_vars = {
                        'df': df.copy(),
                        'pd': pd
                    }
                    
                    # Executing generated code
                    exec(ai_code, {}, local_vars)
                    
                    # Extract and process result_df
                    if 'result_df' in local_vars:
                        result_df = local_vars['result_df']
                        
                        # Convert to DataFrame if a Series is returned as safety fallback (ensures column header is preserved)
                        if isinstance(result_df, pd.Series):
                            col_name = result_df.name if result_df.name else "result"
                            result_df = result_df.to_frame(name=col_name)
                        
                        if isinstance(result_df, pd.DataFrame):
                            # Store outputs in session state
                            st.session_state.result_df = result_df
                            st.session_state.ai_code = ai_code
                            st.session_state.out_filename = sanitize_filename(instruction)
                        else:
                            st.session_state.error_msg = f"RUNTIME ERROR: 'result_df' is of type {type(result_df)}, expected pandas.DataFrame"
                    else:
                        st.session_state.error_msg = "RUNTIME ERROR: Variable 'result_df' was not populated by the compiled code."
                        st.session_state.ai_code = ai_code  # Save code for debugging
                except Exception as e:
                    st.session_state.error_msg = f"COMPILER/RUNTIME EXCEPTION: {str(e)}"
                    
        # Render execution outputs from session state (Clean layout with only requested elements)
        if st.session_state.error_msg:
            term_print(st.session_state.error_msg, type="error")
            if st.session_state.ai_code:
                st.markdown("<pre style='color:#ff3333; background:black; border:none; padding:0; margin:0;'>[COMPILED CODE FOR DEBUGGING]</pre>", unsafe_allow_html=True)
                st.code(st.session_state.ai_code, language="python")
                
        elif st.session_state.result_df is not None:
            # Result Output Area (Only code, preview, and download)
            st.markdown("<pre style='color:#00ff00; background:black; border:none; padding:0; margin:0;'>[CODE_GENERATED]</pre>", unsafe_allow_html=True)
            st.code(st.session_state.ai_code, language="python")
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("<pre style='color:#00ff00; background:black; border:none; padding:0; margin:0;'>[PREVIEW_RESULT]</pre>", unsafe_allow_html=True)
            st.code(st.session_state.result_df.head().to_string(index=False), language="text")
            
            # Generate bytes buffer for download
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                # header=True is implicitly active, ensuring single column retains header and is not mistaken for raw data
                st.session_state.result_df.to_excel(writer, index=False)
            excel_data = buffer.getvalue()
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            # Download button styled as terminal option with dynamic filename
            st.download_button(
                label=f"[ DOWNLOAD: {st.session_state.out_filename.upper()} ]",
                data=excel_data,
                file_name=st.session_state.out_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        term_print(f"INPUT READING ERROR: {str(e)}", type="error")