import streamlit as st
import pandas as pd
import google.generativeai as genai
import os
import io

# Set page config to collapsed/minimal sidebar and terminal page title
st.set_page_config(
    page_title="ANALYST.EXE",
    page_icon="📟",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Custom CSS to force terminal green/black style and hide Streamlit elements
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
    border: 1px dashed #00ff00 !important;
    border-radius: 0px !important;
    padding: 10px !important;
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
    background-color: #000000 !important;
    border: 1px solid #00ff00 !important;
    color: #00ff00 !important;
    border-radius: 0px !important;
    font-family: 'Courier New', Courier, monospace !important;
}

[data-testid="stUploadedFile"] * {
    color: #00ff00 !important;
    font-family: 'Courier New', Courier, monospace !important;
}

[data-testid="stUploadedFile"] svg {
    fill: #00ff00 !important;
    color: #00ff00 !important;
}

[data-testid="stUploadedFile"] button {
    background-color: transparent !important;
    border: none !important;
}

/* Style Download Button to look like terminal option */
[data-testid="stDownloadButton"] button {
    background-color: transparent !important;
    color: #00ff00 !important;
    border: 1px solid #00ff00 !important;
    border-radius: 0px !important;
    font-family: 'Courier New', Courier, monospace !important;
    padding: 8px 20px !important;
    cursor: pointer !important;
    text-transform: uppercase;
}

[data-testid="stDownloadButton"] button:hover {
    background-color: #00ff00 !important;
    color: #000000 !important;
    border: 1px solid #00ff00 !important;
}

[data-testid="stDownloadButton"] button:active {
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

/* Terminal blink cursor simulation */
.cursor {
    animation: blinker 1s linear infinite;
}
@keyframes blinker {
    50% { opacity: 0; }
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
    term_print("OPERATION TERMINATED. CONFIGURE ENVIRONMENT AND RESTART APPLICANT.", type="error")
    st.stop()
else:
    term_print("GEMINI_API_KEY DETECTED SECURELY.")
    term_print("CORE ENGINE STATUS: ONLINE.")

st.markdown("<br>", unsafe_allow_html=True)

# File Uploader
term_print("AWAITING DATA INPUT FILE (.XLSX):")
uploaded_file = st.file_uploader("SOURCE_FILE", type=["xlsx"], label_visibility="collapsed")

if uploaded_file is not None:
    try:
        # Load excel sheet into memory
        df = pd.read_excel(uploaded_file)
        columns = list(df.columns)
        
        term_print(f"FILE LOADED: {uploaded_file.name}", type="success")
        term_print(f"COLUMNS DETECTED: {str(columns)}")
        term_print(f"DIMENSIONS: {df.shape[0]} rows x {df.shape[1]} columns")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # User Instruction Input
        term_print("ENTER NATURAL LANGUAGE INSTRUCTION:")
        instruction = st.text_input("INSTRUCTION_INPUT", label_visibility="collapsed", placeholder="e.g. Filter CGPA > 9, sort by Name")
        
        if instruction:
            term_print("GENERATING CODE VIA GEMINI ENGINE...")
            
            # Configure Generative AI
            genai.configure(api_key=api_key)
            
            # Construct strict system prompt
            system_prompt = (
                "You are an expert pandas compiler. Your only output is raw, execution-ready Python pandas code.\n"
                f"The input DataFrame is named 'df' and contains the following columns: {columns}.\n"
                f"The user wants to perform this action: '{instruction}'.\n"
                "You MUST save the final resulting DataFrame to a variable named 'result_df'.\n"
                "Rules:\n"
                "1. Refer to the input DataFrame strictly as 'df'.\n"
                "2. Do NOT write any markdown blocks (such as ```python ... ```), explanations, HTML, comments, or extra text.\n"
                "3. Your output MUST be 100% executable Python code.\n"
                "4. Make sure 'result_df' is populated with the final DataFrame."
            )
            
            try:
                # Call gemini-2.5-flash model
                model = genai.GenerativeModel(
                    model_name="gemini-2.5-flash",
                    system_instruction="You are a pandas code compiler. You output ONLY raw executable python code. No comments, no explanations, no markdown tags. Assign output to result_df."
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
                
                term_print("COMPILATION SUCCESSFUL. RUNNING IN ISOLATED RUNTIME:")
                st.code(ai_code, language="python")
                
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
                    
                    if isinstance(result_df, pd.DataFrame):
                        term_print(f"EXECUTION SUCCESSFUL. RESULT DATASET SHAPE: {result_df.shape[0]} rows x {result_df.shape[1]} columns", type="success")
                        
                        st.markdown("<br>", unsafe_allow_html=True)
                        term_print("PREVIEWING RESULT (FIRST 5 ROWS):")
                        st.code(result_df.head().to_string(index=False), language="text")
                        
                        # Generate bytes buffer for download
                        buffer = io.BytesIO()
                        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                            result_df.to_excel(writer, index=False)
                        excel_data = buffer.getvalue()
                        
                        st.markdown("<br>", unsafe_allow_html=True)
                        term_print("DOWNLOAD READY:")
                        
                        # Download button styled as terminal option
                        st.download_button(
                            label="[ DOWNLOAD_RESULT.XLSX ]",
                            data=excel_data,
                            file_name="result.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        term_print(f"RUNTIME ERROR: 'result_df' is of type {type(result_df)}, expected pandas.DataFrame", type="error")
                else:
                    term_print("RUNTIME ERROR: Variable 'result_df' was not populated by the compiled code.", type="error")
                    
            except Exception as e:
                term_print(f"COMPILER/RUNTIME EXCEPTION: {str(e)}", type="error")
                
    except Exception as e:
        term_print(f"INPUT READING ERROR: {str(e)}", type="error")
