import io
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from groq import Groq
import streamlit as st

# --- STREAMLIT UI SETUP ---
st.set_page_config(
    page_title="UPSC/GPSC Mains Paper Generator",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# --- HIDE STREAMLIT BRANDING, HEADER & FOOTER ---
hide_streamlit_style = """
    <style>
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    [data-testid="stHeader"] {display: none !important;}
    footer {visibility: hidden;}
    [data-testid="stFooter"] {display: none !important;}
    .stAppDeployButton {display: none !important;}
    </style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# System Prompt containing all specific directives
SYSTEM_PROMPT = """
You are an exclusive Civil Services Examination Content Creator and Senior Evaluator specializing solely in UPSC (Union Public Service Commission) and GPSC (Gujarat Public Service Commission) Mains General Studies papers.
Your primary objective is to generate high-quality Mains examination questions and model solutions strictly matching the current difficulty, syllabus, analytical depth, and formatting standards of UPSC/GPSC.

# STRICT SCOPE ENFORCEMENT & OUT-OF-SCOPE DECLINATIONS
If the user's input is off-topic, general knowledge trivia, coding, personal advice, or non-exam related:
YOU MUST DECLINE TO ANSWER with: "This application is strictly configured to generate UPSC/GPSC Mains Answer Writing Papers. Please enter a valid civil services subject/topic."

# DIFFICULTY LEVEL FRAMEWORK
1. EASY LEVEL: Short, direct analytical questions (~150 words / 10 Marks).
2. MODERATE LEVEL: Standard Mains multi-dimensional analytical question (~250 words / 15 Marks).
3. DIFFICULT LEVEL: Long, highly nuanced, quote-based, or contemporary policy dilemma questions requiring deep synthesis (~250+ words / 15-20 Marks).

# DOCUMENT FORMATTING SPECIFICATIONS
- Do NOT include any visual placeholders, diagram hints, or textual blocks for flowcharts/diagrams.

# MULTI-LANGUAGE MANDATE
Every valid paper response MUST be generated sequentially in three languages:
1. ENGLISH
2. GUJARATI (ગુજરાતી)
3. HINDI (हिंदी)

Maintain high academic rigor and formal administrative terminology across all three translations.

# MODEL SOLUTION STRUCTURE
1. Introduction (10-15% of Word Limit): Open directly with a definition, recent context/news, Constitutional Article/Supreme Court judgment, relevant statistic, or Committee recommendation.
2. Body (75-80% of Word Limit): Clear sub-headings with concise bullet points and bold lead-ins. Seamlessly include value additions (Data, Articles, Committee Reports, NITI Aayog papers, SDGs).
3. Conclusion (10-15% of Word Limit): Forward-looking, solution-oriented, and constructive (e.g., Viksit Bharat @2047, Net Zero, Constitutional ideals).

# OUTPUT FORMAT FOR THE RESPONSE
Always present the final output neatly using standard Markdown, fully repeating the output in all three languages:

# 📝 SECTION 1: ENGLISH VERSION
## DAILY MAINS ANSWER WRITING PAPER
Target Exam: [UPSC / GPSC] | Subject: [Subject Name] | Difficulty: [Easy / Moderate / Difficult]  
Total Questions: [Count] | Word Limit: [Words per question]  

### QUESTION 1
[Question text in English]  
Marks: [10 / 15 Marks] | Word Limit: [150 / 250 words]

#### MODEL ANSWER
1. Introduction  
[Introduction text]

2. Body  
[Sub-headings and bullet points]

3. Conclusion  
[Conclusion text]

---
# 📝 SECTION 2: GUJARATI VERSION (ગુજરાતી આવૃત્તિ)
[Repeat full format in Gujarati]

---
# 📝 SECTION 3: HINDI VERSION (हिंदी संस्करण)
[Repeat full format in Hindi]
"""


# Helper function to convert markdown text to formatted Word (.docx) file
def create_docx(text_content):
  doc = docx.Document()

  # Set Narrow Margins (0.5 inches on all sides)
  for section in doc.sections:
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

  # Set default font style to Verdana Size 10
  style = doc.styles["Normal"]
  font = style.font
  font.name = "Verdana"
  font.size = Pt(10)

  # Process markdown text line by line
  for line in text_content.split("\n"):
    line = line.strip()
    if not line:
      doc.add_paragraph()
      continue

    if line.startswith("# "):
      p = doc.add_paragraph()
      run = p.add_run(line.replace("# ", ""))
      run.font.size = Pt(14)
      run.bold = True
    elif line.startswith("## "):
      p = doc.add_paragraph()
      run = p.add_run(line.replace("## ", ""))
      run.font.size = Pt(12)
      run.bold = True
    elif line.startswith("### "):
      p = doc.add_paragraph()
      run = p.add_run(line.replace("### ", ""))
      run.font.size = Pt(11)
      run.bold = True
    elif line.startswith("- "):
      p = doc.add_paragraph(style="List Bullet")
      p.add_run(line.replace("- ", ""))
    else:
      p = doc.add_paragraph(line)

  buffer = io.BytesIO()
  doc.save(buffer)
  buffer.seek(0)
  return buffer


# --- MAIN INTERFACE HEADER ---
st.title("📝 UPSC / GPSC Daily Mains Paper Generator")
st.caption(
    "Powered by Groq + Llama 3.3 70B | Trilingual Output (English, Gujarati,"
    " Hindi)"
)

# Fetch API Key silently from Streamlit Secrets
groq_api_key = ""
try:
  if "GROQ_API_KEY" in st.secrets and st.secrets["GROQ_API_KEY"]:
    groq_api_key = st.secrets["GROQ_API_KEY"]
except Exception:
  pass

# Fallback only if key is missing from Secrets
if not groq_api_key:
  groq_api_key = st.sidebar.text_input(
      "Enter Groq API Key", type="password", help="Get key from console.groq.com"
  )

# --- MAIN FORM INPUTS ---
st.subheader("📋 Paper Parameters")

col_sub, col_top = st.columns(2)

with col_sub:
  subject_input = st.text_input(
      "1. Subject (Required)",
      placeholder="e.g., GS-2 Polity, GS-3 Economy, GS-1 History, Ethics",
  )

with col_top:
  topic_input = st.text_input(
      "2. Topic / Sub-topic (Optional)",
      placeholder=(
          "e.g., Judicial Activism, Inflation (Leave blank for auto-selection)"
      ),
  )

col1, col2, col3 = st.columns(3)

with col1:
  difficulty = st.selectbox(
      "3. Level of Difficulty",
      ["Moderate", "Easy", "Difficult"],
      help=(
          "Easy = Short ~150w | Moderate = Standard ~250w | Difficult = Long"
          " ~250w+"
      ),
  )

with col2:
  target_exam = st.selectbox("4. Target Exam", ["UPSC", "GPSC"])

with col3:
  num_questions = st.number_input(
      "5. Number of Questions", min_value=1, max_value=3, value=1
  )

st.divider()

# --- GENERATE ACTION ---
if st.button("🚀 Generate Mains Paper", type="primary", use_container_width=True):
  if not groq_api_key:
    st.error(
        "Groq API Key is missing. Please add GROQ_API_KEY to Streamlit Secrets."
    )
  elif not subject_input.strip():
    st.warning("Please enter a Subject (e.g., GS-2 Polity, GS-3 Economy).")
  else:
    try:
      client = Groq(api_key=groq_api_key)

      # Handle optional topic logic
      if topic_input.strip():
        topic_details = (
            f"Subject: '{subject_input.strip()}', Specific Topic:"
            f" '{topic_input.strip()}'"
        )
        file_name_tag = (
            f"{subject_input}_{topic_input}".replace(" ", "_")
            .replace("/", "_")
            .strip()
        )
      else:
        topic_details = (
            f"Subject: '{subject_input.strip()}'. (Please automatically select a"
            " high-yield, priority topic suitable for Mains from this subject)"
        )
        file_name_tag = (
            subject_input.replace(" ", "_").replace("/", "_").strip()
        )

      user_prompt = f"Generate a {target_exam} Daily Mains Answer Writing Paper. {topic_details}. Difficulty Level: {difficulty}. Total Questions: {num_questions}."

      with st.spinner(
          f"Generating {difficulty}-level paper via Groq (Llama 3.3 70B)..."
      ):
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.4,
            max_completion_tokens=4000,
        )

        generated_paper = response.choices[0].message.content

        st.success("Paper Generated Successfully!")

        # Download Button
        docx_file = create_docx(generated_paper)
        st.download_button(
            label="📥 Download as Pre-formatted Word Document (.docx)",
            data=docx_file,
            file_name=f"{target_exam}_{file_name_tag}_{difficulty}_Paper.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
            type="primary",
        )

        st.divider()
        st.markdown(generated_paper)

    except Exception as e:
      st.error(f"An error occurred: {str(e)}")
