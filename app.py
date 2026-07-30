import io
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from groq import Groq
import streamlit as st

# System Prompt containing all specific directives
SYSTEM_PROMPT = """
You are an exclusive Civil Services Examination Content Creator and Senior Evaluator specializing solely in UPSC (Union Public Service Commission) and GPSC (Gujarat Public Service Commission) Mains General Studies papers.
Your primary objective is to generate high-quality Mains examination questions and model solutions strictly matching the current difficulty, syllabus, analytical depth, and formatting standards of UPSC/GPSC.

# STRICT SCOPE ENFORCEMENT & OUT-OF-SCOPE DECLINATIONS
If the user's topic input is off-topic, general knowledge trivia, coding, personal advice, or non-exam related:
YOU MUST DECLINE TO ANSWER with: "This application is strictly configured to generate UPSC/GPSC Mains Answer Writing Papers. Please enter a valid syllabus topic."

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
  style = doc.styles['Normal']
  font = style.font
  font.name = 'Verdana'
  font.size = Pt(10)

  # Process markdown text line by line
  for line in text_content.split('\n'):
    line = line.strip()
    if not line:
      doc.add_paragraph()
      continue

    if line.startswith('# '):
      p = doc.add_paragraph()
      run = p.add_run(line.replace('# ', ''))
      run.font.size = Pt(14)
      run.bold = True
    elif line.startswith('## '):
      p = doc.add_paragraph()
      run = p.add_run(line.replace('## ', ''))
      run.font.size = Pt(12)
      run.bold = True
    elif line.startswith('### '):
      p = doc.add_paragraph()
      run = p.add_run(line.replace('### ', ''))
      run.font.size = Pt(11)
      run.bold = True
    elif line.startswith('- '):
      p = doc.add_paragraph(style='List Bullet')
      p.add_run(line.replace('- ', ''))
    else:
      p = doc.add_paragraph(line)

  buffer = io.BytesIO()
  doc.save(buffer)
  buffer.seek(0)
  return buffer


# --- STREAMLIT UI SETUP ---
st.set_page_config(
    page_title="UPSC/GPSC Mains Paper Generator",
    page_icon="📝",
    layout="wide",
)

st.title("📝 UPSC / GPSC Daily Mains Paper Generator")
st.caption(
    "Powered by Groq + Llama 3.3 70B | Trilingual Output (English, Gujarati,"
    " Hindi)"
)

# Fetch API Key automatically from Streamlit Secrets if configured, else sidebar
api_key_secret = st.secrets.get("GROQ_API_KEY", "")

st.sidebar.header("⚙️ Settings")
if api_key_secret:
  groq_api_key = api_key_secret
  st.sidebar.success("API Key auto-loaded from App Secrets!")
else:
  groq_api_key = st.sidebar.text_input("Enter Groq API Key", type="password")

target_exam = st.sidebar.selectbox("Target Exam", ["UPSC", "GPSC"])
difficulty = st.sidebar.selectbox(
    "Difficulty Level", ["Moderate", "Easy", "Difficult"]
)
num_questions = st.sidebar.number_input(
    "Number of Questions", min_value=1, max_value=3, value=1
)

# Main Form
topic_input = st.text_input(
    "Subject / Topic Name",
    placeholder="e.g., GS-2 Judiciary, GS-3 Renewable Energy, GS-1 Modern History",
)

if st.button("🚀 Generate Mains Paper", type="primary"):
  if not groq_api_key:
    st.error(
        "Groq API Key is missing. Please add it to Streamlit Secrets or sidebar."
    )
  elif not topic_input:
    st.warning("Please enter a subject or topic.")
  else:
    try:
      client = Groq(api_key=groq_api_key)

      user_prompt = f"Generate a {target_exam} Daily Mains Answer Writing Paper on the topic: '{topic_input}'. Difficulty Level: {difficulty}. Total Questions: {num_questions}."

      with st.spinner(
          "Analyzing PYQs and generating trilingual answers via Groq (Llama"
          " 3.3 70B)..."
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
        st.markdown(generated_paper)

        # Download Button
        docx_file = create_docx(generated_paper)
        st.download_button(
            label="📥 Download as Pre-formatted Word Document (.docx)",
            data=docx_file,
            file_name=(
                f"{target_exam}_{topic_input.replace(' ', '_')}_Mains_Paper.docx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
        )

    except Exception as e:
      st.error(f"An error occurred: {str(e)}")
