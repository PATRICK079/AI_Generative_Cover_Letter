import streamlit as st
import google.generativeai as genai
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import simpleSplit
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from io import BytesIO

# Helper Functions
def configure(gemini_api_key):
    genai.configure(api_key=gemini_api_key)

def prompt(job_title, company_name, job_description, platform, recipient_name="Hiring Manager"):
    return f"""This is a canditate's CV in pdf format. Extract the text considering layouts, headings and subheadings.
Then write a professional cover letter of 250-330 words based on this CV and the following details:
- Job Title: {job_title}
- Company: {company_name}
- Recipient’s Name: {recipient_name}
- Job Description: {job_description}
- Platform of the Advertisement: {platform}

Please ensure the cover letter is tailored to the specific role and company. It should:
1. Clearly state the purpose of the letter and introduce the candidate.
2. Demonstrate genuine interest in the job and company.
3. Complement the CV by offering a little more detail about key experiences, not to simply repeat the CV in paragraph form.
4. Focus on transferable skills if no direct experience exists.
5. Maintain a professional tone, be concise, and thank the reader for their time and consideration.

Structure:
- Introduction: State why the candidate is interested in the role and company.
- Body: Provide 2-3 key examples from the CV that show how the candidate's skills and experience align with the role.
- Closing: Reaffirm the candidate's interest and gratitude for the opportunity.

Be sure to personalize the letter and avoid generic language. The body paragraphs must be equal in size.
"""

def generator(cv, prompt):
    cv = genai.upload_file(cv, mime_type="application/pdf")
    version = 'models/gemini-1.5-flash'
    model = genai.GenerativeModel(version, generation_config={"temperature": 1.5, 
                                                            "top_p": 0.96, 
                                                            "max_output_tokens": 600})
    response = model.generate_content([cv, prompt])
    cv.delete()
    return response.text

def cover_letter_pdf(text):
    # Set up the canvas and page properties
    pdf_buffer = BytesIO()
    c = canvas.Canvas(pdf_buffer, pagesize=A4)
    width, height = A4
    c.setFont("Times-Roman", 12)

    # Margins
    margin = 72  # 1 inch = 72 points
    text_width = width - 2 * margin

    # Start position
    x, y = margin, height - margin

    # Line spacing
    line_height = 14  # Single-spaced lines, font size 12 + 2pt padding

    # Process text
    paragraphs = text.split("\n\n")
    for i, paragraph in enumerate(paragraphs):
        if i > 0:
            y -= line_height
        lines = c.beginText(x, y)
        wrapped_lines = simpleSplit(paragraph, "Times-Roman", 12, text_width)
        for line in wrapped_lines:
            lines.textLine(line)
            y -= line_height
        c.drawText(lines)

    c.save()
    pdf_buffer.seek(0)
    return pdf_buffer

def cover_letter_docx(text):
    doc = Document()
    section = doc.sections[0]
    
    # Set A4 dimensions: 21.0 x 29.7 cm (8.27 x 11.69 inches)
    section.page_height = Inches(11.69)
    section.page_width = Inches(8.27)
    
    # Set one-inch margins
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    
    # Set default paragraph formatting
    style = doc.styles['Normal']
    style.paragraph_format.space_after = Pt(14)
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.line_spacing = Pt(14)
    style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)

    # Split text into paragraphs and add to the document
    paragraphs = text.split("\n\n")
    for i, para in enumerate(paragraphs):
        if para.strip():
            paragraph = doc.add_paragraph(para.strip())
            
            # Set space after for all but the last paragraph
            if i == len(paragraphs) - 1:
                paragraph.paragraph_format.space_after = Pt(0)

    # Save document to buffer
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# Streamlit Interface
col1, col2 = st.columns([1, 6])
with col1:
    st.image("gemini_generated_image.jpg", width=1000)
with col2:
    st.markdown("""
        <div style="display: flex; align-items: center; height: 100%;">
            <h1 style="font-size: 32px; margin: 0;">Super AI-Driven Cover Letter Builder</h1>
        </div>
    """, unsafe_allow_html=True)

st.write("""Welcome to the Super AI-Driven Cover Letter Builder! 
         Powered by Google Gemini 1.5 Flash, this tool helps you craft personalized, professional cover letters in just a few clicks. 
         Simply submit your details below, and let the app generate your cover letter. After creating your cover letter, you can download it in PDF or DOCX format.
         You can also refine it further using [this guide](https://capd.mit.edu/resources/how-to-write-an-effective-cover-letter/) for an even more tailored result.""")

st.write("""
To use this app, you need a Gemini API key. Please follow the steps below if you don't have one:

1. [Get a Gemini API key in Google AI Studio](https://aistudio.google.com/app/apikey?_gl=1*1elzb2j*_ga*Nzg0MDg5MzQ5LjE3MzIwNDcxOTk.*_ga_P1DBVKWT6V*MTczMjM1ODk3Ny4xOS4xLjE3MzIzNTk1MDguMjQuMC4xMzEwOTE1NTgy).
2. Follow the instructions to generate your API key.
3. Once you have your API key, come back to this app and enter it below.
""")

st.markdown(
    "<h1 style='color:red; text-transform: uppercase; font-size: 20px; '>Do not submit sensitive, confidential, or personal information!</h1>",
    unsafe_allow_html=True
)
st.write("Submit the details below to receive a custom, professional cover letter.")

job_title = st.text_input("Enter the job title *")
company_name = st.text_input("Enter the company name *")
recipient_name = st.text_input("Enter the recipient's name")
job_description = st.text_area("Enter the job description *")
platform = st.text_input("Platform where you saw the advertisement")
cv = uploaded_file = st.file_uploader("""Upload your CV in PDF format *     
                                      (REMOVE OR MASK ANY PERSONAL INFORMATION FROM YOUR CV BEFORE UPLOADING!)""", type="pdf")

if cv is not None:
    st.success("File uploaded successfully!")

gemini_api_key = st.text_input("Enter your Gemini API Key:", type="password")

if "generated_text" not in st.session_state:
    st.session_state.generated_text = ""
if "file_format" not in st.session_state:
    st.session_state.file_format = "PDF"

if st.button("Generate Cover Letter"):
    if not job_title or not company_name or not job_description or not cv or not gemini_api_key:
        st.warning("Please fill in all required fields before generating a cover letter.")
    else:
        with st.spinner("Generating your cover letter..."):
            try:
                configure(gemini_api_key)
                prompt_text = prompt(job_title, company_name, job_description, platform, recipient_name)
                response = generator(cv, prompt_text)

                st.session_state.generated_text = response

            except Exception as e:
                st.error(f"Error generating cover letter: {e}")

if not st.session_state.generated_text == "":
    try:
        st.subheader("Generated Cover Letter")
        st.text(st.session_state.generated_text)

        col1, col2, col3 = st.columns([3, 3, 2])
        with col2:  
            # Radio button to choose file format
            st.subheader("Download")
            file_format = st.radio("Choose a file format to download:", ["PDF", "DOCX"], horizontal=True)
            st.session_state.file_format = file_format

            # Create the file based on the selected format
            if st.session_state.file_format == "DOCX":
                file_data = cover_letter_docx(st.session_state.generated_text)
                file_name = "Firstname_Lastname_CoverLetter.docx"
                mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            elif st.session_state.file_format == "PDF":
                file_data = cover_letter_pdf(st.session_state.generated_text)
                file_name = "Firstname_Lastname_CoverLetter.pdf"
                mime_type = "application/pdf"

            # Show download button
            st.download_button(
                label="Download",
                data=file_data,
                file_name=file_name,
                mime=mime_type,
                use_container_width=True
            )

    except Exception as e:
        st.error(f"Error generating the file: {e}")
