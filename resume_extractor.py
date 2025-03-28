import pdfplumber
import pandas as pd
import re
import os
from datetime import datetime

# Define output file path (update as needed)
output_path = "C:/Users/ADMIN/resume_project/resume_with_links.xlsx"
resume_folder = "Resumes/"  # Folder where resumes are stored

# Function to extract text from a PDF
def extract_text_from_pdf(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            extracted_text = page.extract_text()
            if extracted_text:
                text += extracted_text + "\n"
    return text.strip()

# Function to clean text (removes unwanted characters)
def clean_text(text):
    return re.sub(r'[\x00-\x1F\x7F]', '', text)  # Removes control characters

# Function to extract the introduction section
def extract_intro(text):
    intro_match = re.search(r"(?i)(Summary|Profile|Objective|About Me)(.*?)(?=Experience|Professional Experience|Employment History|$)", text, re.DOTALL)
    return clean_text(intro_match.group(2).strip()) if intro_match else "Not Found"

# Function to extract work experience as a single block
def extract_work_experience(text):
    work_exp_match = re.search(r"(?i)Experience(.*?)(?=Education|Projects|Certifications|Skills|$)", text, re.DOTALL)
    return clean_text(work_exp_match.group(1).strip()) if work_exp_match else "Not Found"

# Function to extract education section and clean formatting
def extract_education(text):
    education_match = re.search(r"(?i)Education(.*?)(?=Publications|Skills|Projects|Experience|$)", text, re.DOTALL)
    if education_match:
        cleaned_text = re.sub(r"[\-\.\|]", " ", education_match.group(1).strip())  # Replace hyphens, dots, and pipes with spaces
        cleaned_text = re.sub(r"\s+", " ", cleaned_text).strip()  # Normalize spacing
        return cleaned_text
    return "Not Found"

# Function to check if Education contains IIT/IIIT/BITS/NIT
def check_top_institute(education_text):
    if pd.isna(education_text):
        return "No"
    
    iit_keywords = ["IIT", "Indian Institute of Technology", 
                    "IIIT", "International Institute of Information Technology", 
                    "BITS", "Birla Institute of Technology and Science",
                    "NIT", "National Institute of Technology"]
    
    for keyword in iit_keywords:
        if keyword.lower() in str(education_text).lower():
            return "Yes"
    
    return "No"

# Function to calculate total experience in years
def calculate_total_experience(work_exp_text):
    if pd.isna(work_exp_text) or work_exp_text == "Not Found":
        return 0  # No valid work experience found

    # Find all year ranges (e.g., "2015-2018", "2020–Present")
    year_matches = re.findall(r"(\d{4})\s*[-–]\s*(\d{4}|Present)", work_exp_text)

    total_experience = 0
    current_year = datetime.now().year  # Get the current year

    for start, end in year_matches:
        start_year = int(start)
        end_year = current_year if end == "Present" else int(end)
        
        # Add the experience duration to total
        total_experience += (end_year - start_year)

    return total_experience

# Function to generate clickable resume link
def generate_resume_link(file_name):
    if pd.isna(file_name) or file_name.strip() == "":
        return "File Not Found"
    
    # Create the relative file path
    resume_path = os.path.join(resume_folder, file_name)
    
    # Ensure path formatting for Excel hyperlinks
    return f'=HYPERLINK("{resume_path}", "Open Resume")'

# Folder containing resumes
resume_folder_local = "C:/Users/ADMIN/resume_project/resumes"
resume_data = []

# Process all PDF files in the folder
for filename in os.listdir(resume_folder_local):
    if filename.endswith(".pdf"):
        pdf_path = os.path.join(resume_folder_local, filename)
        
        text = extract_text_from_pdf(pdf_path)  # Extract text from PDF
        work_experience = extract_work_experience(text)  # Extract Work Experience
        education = extract_education(text)  # Extract Education
        
        details = {
            "File Name": filename,
            "Name": text.split("\n")[0].strip() if text else "Not Found",
            "Phone": re.search(r"\b\d{10}\b", text).group(0) if re.search(r"\b\d{10}\b", text) else "Not Found",
            "Email": re.search(r"[\w\.-]+@[\w\.-]+\.\w+", text).group(0) if re.search(r"[\w\.-]+@[\w\.-]+\.\w+", text) else "Not Found",
            "LinkedIn": re.search(r"(https?://[^\s]+linkedin\.com[^\s]+)", text, re.IGNORECASE).group(1) if re.search(r"(https?://[^\s]+linkedin\.com[^\s]+)", text, re.IGNORECASE) else "Not Found",
            "Intro": extract_intro(text),
            "Work Experience": work_experience,
            "Total Experience (Years)": calculate_total_experience(work_experience),  # New column
            "Education": education,
            "Top Institute Graduate": check_top_institute(education),  # IIT/IIIT/BITS/NIT check
            "Resume Link": generate_resume_link(filename)  # Clickable link
        }
        resume_data.append(details)

# Convert to DataFrame
df = pd.DataFrame(resume_data)

# Save to Excel with correct hyperlink format
with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
    df.to_excel(writer, index=False)

print(f"✅ Resume data extracted and saved to: {output_path}")
