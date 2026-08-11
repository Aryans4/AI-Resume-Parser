import os
from time import sleep
from pathlib import Path
from dotenv import load_dotenv
from groq import Groq
from pydantic import BaseModel,Field

load_dotenv()
my_api_key=os.getenv("GROQ_API_KEY")
if not my_api_key:
    raise ValueError("api not found")

client=Groq(api_key=my_api_key)
model="llama-3.1-8b-instant"
job_description="""
Description
Do you want to solve real customer problems through innovative technology? Do you enjoy working on scalable services in a collaborative team environment? Do you want to see your code directly impact millions of customers worldwide?

At Amazon, we hire the best minds in technology to innovate and build on behalf of our customers. Customer obsession is part of our company DNA, which has made us one of the world's most beloved brands.

Our Software Development Engineers (SDEs) use modern technology to solve complex problems while seeing their work's impact first-hand. The challenges SDEs solve at Amazon are meaningful and influence millions of customers, sellers, and products globally. We seek individuals passionate about creating new products, features, and services while managing ambiguity in an environment where development cycles are measured in weeks, not years.

At Amazon, we believe in ownership at every level. As an SDE-I, you'll own the entire lifecycle of your code - from design through deployment and ongoing operations. This ownership mindset, combined with our commitment to operational excellence, ensures we deliver the highest quality solutions for our customers.

We're looking for curious minds who think big and want to define tomorrow's technology. At Amazon, you'll grow into the high-impact engineer you know you can be, supported by a culture of learning and mentorship. Every day brings exciting new challenges and opportunities for personal growth.
Key job responsibilities
• Collaborate and communicate effectively with experienced cross-disciplinary Amazonians to design, build, and operate innovative products and services that delight our customers, while participating in technical discussions to drive solutions forward.
• Design and develop scalable solutions using cloud-native architectures and microservices in a large distributed computing environment.
• Participate in code reviews and contribute to technical documentation.
• Build and maintain resilient distributed systems that are scalable, fault-tolerant, and cost-effective.
• Leverage and contribute to the development of GenAI and AI-powered tools to enhance development productivity while staying current with emerging technologies.
• Write clean, maintainable code following best practices and design patterns.
• Work in an agile environment practicing CI/CD principles while participating in operational responsibilities including on-call duties.
• Demonstrate operational excellence through monitoring, troubleshooting, and resolving production issues.
Basic Qualifications
- Experience with at least one general-purpose programming language such as Java, Python, C++, C#, Go, Rust, or TypeScript
- Experience with data structure implementation, basic algorithm development, and/or object-oriented design principles
- Currently has, or is in the process of obtaining a bachelor’s degree in Computer Science, Computer Engineering, Data Science, Information Systems, or related STEM fields
- Must be 18 years of age of older
Preferred Qualifications
- Experience from previous technical internship(s) or demonstrated project experience
- Experience with one or more of the following: AI tools for development productivity, Cloud platforms (preferably AWS), Database systems (SQL and NoSQL), Contributing to open-source projects, Version control systems, Debugging and troubleshooting complex systems
- Demonstrated ability to learn and adapt to new technologies quickly
- Basic understanding of software development lifecycle (SDLC)
- Strong problem-solving and analytical skills
- Excellent written and verbal communication skills
"""
#JD ki class banayenge jo batayegi JD me kya chahiye.
class JobD(BaseModel):
 role: str | None = None
 required_skills: list[str] | None = Field(default_factory=list)
 preferred_skills: list[str] | None = Field(default_factory=list)
 minimum_experience: float | None = None
 education_requirement: list[str] | None = Field(default_factory=list)
 responsibilities: list[str] | None = Field(default_factory=list)

JobD_schema=JobD.model_json_schema()
#system ko kya krna hai.
system_prompt=f"""
You are an expert HR assistant.

Your job is to analyze job descriptions and extract
structured information from them.

Return ONLY valid JSON matching this schema:

{JobD_schema}
IMPORTANT:
Do NOT return the schema itself.
Do NOT return fields like "properties", "title" or "type".
Fill the schema with actual information extracted from the job description.

If minimum experience is not mentioned, return null.
IMPORTANT:
minimum_experience must be either a number or JSON null.
Never return "null" as a string.
If information for a list is missing, return an empty list.
Do no t invent information.
Return exactly these fields:
role: string
required_skills: list of strings
preferred_skills: list of strings
minimum_experience: number or null
education_requirement: list of strings
responsibilities: list of strings

The role field MUST be a plain string.
Example:
"role": "Software Development Engineer (SDE-I)"

Do NOT make role an object.
Do NOT use {{"title": "..."}}for role.
"""
#user kya krega.
user_prompt=f"""
Analye the follwing job description.
{job_description}
"""
#Job Description ka response JSON format me.
message_system={
   "role": "system",
   "content": system_prompt
}
message_user={
   "role": "user",
   "content": user_prompt
}
response_format={
   "type": "json_object"
}
messages=[message_system,message_user]
response=client.chat.completions.create(model=model,messages=messages,response_format=response_format)
answer=response.choices[0].message.content
raw_json=answer

#print raw json
import json
job_data=json.loads(raw_json)
job=JobD(**job_data)# here we got job descripton in json format

#same for resume extraction
class matchResult(BaseModel):
   score: float | None = 0.0
   candidate_name: str | None = "Unknown"
   matching_skills: list[str] | None = Field(default_factory=list)
   missing_skills: list[str] | None = Field(default_factory=list)
   experience_requirement_met: bool | None = False
   verdict: str | None = "No verdict"

match_schema=matchResult.model_json_schema()   

from typing import Any

class Experience(BaseModel):
   company: str | None=None
   role: str | None=None
   duration: str | None=None
   description:str | None=None
   skills_used:list[Any] | None = Field(default_factory=list)

class Resume(BaseModel):
   name: str | None=None
   email:str | None=None 
   phone:str | None=None
   total_experience_years:int | None=None
   experience:list[Experience] | None = Field(default_factory=list)
   education: list[Any] | None = Field(default_factory=list)
   skills: list[Any] | None = Field(default_factory=list)
   projects:list[Any] | None = Field(default_factory=list)
   certification:list[Any] | None = Field(default_factory=list)

resume_schema=Resume.model_json_schema()

#Uploading RESUME
from pypdf import PdfReader
from docx import Document

#read_file inme se kisi ko call krega.inko phle define krna padta hai.
def read_pdf(file_path):
   reader=PdfReader(file_path)
   text=""
   for page in reader.pages:
      page_txt=page.extract_text()
      if page_txt:
         text+=page_txt + "\n"
   return text

def read_docx(file_path):
   document= Document(file_path)
   text=""
   for paragraph in document.paragraphs:
     para_text=paragraph.text.strip()
     if para_text:
        text+=para_text + "\n"

   for table in document.tables:
      for row in table.rows:
         for cell in row.cells:
             if cell.text.strip():
               text+=cell.text + "\n"
   return text           

#uploaded resume phle is function me aayega.
def read_file(file_path):
   if file_path.suffix.lower()==".pdf":
      return read_pdf(file_path)
   elif file_path.suffix.lower()==".docx":
      return read_docx(file_path)
   else:
      return None

def resume_parser(resume_text):
   system_prompt=f"""
 You are an expert resume parser.

    Extract information from the resume based on its meaning,
    not only based on exact section headings.

    Different resumes may use different headings.

    For example:
    - Experience
    - Professional Experience
    - Work History
    - Employment
    - Internships

    These may all contain relevant experience.

    Skills may also appear in the skills section, work experience,
    internships or projects.

    Return ONLY valid JSON matching this schema:

    {resume_schema}

    Important rules:

    1. Do not invent information.
    2. If a value is not available, return null.
    3. If a list has no information, return an empty list.
    4. Include internships inside experiences.
    5. Extract skills mentioned across the entire resume.
   """
   user_prompt=f"""
    parse the following resume.
    {resume_text}
   """
   message_system={
      "role":"system",
      "content":system_prompt 
   }
   message_user={
      "role": "user",
      "content": user_prompt
   }
   response_format={
      "type": "json_object"
   }
   messages=[message_system,message_user]
   response=client.chat.completions.create(model=model,messages=messages,response_format=response_format)
   raw_output=response.choices[0].message.content
   import json
   data=json.loads(raw_output)
   resume=Resume(**data)
   return resume

def final_match(job,resume):
    prompt=f"""
    You are an HR recruiter.

    Compare the candidate's resume with the job description.

    JOB DESCRIPTION:
    {job.model_dump_json(indent=2)}

    CANDIDATE RESUME:
    {resume.model_dump_json(indent=2)}
    Return JSON matching this schema:

    {match_schema}

    Give me:

    1. Candidate name
    2. Matching skills
    3. Missing important skills
    4. Whether experience requirement is met
    5. Overall match percentage from 0 to 100
    6. A short final verdict

    Keep the response concise and easy to read.
    """
    response_format={
        "type": "json_object"
    }
    message={
      "role": "user",
      "content":prompt
    }
    messages=[message]
    response=client.chat.completions.create(model=model,messages=messages,response_format=response_format, temperature=0.0)
    data=response.choices[0].message.content
    import json
    data=json.loads(data)
    if isinstance(data, list):
        data = data[0]
    return matchResult(**data)

#Program Execution Started.
resume_folder=Path("resumes")
all_results=[]
for file_path in resume_folder.iterdir():
   if file_path.suffix.lower() not in [".pdf",".docx"]:
      continue
   print("Processing:", file_path.name)
   resume_text=read_file(file_path) #function call
   parsed_resume=resume_parser(resume_text) #LLM Call 1
   sleep(5)
   result=final_match(job,parsed_resume) #LLM Call 2
   sleep(5)
   print("score:", result.score)
   all_results.append({
      "name": parsed_resume.name,
      "score": result.score,
      "verdict":result.verdict
   })




