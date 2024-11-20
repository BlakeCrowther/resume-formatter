# Resume Formatter
This tool tailors a resume to a job description with the OpenAI API while keeping the original formatting and style. Does require some manual editing of your resume and the resume_schema.json, but well worth it in the long run.

## Setup
1. Create an OpenAI account and get an API key.
2. Set the `OPENAI_API_KEY` environment variable in a `.env` file.
3. Install the dependencies in a virtual environment.
4. Run the docx parser to extract the document structure. This is what the DocXMLModifier will look for to replace the resume content.
5. Suggested Formatting:
   1. Add a table for each section of Experience or Projects.
   2. Add a row for job title, company, and dates of employment.
   3. Add a row for each bullet point in the job description.
6. Edit the resume_schema.json according to your needs. I suggest using a tool like ChatGPT to help you with this.
7. Run the resume generator to tailor the resume to a job description.
8. The output will be in the specified output directory.

### Installation
```bash
python3 -m venv venv
source venv/bin/activate
pip install -e .
```

# Usage
```bash
python src/docx_parser.py # Extracts the document structure
python src/generate_resume.py # Tailors the resume
```
