import json
import os
from datetime import datetime
from dotenv import load_dotenv
from docx import Document
from docx.shared import Pt
import asyncio

from docx_parser import DocxXMLParser
from docx_modifier import DocXMLModifier
from openai_handler import OpenAIHandler


class ResumeTailoringSystem:
    def __init__(self, resume_name, output_dir_name):
        # Get the root directory (one level up from src)
        root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        # Initialize input and output directories
        self.inputs = os.path.join(root_dir, "inputs")
        self.outputs = os.path.join(root_dir, "outputs")
        self.resume_path = os.path.join(self.inputs, f"{resume_name}.docx")
        self.output_dir = os.path.join(self.outputs, output_dir_name)
        self.config_path = os.path.join(self.inputs, "input_history.json")
        self._setup_directories()

    def _setup_directories(self):
        """Create necessary directories if they don't exist."""
        os.makedirs(self.output_dir, exist_ok=True)

    def extract_structure(self):
        """Extract document structure and save to inputs directory."""
        parser = DocxXMLParser(self.resume_path)
        structure = parser.extract_document_structure(self.inputs)

    async def generate_tailored_resume(self, job_description):
        """Main function to generate tailored resume from job description."""
        os.makedirs(self.output_dir, exist_ok=True)

        openai_handler = OpenAIHandler(os.getenv("OPENAI_API_KEY"))

        # Get keywords from OpenAI
        print("Getting keywords from OpenAI")
        keywords = await openai_handler.get_keywords(job_description)

        # Load and validate schema
        print("Loading and validating original resume schema")
        with open(os.path.join(self.inputs, "resume_schema.json"), "r") as f:
            resume_schema = json.load(f)

        # Get tailored resume content
        print("Tailoring resume content")
        tailored_schema = await openai_handler.tailor_resume(
            resume_schema, keywords, job_description
        )

        # Save the tailored schema to a file
        print("Saving tailored resume schema to outputs directory")
        tailored_schema["keywords"] = keywords
        with open(os.path.join(self.output_dir, "tailored_schema.json"), "w") as f:
            json.dump(tailored_schema, f, indent=2)

        # Initialize DocXMLModifier and modify resume
        print("Modifying original resume with tailored schema")
        modifier = DocXMLModifier(self.resume_path)
        output_path = modifier.modify_docx(tailored_schema, self.output_dir)

        return output_path

    def _save_last_inputs(self, resume_name, output_dir_name):
        """Save the last used inputs to config file."""
        config = {"last_resume_name": resume_name, "last_output_dir": output_dir_name}
        with open(self.config_path, "w") as f:
            json.dump(config, f)

    def _get_last_inputs(self):
        """Get the last used inputs from config file."""
        try:
            with open(self.config_path, "r") as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            return {"last_resume_name": "", "last_output_dir": ""}


async def main():
    # load env
    load_dotenv()

    # Create a temporary instance to access config methods
    temp_system = ResumeTailoringSystem("temp", "temp")
    last_inputs = temp_system._get_last_inputs()

    resume_name = "Blake-Crowther-Resume"

    # Get output directory name with default value
    default_output = last_inputs["last_output_dir"]
    output_prompt = (
        f"Enter output directory name [press Enter for '{default_output}']: "
        if default_output
        else "Enter output directory name: "
    )
    output_dir_name = input(output_prompt)
    output_dir_name = output_dir_name or default_output

    # Initialize resume tailoring system
    resume_tailor = ResumeTailoringSystem(resume_name, output_dir_name)

    # Save the inputs
    resume_tailor._save_last_inputs(resume_name, output_dir_name)

    # Validate resume exists
    if not os.path.exists(resume_tailor.resume_path):
        print(f"Error: Resume not found at {resume_tailor.resume_path}")
        return

    # Get job description
    print("Paste job description and press Enter twice when done:")
    job_description_lines = []

    empty_line_count = 0
    while empty_line_count < 2:
        line = input().strip()  # Strip whitespace immediately on input
        if not line:
            empty_line_count += 1
        else:
            empty_line_count = 0
            job_description_lines.append(line)

    job_description = " ".join(job_description_lines)

    if not job_description:
        print("Error: Job description cannot be empty")
        return
    else:
        # Generate tailored resume
        await resume_tailor.generate_tailored_resume(job_description)


if __name__ == "__main__":
    asyncio.run(main())
