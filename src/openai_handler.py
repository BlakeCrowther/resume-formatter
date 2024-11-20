# openai_handler.py
import openai
import json
import time
from typing import List, Dict, Any
from tenacity import retry, stop_after_attempt, wait_exponential


class OpenAIHandler:
    def __init__(self, api_key: str, model: str = "gpt-4o-mini"):
        """
        Initialize OpenAI handler with API key and model selection.

        Args:
            api_key (str): OpenAI API key
            model (str): Model to use for completions (default: gpt-4o-mini)
        """
        self.client = openai.AsyncClient(api_key=api_key)
        self.model = model

    @retry(
        stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=4, max=10)
    )
    async def get_keywords(self, job_description: str) -> List[str]:
        """
        Extract relevant keywords from job description using OpenAI.

        Args:
            job_description (str): Job description text

        Returns:
            List[str]: List of extracted keywords
        """
        try:
            prompt = f"""You are an expert recruiter analyzing job descriptions.
            Given this job description, identify the most important technical skills, 
            soft skills, tools, and qualifications that would be valuable to see in 
            an applicant's resume.

            Job Description:
            {job_description}

            Return ONLY a comma-separated list of the most relevant keywords and phrases.
            Focus on specific, concrete terms rather than general concepts.
            Limit your response to 20-25 of the most important terms."""

            response = await self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {
                        "role": "system",
                        "content": "You are a helpful recruiter analyzing job descriptions.",
                    },
                    {"role": "user", "content": prompt},
                ],
                temperature=0.3,
            )

            # Parse comma-separated response into list
            keywords = [
                keyword.strip()
                for keyword in response.choices[0].message.content.split(",")
            ]

            print(f"Successfully extracted {len(keywords)} keywords")
            return keywords

        except Exception as e:
            print(f"Error in keyword extraction: {str(e)}")
            raise

    @retry(
        stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=4, max=10)
    )
    async def tailor_resume(
        self, resume_schema: Dict[str, Any], keywords: List[str], job_description: str
    ) -> Dict[str, Any]:
        """
        Tailor resume content based on job description and keywords.

        Args:
            resume_schema (dict): Original resume schema
            keywords (list): List of important keywords
            job_description (str): Original job description

        Returns:
            dict: Updated resume schema with tailored content
        """
        try:
            # Create a focused prompt for the AI
            prompt = f"""You are an expert resume writer tailoring a resume to a specific job.
            
            Job Description:
            {job_description}

            Important Keywords to Include (where relevant):
            {', '.join(keywords)}

            Current Resume Content (in JSON format):
            {json.dumps(resume_schema, indent=2)}

            Task:
            1. Rewrite each bullet point to naturally incorporate relevant keywords
            2. Maintain the original meaning and achievements
            3. Respect the min_chars and max_chars constraints for each bullet point
            4. For each bullet point, add a "keywords" list showing which keywords were used from the important keywords list
            5. Keep the tone professional and achievement-focused
            6. Use strong action verbs
            7. Quantify achievements where possible
            8. Do not invent new achievements or skills

            Return ONLY valid JSON with no additional text or formatting.
            The JSON should contain the complete resume schema with updated bullet points and keyword tracking."""

            response = await self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {
                        "role": "system",
                        "content": "You are a professional resume writer. Return only valid JSON.",
                    },
                    {"role": "user", "content": prompt},
                ],
                temperature=0.4,
            )

            # Extract response content
            response_text = response.choices[0].message.content.strip()

            # Remove any markdown code block indicators if present
            response_text = (
                response_text.replace("```json", "").replace("```", "").strip()
            )

            # Parse the response and validate it's proper JSON
            try:
                tailored_schema = json.loads(response_text)
            except json.JSONDecodeError as e:
                print(f"OpenAI response was not valid JSON: {str(e)}")
                print(f"Response content: {response_text}")
                raise ValueError("Failed to generate valid JSON response")

            # Validate the schema structure
            self._validate_schema(tailored_schema)

            print("Successfully tailored resume content")
            return tailored_schema

        except Exception as e:
            print(f"Error in resume tailoring: {str(e)}")
            raise

    def _validate_schema(self, schema: Dict[str, Any]) -> None:
        """Validate the basic structure of the resume schema."""
        required_sections = ["experiences", "projects"]
        for section in required_sections:
            if section not in schema:
                raise ValueError(f"Missing required section: {section}")

        # Validate each section has the required fields
        for experience in schema["experiences"]:
            if "bullet_points" not in experience:
                raise ValueError("Experience missing bullet_points")
            for bullet in experience["bullet_points"]:
                if not all(key in bullet for key in ["text", "keywords"]):
                    raise ValueError("Bullet point missing required fields")

        for project in schema["projects"]:
            if "bullet_points" not in project:
                raise ValueError("Project missing bullet_points")
            for bullet in project["bullet_points"]:
                if not all(key in bullet for key in ["text", "keywords"]):
                    raise ValueError("Bullet point missing required fields")


# Example usage
async def main():

    # Initialize the handler
    handler = OpenAIHandler(api_key="your-api-key")

    # Example job description
    job_description = """
    We are seeking a Senior Software Engineer with 5+ years of experience in Python
    and React. Must have experience with AWS, Docker, and CI/CD pipelines.
    Experience with machine learning and data analysis is a plus.
    """

    # Get keywords
    keywords = await handler.get_keywords(job_description)
    print("Extracted Keywords:", keywords)

    # Example resume schema
    resume_schema = {
        "experiences": [
            {
                "title": "Software Engineer",
                "company": "Tech Corp",
                "bullet_points": [
                    {
                        "text": "Developed and maintained Python microservices",
                        "min_chars": 50,
                        "max_chars": 200,
                        "keywords": [],
                    }
                ],
            }
        ],
        "projects": [
            {
                "title": "ML Pipeline",
                "bullet_points": [
                    {
                        "text": "Built data processing pipeline using Python and AWS",
                        "min_chars": 50,
                        "max_chars": 150,
                        "keywords": [],
                    }
                ],
            }
        ],
    }

    # Tailor resume
    tailored_schema = await handler.tailor_resume(
        resume_schema, keywords, job_description
    )
    print("Tailored Schema:", json.dumps(tailored_schema, indent=2))


if __name__ == "__main__":
    import asyncio

    asyncio.run(main())
