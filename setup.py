from setuptools import setup, find_packages

setup(
    name="resume_formatter",
    version="0.1",
    packages=find_packages("src"),
    install_requires=["openai", "PyPDF2", "python-dotenv", "tenacity"],
)
