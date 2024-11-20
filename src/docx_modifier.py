import os
import json
import zipfile
import xml.etree.ElementTree as ET
import re
from datetime import datetime


class DocXMLModifier:
    def __init__(self, original_docx_path):
        """
        Initialize the Docx XML Modifier

        Args:
            original_docx_path (str): Path to the original DOCX file
        """
        self.original_docx_path = original_docx_path
        self.resume_name = os.path.basename(original_docx_path)
        self.namespaces = {
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        }
        # Register namespaces to make XML parsing easier
        for prefix, uri in self.namespaces.items():
            ET.register_namespace(prefix, uri)

    def modify_docx(self, tailored_schema, output_dir):
        """
        Modify the DOCX based on the provided schema

        Args:
            tailored_schema (dict): Tailored resume content
            output_dir (str): Directory to save the modified DOCX

        Returns:
            str: Path to the modified DOCX
        """
        # Create a temporary working directory within output_dir
        temp_dir = os.path.join(output_dir, "docx_temp")
        os.makedirs(temp_dir, exist_ok=True)

        # Unzip the original DOCX
        with zipfile.ZipFile(self.original_docx_path, "r") as zip_ref:
            zip_ref.extractall(temp_dir)

        # Path to the document.xml
        document_xml_path = os.path.join(temp_dir, "word", "document.xml")

        # Parse the XML
        tree = ET.parse(document_xml_path)
        root = tree.getroot()

        # Modify experiences and projects
        self._modify_sections(
            root,
            tailored_schema.get("experiences", [])
            + tailored_schema.get("projects", []),
        )

        # Write modified XML back
        tree.write(document_xml_path, encoding="UTF-8", xml_declaration=True)

        # Rezip the modified files into a new DOCX
        output_path = os.path.join(output_dir, self.resume_name)
        self._create_docx(temp_dir, output_path)

        # Clean up temporary directory
        import shutil

        shutil.rmtree(temp_dir)

        return output_path

    def _modify_sections(self, root, all_sections):
        """
        Modify sections by matching existing content

        Args:
            root (ET.Element): Root XML element
            all_sections (list): Combined list of experiences and projects
        """
        # Find all tables in the document
        tables = root.findall(".//w:tbl", self.namespaces)

        for table in tables:
            # Extract full text content of the table
            table_text = self._extract_table_text(table)

            # Find matching section
            matched_section = self._find_matching_section(table_text, all_sections)

            if matched_section:
                # Update table content
                self._update_table_content(table, matched_section)

    def _extract_table_text(self, table):
        """
        Extract full text content of a table

        Args:
            table (ET.Element): Table XML element

        Returns:
            str: Extracted text content
        """
        # Find all text elements in the table
        text_elements = table.findall(".//w:t", self.namespaces)

        # Combine text, removing extra whitespace
        full_text = " ".join(
            [t.text.strip() for t in text_elements if t.text and t.text.strip()]
        )

        return full_text

    def _find_matching_section(self, table_text, sections):
        """
        Find a matching section based on text content

        Args:
            table_text (str): Text content of the table
            sections (list): List of sections to match against

        Returns:
            dict or None: Matched section
        """
        # Normalize table text
        table_text = self._normalize_text(table_text)

        for section in sections:
            # First try to match by title/company (more reliable)
            title = self._normalize_text(section.get("title", ""))
            company = self._normalize_text(section.get("company", ""))

            # Check if both title and company are present in the table text
            if title and company:
                if title in table_text and company in table_text:
                    return section
            # If only title is present, use that for matching
            elif title and title in table_text:
                return section

            # As a fallback, check for strong bullet point matches
            bullet_matches = 0
            total_bullets = len(section.get("bullet_points", []))

            if total_bullets > 0:
                for bullet in section.get("bullet_points", []):
                    normalized_bullet = self._normalize_text(bullet.get("text", ""))
                    if normalized_bullet in table_text:
                        bullet_matches += 1

                # Only consider it a match if majority of bullets match (at least 50%)
                if bullet_matches / total_bullets >= 0.5:
                    return section

        return None

    def _normalize_text(self, text):
        """
        Normalize text for comparison

        Args:
            text (str): Input text

        Returns:
            str: Normalized text
        """
        # Convert to lowercase, remove extra whitespace
        return re.sub(r"\s+", " ", text.lower().strip())

    def _update_table_content(self, table, section):
        """
        Update the table content with new bullet points

        Args:
            table (ET.Element): Table XML element
            section (dict): Section with new content
        """
        # Get all rows in the table
        rows = table.findall(".//w:tr", self.namespaces)

        # Get bullet points from the section
        bullet_points = section.get("bullet_points", [])

        # Find the starting row for bullet points (skip header row if it contains title/company)
        start_row_index = 0
        first_row_text = self._extract_row_text(rows[0])
        if (
            section.get("title", "") in first_row_text
            or section.get("company", "") in first_row_text
        ):
            start_row_index = 1

        # Get bullet point rows
        bullet_rows = rows[start_row_index:]

        # Update or remove existing bullet point rows
        for i, row in enumerate(bullet_rows):
            if i < len(bullet_points):
                # Update existing row with new bullet point
                self._update_row_content(row, bullet_points[i]["text"])
            else:
                # Remove extra rows
                table.remove(row)

        # Add new rows if needed
        for i in range(len(bullet_rows), len(bullet_points)):
            # Clone the last existing bullet point row as template
            template_row = ET.fromstring(ET.tostring(bullet_rows[-1]))
            # Update content
            self._update_row_content(template_row, bullet_points[i]["text"])
            # Add to table
            table.append(template_row)

    def _update_row_content(self, row, new_text):
        """
        Update a single row's content while preserving formatting

        Args:
            row (ET.Element): Row XML element
            new_text (str): New text content
        """
        # Find the cell (should be first/only cell in bullet point rows)
        cell = row.find(".//w:tc", self.namespaces)
        if cell is None:
            return

        # Find existing paragraph
        paragraph = cell.find(".//w:p", self.namespaces)
        if paragraph is None:
            return

        # Preserve paragraph properties (pPr) if they exist
        p_pr = paragraph.find(".//w:pPr", self.namespaces)

        # Find existing run properties (rPr) to preserve font settings
        existing_run = paragraph.find(".//w:r", self.namespaces)
        r_pr = (
            existing_run.find(".//w:rPr", self.namespaces)
            if existing_run is not None
            else None
        )

        # Create new paragraph with preserved properties
        new_paragraph = ET.Element(f"{{{self.namespaces['w']}}}p")
        if p_pr is not None:
            new_paragraph.append(p_pr)

        # Create run with new text and preserved formatting
        run = ET.Element(f"{{{self.namespaces['w']}}}r")

        # Add run properties if they exist
        if r_pr is not None:
            run.append(r_pr)

        text = ET.Element(f"{{{self.namespaces['w']}}}t")
        text.text = new_text
        run.append(text)
        new_paragraph.append(run)

        # Replace old paragraph with new one
        cell.remove(paragraph)
        cell.append(new_paragraph)

    def _extract_row_text(self, row):
        """
        Extract text content from a table row

        Args:
            row (ET.Element): Row XML element

        Returns:
            str: Combined text content of the row
        """
        text_elements = row.findall(".//w:t", self.namespaces)
        return " ".join(
            [t.text.strip() for t in text_elements if t.text and t.text.strip()]
        )

    def _create_docx(self, temp_dir, output_path):
        """
        Rezip the modified files into a new DOCX

        Args:
            temp_dir (str): Temporary directory with extracted files
            output_path (str): Path to save the new DOCX
        """
        # Create a new zip file (DOCX)
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zipf:
            # Walk through all files in the temp directory
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    # Calculate the arcname (path within the zip)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zipf.write(file_path, arcname)


def main():
    import sys

    # Check for correct number of arguments
    if len(sys.argv) != 4:
        print(
            "Usage: python script.py <original_docx> <tailored_schema_json> <output_dir>"
        )
        sys.exit(1)

    original_resume = sys.argv[1]
    schema_path = sys.argv[2]
    base_output_dir = sys.argv[3]

    # Create timestamped output directory
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = os.path.join(base_output_dir, timestamp)
    os.makedirs(output_dir, exist_ok=True)

    # Load schema
    with open(schema_path, "r") as f:
        tailored_schema = json.load(f)

    # Initialize and tailor the document
    tailor = DocXMLModifier(original_resume)
    output_path = tailor.modify_docx(tailored_schema, output_dir)

    # Copy schema to output directory for reference
    import shutil

    schema_output = os.path.join(output_dir, "schema.json")

    shutil.copy2(schema_path, schema_output)

    print(f"Tailored resume saved to {output_path}")
    print(f"Schema copied to {schema_output}")


if __name__ == "__main__":
    main()
