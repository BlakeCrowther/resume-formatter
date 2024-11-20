import os
import json
import zipfile
import xml.etree.ElementTree as ET
import re


class DocxXMLParser:
    def __init__(self, docx_path):
        """
        Initialize parser for a DOCX file by extracting its XML contents

        Args:
            docx_path (str): Path to the .docx file
        """
        self.docx_path = docx_path
        self.namespaces = {
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        }

    def extract_document_structure(self, output_dir):
        """
        Extract comprehensive document structure by parsing DOCX XML

        Args:
            output_dir (str): Directory to save extracted structure

        Returns:
            dict: Comprehensive document structure
        """
        try:
            # Ensure output directory exists
            os.makedirs(output_dir, exist_ok=True)

            # Open the DOCX file as a zip
            with zipfile.ZipFile(self.docx_path) as docx:
                # Extract document.xml (main document content)
                document_xml = docx.read("word/document.xml")
                root = ET.fromstring(document_xml)

                # Comprehensive structure to collect
                structure = {
                    # "paragraphs": [],
                    # "text_boxes": [],
                    "tables": [],
                    # "images": [],
                    "styles": {},
                }

                # Parse paragraphs
                # self._parse_paragraphs(root, structure)

                # Parse text boxes (drawing objects)
                # self._parse_text_boxes(docx, root, structure)

                # Parse tables
                self._parse_tables(root, structure)

                # Parse images
                # self._parse_images(docx, root, structure)

                # Parse styles
                styles_xml = docx.read("word/styles.xml")
                self._parse_styles(styles_xml, structure)

            # Save structure to JSON
            structure_path = os.path.join(output_dir, "document_structure.json")
            with open(structure_path, "w", encoding="utf-8") as f:
                json.dump(structure, f, indent=2, ensure_ascii=False)

            return structure

        except Exception as e:
            print(f"Error extracting document structure: {e}")
            raise

    def _parse_paragraphs(self, root, structure):
        """
        Extract paragraph information

        Args:
            root (ET.Element): Root XML element
            structure (dict): Structure dictionary to populate
        """
        paragraphs = root.findall(".//w:p", self.namespaces)
        for para in paragraphs:
            para_text = self._extract_paragraph_text(para)
            para_style = para.find(".//w:pStyle", self.namespaces)

            paragraph_info = {
                "text": para_text,
                "style": (
                    para_style.get(f'{{{self.namespaces["w"]}}}val')
                    if para_style is not None
                    else None
                ),
                "char_count": len(para_text),
            }
            structure["paragraphs"].append(paragraph_info)

    def _parse_text_boxes(self, docx, root, structure):
        """
        Extract text box information

        Args:
            docx (zipfile.ZipFile): DOCX file as a zip
            root (ET.Element): Root XML element
            structure (dict): Structure dictionary to populate
        """
        # Find all drawing elements (potential text boxes)
        drawings = root.findall(".//w:drawing", self.namespaces)
        for drawing in drawings:
            textbox = drawing.find(".//a:textbox", self.namespaces)
            if textbox is not None:
                # Extract text from text box
                textbox_text = self._extract_textbox_text(textbox)

                # Try to extract positioning information
                ext = drawing.find(".//a:ext", self.namespaces)
                position = drawing.find(".//a:off", self.namespaces)

                text_box_info = {
                    "text": textbox_text,
                    "char_count": len(textbox_text),
                    "width": (
                        float(ext.get("cx", 0)) / 360000 if ext is not None else None
                    ),
                    "height": (
                        float(ext.get("cy", 0)) / 360000 if ext is not None else None
                    ),
                    "x": (
                        float(position.get("x", 0)) / 360000
                        if position is not None
                        else None
                    ),
                    "y": (
                        float(position.get("y", 0)) / 360000
                        if position is not None
                        else None
                    ),
                }
                structure["text_boxes"].append(text_box_info)

    def _parse_tables(self, root, structure):
        """
        Extract table information

        Args:
            root (ET.Element): Root XML element
            structure (dict): Structure dictionary to populate
        """
        tables = root.findall(".//w:tbl", self.namespaces)
        for table in tables:
            table_info = {"rows": [], "total_rows": 0, "total_columns": 0}

            rows = table.findall(".//w:tr", self.namespaces)
            table_info["total_rows"] = len(rows)

            for row in rows:
                cells = row.findall(".//w:tc", self.namespaces)
                table_info["total_columns"] = max(
                    table_info["total_columns"], len(cells)
                )

                row_text = []
                for cell in cells:
                    cell_text = self._extract_paragraph_text(cell)
                    row_text.append(cell_text)

                table_info["rows"].append(row_text)

            structure["tables"].append(table_info)

    def _parse_images(self, docx, root, structure):
        """
        Extract image references

        Args:
            docx (zipfile.ZipFile): DOCX file as a zip
            root (ET.Element): Root XML element
            structure (dict): Structure dictionary to populate
        """
        images = root.findall(".//pic:pic", self.namespaces)
        for image in images:
            name = image.find(".//pic:cNvPr", self.namespaces)
            ext = image.find(".//a:ext", self.namespaces)

            image_info = {
                "name": name.get("name") if name is not None else None,
                "width": float(ext.get("cx", 0)) / 360000 if ext is not None else None,
                "height": float(ext.get("cy", 0)) / 360000 if ext is not None else None,
            }
            structure["images"].append(image_info)

    def _parse_styles(self, styles_xml, structure):
        """
        Extract document styles

        Args:
            styles_xml (bytes): XML content of styles
            structure (dict): Structure dictionary to populate
        """
        root = ET.fromstring(styles_xml)
        styles = root.findall(".//w:style", self.namespaces)

        for style in styles:
            style_id = style.get(f'{{{self.namespaces["w"]}}}styleId')
            name = style.find(".//w:name", self.namespaces)

            if style_id and name is not None:
                structure["styles"][style_id] = {
                    "name": name.get(f'{{{self.namespaces["w"]}}}val'),
                    "type": style.get(f'{{{self.namespaces["w"]}}}type'),
                }

    def _extract_paragraph_text(self, element):
        """
        Extract text from a paragraph or cell element

        Args:
            element (ET.Element): XML element to extract text from

        Returns:
            str: Extracted text
        """
        texts = element.findall(".//w:t", self.namespaces)
        return "".join([t.text for t in texts if t.text]) if texts else ""

    def _extract_textbox_text(self, textbox):
        """
        Extract text from a text box element

        Args:
            textbox (ET.Element): Text box XML element

        Returns:
            str: Extracted text
        """
        texts = textbox.findall(".//w:t", self.namespaces)
        return "".join([t.text for t in texts if t.text]) if texts else ""


def main():
    import sys

    if len(sys.argv) < 2:
        print("Usage: python docx-xml-parser.py <docx_path>")
        sys.exit(1)

    docx_path = sys.argv[1]
    output_dir = sys.argv[2]

    parser = DocxXMLParser(docx_path)
    structure = parser.extract_document_structure(output_dir)

    # Print some insights
    # print(f"Total Paragraphs: {len(structure['paragraphs'])}")
    # print(f"Total Text Boxes: {len(structure['text_boxes'])}")
    print(f"Total Tables: {len(structure['tables'])}")
    # print(f"Total Images: {len(structure['images'])}")


if __name__ == "__main__":
    main()
