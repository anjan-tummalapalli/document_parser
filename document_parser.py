import os
import csv
import json
import argparse
from typing import Any, Dict, List, Optional

# Supported dependencies:
# pip3 install pypdf2 python-docx pandas openpyxl beautifulsoup4

try:
    import PyPDF2
except ImportError:
    PyPDF2 = None

try:
    import docx
except ImportError:
    docx = None

try:
    import pandas as pd
except ImportError:
    pd = None

try:
    from bs4 import BeautifulSoup
except ImportError:
    BeautifulSoup = None


class DocumentParser:
    """A universal document parser for extracting text and data from various file formats."""

    @staticmethod
    def parse_txt(file_path: str) -> str:
        with open(file_path, "r", encoding="utf-8") as f:
            return f.read()

    @staticmethod
    def parse_csv(file_path: str) -> str:
        content = []
        with open(file_path, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            for row in reader:
                content.append(", ".join(row))
        return "\n".join(content)

    @staticmethod
    def parse_json(file_path: str) -> str:
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return json.dumps(data, indent=2)

    @staticmethod
    def parse_pdf(file_path: str) -> str:
        if not PyPDF2:
            return "[Error: PyPDF2 is not installed. Run 'pip install PyPDF2']"
        
        text = []
        with open(file_path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text.append(page_text)
        return "\n\n".join(text)

    @staticmethod
    def parse_docx(file_path: str) -> str:
        if not docx:
            return "[Error: python-docx is not installed. Run 'pip install python-docx']"
        
        doc = docx.Document(file_path)
        return "\n".join([para.text for para in doc.paragraphs])

    @staticmethod
    def parse_excel(file_path: str) -> str:
        if not pd:
            return "[Error: pandas and openpyxl are not installed. Run 'pip install pandas openpyxl']"
        
        # Read all sheets into a dictionary of dataframes
        df_dict = pd.read_excel(file_path, sheet_name=None)
        output = []
        for sheet_name, df in df_dict.items():
            output.append(f"--- Sheet: {sheet_name} ---")
            output.append(df.to_string(index=False))
        return "\n\n".join(output)

    @staticmethod
    def parse_html(file_path: str) -> str:
        if not BeautifulSoup:
            return "[Error: beautifulsoup4 is not installed. Run 'pip install beautifulsoup4']"
        
        with open(file_path, "r", encoding="utf-8") as f:
            soup = BeautifulSoup(f.read(), "html.parser")
            return soup.get_text(separator="\n", strip=True)

    @classmethod
    def parse(cls, file_path: str) -> str:
        """Determines the file type by extension and parses it."""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")

        ext = os.path.splitext(file_path)[1].lower()

        if ext == ".txt":
            return cls.parse_txt(file_path)
        elif ext == ".csv":
            return cls.parse_csv(file_path)
        elif ext == ".json":
            return cls.parse_json(file_path)
        elif ext == ".pdf":
            return cls.parse_pdf(file_path)
        elif ext in [".docx", ".doc"]:
            return cls.parse_docx(file_path)
        elif ext in [".xlsx", ".xls"]:
            return cls.parse_excel(file_path)
        elif ext in [".html", ".htm"]:
            return cls.parse_html(file_path)
        else:
            raise ValueError(f"Unsupported file format: {ext}")


def main():
    parser = argparse.ArgumentParser(description="Universal Document Parser")
    parser.add_argument("filepath", type=str, help="Path to the document to parse")
    args = parser.parse_args()

    try:
        print(f"Parsing document: {args.filepath}...\n")
        parsed_text = DocumentParser.parse(args.filepath)
        print("--- EXTRACTED CONTENT ---")
        print(parsed_text)
        print("-------------------------")
    except Exception as e:
        print(f"Error parsing document: {e}")


if __name__ == "__main__":
    main()
