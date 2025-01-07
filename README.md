# DocxWriterPy: A Python Tool for DOCX Mail Merge and Manipulation
**DocxWriterPy** is a Python-based utility designed to simplify working with Microsoft Word (DOCX) documents. This tool is particularly useful for organizations that restrict the use of external libraries, as it provides a self-contained solution for automating document processing without needing to install the `python-docx` package online. By implementing mail merging capabilities, DocxWriter enables users to efficiently manage data processing tasks, enhancing productivity and reducing manual errors.

# Why DocxWriter?
The development of **DocxWriterPy** stems from the constraints imposed by my company, which does not permit the use of online installations for the python-docx package. This limitation necessitated the creation of a robust alternative that can effectively implement mail merging with DOCX files, thereby automating data processing tasks.

# Features
* **Mail Merge**: Automate document creation by replacing placeholders with dynamic data.
* **Image Extraction**: Extract all embedded images from a DOCX file to a folder.
* **Image Replacement**: Replace placeholder images directly inside a DOCX file.
* **Text Manipulation**: Find and replace text in paragraphs, headers, or textboxes.
* **Convert to PDF**: Save updated DOCX files as PDFs.
* **Textbox Content Extraction**: Retrieve text from all textboxes in the document, addressing a limitation of `python-docx`.

# Installation
1. Clone the repository:
``` bash
git clone https://github.com/your-username/docxwriter.git
cd docxwriter
```

2. Install required dependencies:
``` bash
pip install lxml pywin32
```

3. Ensure **Microsoft Word** is installed (required for PDF conversion).

# Example Usage
Refer to the `example.ipynb` file included in this repository for practical examples demonstrating how to utilize DocxWriter effectively. The notebook provides step-by-step instructions on implementing various features, ensuring you can harness the full potential of this tool. With DocxWriter, you can efficiently manage your DOCX documents without relying on external packages, making it an essential tool for professionals who require flexibility and automation in their document handling processes.