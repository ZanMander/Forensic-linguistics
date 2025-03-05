# Forensic-linguistics
Various application serving as forensic linguistic instruments


Overview

This application draws on research methods described in Chapter 5 of Clare Johnson’s PhD thesis:

Johnson, C. (2023). *Understanding academic integrity and plagiarism in the digital age: Can digital forensics techniques help prevent and detect academic misconduct?* Doctoral thesis, University of South Wales.

Special thanks to Dr. Johnson for the foundational ideas and techniques that made this project possible.


This application is a Streamlit-based tool designed to analyze revision history and author contributions in Microsoft Word (.docx) documents. The tool extracts XML data from .docx files, parses revision tracking information, and provides insights into editing patterns, formatting changes, and author contributions.

 Features
- Extract Document XML: Unzips and extracts XML files from .docx documents.
- Parse Revision History: Identifies insertions, deletions, formatting changes, and tracks authors and timestamps.
- Analyze Typing Patterns: Detects possible copy-pasting or manual typing patterns.
- Visualize RSID Contributions: Maps RSID (Revision Save ID) metadata to understand document evolution.
- User-Friendly Interface: Uses Streamlit for an interactive and intuitive user experience.

Installation
To install and run the application, follow these steps:

Prerequisites
Ensure you have Python installed (recommended version: Python 3.8+).

Install Required Packages
```sh
pip install streamlit pandas matplotlib seaborn numpy
```

Running the Application
```sh
streamlit run app.py
```

How It Works
1. Upload a .docx file: The tool extracts XML files embedded in the document.
2. Extract and parse revision data: Detects edits, formatting changes, and authors.
3. Analyze RSID metadata: Understands document progression and typing behavior.
4. Display insights: Provides statistical data and visualizations.

 File Structure
- `extract_docx_xml(docx_path)`: Extracts XML files from .docx.
- `parse_revision_xml(document_xml, core_xml, settings_xml)`: Parses revision tracking data.
- `parse_document_xml(xml_data)`: Extracts text and metadata for RSID analysis.
- `analyze_typing_patterns(rsid_metadata, rsid_timeline)`: Identifies typing vs copy-pasting behavior.

Usage
- Open the Streamlit application.
- Upload a .docx file.
- View detailed insights and visualizations of document revisions.

Dependencies
- `streamlit`
- `zipfile`
- `xml.etree.ElementTree`
- `pandas`
- `matplotlib`
- `seaborn`
- `numpy`

License
This project is licensed under the MIT License.

Author and Developer:
Zander Janse Van Rensburg &  Simphiwe Nhlapo

Link : https://forensic-linguistics-9yvvfucxcvzfxnjilxvg7j.streamlit.app/

