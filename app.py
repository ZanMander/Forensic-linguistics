import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
import os
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import random

# ✅ Function to extract XML content from a .docx file
def extract_docx_xml(docx_path):
    """Extracts the document.xml content from a .docx file."""
    with zipfile.ZipFile(docx_path, 'r') as docx_zip:
        with docx_zip.open('word/document.xml') as xml_file:
            return xml_file.read().decode('utf-8')

# ✅ Function to parse document XML and extract RSID values
def parse_document_xml(xml_data):
    """Parses the XML to extract text runs and their RSIDs with color coding."""
    root = ET.fromstring(xml_data)
    namespace = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    runs_data = []
    rsid_colors = {}  # Store colors for each RSID

    for paragraph in root.findall('.//w:p', namespace):
        rsidR = paragraph.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidR', 'Unknown')
        
        # Assign a random color if RSID is new
        if rsidR not in rsid_colors:
            rsid_colors[rsidR] = "#{:06x}".format(random.randint(0, 0xFFFFFF))

        for run in paragraph.findall('.//w:r', namespace):
            text_element = run.find('w:t', namespace)
            if text_element is not None:
                runs_data.append((text_element.text, rsidR, rsid_colors[rsidR]))
    
    return runs_data, rsid_colors

# ✅ Function to compute RSID statistics
def compute_rsid_stats(runs_data):
    """Counts the number of words associated with each RSID."""
    rsid_wordcount = defaultdict(int)
    for text, rsid, _ in runs_data:
        word_count = len(text.split())  # Counting words in the text
        rsid_wordcount[rsid] += word_count
    return dict(rsid_wordcount)

# ✅ Function to generate a bar chart
def create_rsid_bar_chart(rsid_wordcount):
    """Generates a bar chart of RSID word counts."""
    plt.figure(figsize=(10, 5))
    plt.bar(rsid_wordcount.keys(), rsid_wordcount.values(), color='skyblue')
    plt.xlabel("RSID")
    plt.ylabel("Word Count")
    plt.title("RSID Word Count Distribution")
    plt.xticks(rotation=45, ha="right")
    plt.tight_layout()

    bar_chart_path = "rsid_bar_chart.png"
    plt.savefig(bar_chart_path)
    return bar_chart_path

# ✅ Function to generate a heatmap
def create_rsid_heatmap(rsid_wordcount):
    """Generates a heatmap for RSID word counts."""
    fig, ax = plt.subplots(figsize=(10, 5))
    rsid_values = list(rsid_wordcount.keys())
    word_counts = list(rsid_wordcount.values())

    sns.heatmap([word_counts], cmap="coolwarm", annot=True, fmt="d", xticklabels=rsid_values, yticklabels=["Word Count"], ax=ax)
    
    ax.set_xlabel("RSID")
    ax.set_title("RSID Word Count Heatmap")
    plt.xticks(rotation=45, ha="right")

    heatmap_path = "rsid_heatmap.png"
    fig.savefig(heatmap_path)
    return heatmap_path

# ✅ Function to generate the RSID report
def generate_report(docx_file):
    """Extracts, parses, computes statistics, generates charts, and creates an HTML report."""
    xml_data = extract_docx_xml(docx_file)
    runs_data, rsid_colors = parse_document_xml(xml_data)
    rsid_stats = compute_rsid_stats(runs_data)

    # Generate both bar chart and heatmap
    bar_chart_path = create_rsid_bar_chart(rsid_stats)
    heatmap_path = create_rsid_heatmap(rsid_stats)

    # Generate HTML report
    report_path = "rsid_report.html"
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("<html><head><title>RSID Report</title></head><body>")
        f.write("<h1>RSID Word Count Statistics</h1>")
        f.write("<table border='1'><tr><th>RSID</th><th>Word Count</th></tr>")
        for rsid, count in rsid_stats.items():
            f.write(f"<tr><td>{rsid}</td><td>{count}</td></tr>")
        f.write("</table>")
        f.write("<h2>Bar Chart</h2>")
        f.write(f"<img src='{bar_chart_path}' width='600px'>")
        f.write("<h2>Heatmap</h2>")
        f.write(f"<img src='{heatmap_path}' width='600px'>")
        f.write("</body></html>")
    
    return rsid_stats, runs_data, rsid_colors, bar_chart_path, heatmap_path, report_path

# ✅ Streamlit UI
st.title("📄 Word Document RSID Analyzer")
st.write("Upload a .docx file to analyze RSID statistics, view extracted text, and generate a report.")

uploaded_file = st.file_uploader("📂 Upload a .docx file", type=["docx"])

if uploaded_file:
    st.success("✅ File uploaded successfully!")

    temp_path = "temp.docx"
    with open(temp_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # Run analysis
    rsid_stats, runs_data, rsid_colors, bar_chart_path, heatmap_path, report_path = generate_report(temp_path)

    # Display RSID Statistics
    st.subheader("📊 RSID Word Count Statistics")
    stats_df = pd.DataFrame(rsid_stats.items(), columns=["RSID", "Word Count"])
    st.dataframe(stats_df, use_container_width=True)

    # **Text-based RSID visualization**
    st.subheader("📝 RSID-Based Text Visualization")

    rsid_text_html = "<div style='font-family: Arial, sans-serif; line-height: 1.5;'>"
    for text, rsid, color in runs_data:
        rsid_text_html += f"<span style='background-color:{color}; padding:3px; margin:2px; border-radius:3px;' title='RSID: {rsid}'>{text} </span>"
    rsid_text_html += "</div>"

    st.markdown(rsid_text_html, unsafe_allow_html=True)

    # **Display Bar Chart**
    st.subheader("📉 RSID Word Count Bar Chart")
    st.image(bar_chart_path, use_column_width=True)

    # **Display Heatmap**
    st.subheader("🔥 RSID Heatmap Visualization")
    st.image(heatmap_path, use_column_width=True)

    # **Download RSID Statistics (CSV)**
    csv_data = stats_df.to_csv(index=False).encode('utf-8')
    st.download_button("📥 Download RSID Statistics (CSV)", csv_data, "rsid_statistics.csv", "text/csv")

    # **Download Full RSID Report (HTML)**
    with open(report_path, "r", encoding="utf-8") as f:
        html_data = f.read()
    st.download_button("📄 Download Full RSID Report (HTML)", html_data, "rsid_report.html", "text/html")

    os.remove(temp_path)
