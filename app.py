import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from typing import Tuple, Dict, List, Optional, Any
import numpy as np
import colorsys
import matplotlib.dates as mdates
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
import re
import base64
from io import BytesIO
import os
from PIL import Image
import tempfile






class WordDocumentAnalyzer:
    """A class to analyze Word documents for revision history, typing patterns, and academic integrity."""
    
    def __init__(self, docx_path: str):
        """Initialize analyzer with a docx file path."""
        self.docx_path = docx_path
        self.namespace = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
            'dc': 'http://purl.org/dc/elements/1.1/',
            'dcterms': 'http://purl.org/dc/terms/',
            'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
            'vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
            'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
            'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
            'w15': 'http://schemas.microsoft.com/office/word/2012/wordml'
        }
        
        # Extract and store all relevant XML files
        self.xml_files = self.extract_all_xml_files()
        
        # Parse document roots
        try:
            self.document_root = ET.fromstring(self.xml_files.get('document.xml', '')) if 'document.xml' in self.xml_files else None
            self.core_root = ET.fromstring(self.xml_files.get('core.xml', '')) if 'core.xml' in self.xml_files else None
            self.settings_root = ET.fromstring(self.xml_files.get('settings.xml', '')) if 'settings.xml' in self.xml_files else None
            self.app_root = ET.fromstring(self.xml_files.get('app.xml', '')) if 'app.xml' in self.xml_files else None
            self.styles_root = ET.fromstring(self.xml_files.get('styles.xml', '')) if 'styles.xml' in self.xml_files else None
        except ET.ParseError as e:
            st.error(f"Error parsing XML: {e}")
            self.document_root = self.core_root = self.settings_root = self.app_root = self.styles_root = None
    
    def extract_all_xml_files(self) -> Dict[str, str]:
        """Extracts all relevant XML files from a .docx file."""
        xml_files = {}
        file_paths = [
            ('document.xml', 'word/document.xml'),
            ('core.xml', 'docProps/core.xml'),
            ('settings.xml', 'word/settings.xml'),
            ('app.xml', 'docProps/app.xml'),
            ('styles.xml', 'word/styles.xml'),
            ('fontTable.xml', 'word/fontTable.xml'),
            ('theme1.xml', 'word/theme/theme1.xml'),
            ('numbering.xml', 'word/numbering.xml'),
            ('webSettings.xml', 'word/webSettings.xml'),
            ('comments.xml', 'word/comments.xml')
        ]
        
        try:
            with zipfile.ZipFile(self.docx_path, 'r') as docx_zip:
                # Extract all standard files
                for file_name, path in file_paths:
                    try:
                        with docx_zip.open(path) as xml_file:
                            xml_files[file_name] = xml_file.read().decode('utf-8')
                    except KeyError:
                                st.info(f"{path} not found in the document")
                
                # Additionally extract any custom XML files
                custom_xml_files = [f for f in docx_zip.namelist() if f.startswith('customXml/')]
                for custom_path in custom_xml_files:
                    try:
                        with docx_zip.open(custom_path) as xml_file:
                            file_name = os.path.basename(custom_path)
                            xml_files[f'custom_{file_name}'] = xml_file.read().decode('utf-8')
                    except Exception as e:
                            st.warning(f"Error extracting {custom_path}: {e}")
        except Exception as e:
                st.error(f"Error opening docx file: {e}")
            
        return xml_files
    
    def parse_metadata(self) -> Dict[str, Any]:
        """Parse comprehensive document metadata from core.xml and app.xml."""
        metadata = {
            'title': 'Unknown',
            'creator': 'Unknown',
            'last_modified_by': 'Unknown',
            'created': 'Unknown',
            'modified': 'Unknown',
            'company': 'Unknown',
            'application': 'Unknown',
            'app_version': 'Unknown',
            'total_edit_time': 0,
            'last_printed': 'Unknown',
            'revision': 0,
            'content_status': 'Unknown',
            'template': 'Unknown',
            'subject': 'Unknown',
            'category': 'Unknown'
        }
        
        # Parse core.xml
        if self.core_root is not None:
            try:
                # Extract basic metadata
                elements_to_extract = [
                    ('title', './/dc:title', self.namespace),
                    ('creator', './/dc:creator', self.namespace),
                    ('last_modified_by', './/cp:lastModifiedBy', self.namespace),
                    ('created', './/dcterms:created', self.namespace),
                    ('modified', './/dcterms:modified', self.namespace),
                    ('revision', './/cp:revision', self.namespace),
                    ('subject', './/dc:subject', self.namespace),
                    ('category', './/cp:category', self.namespace),
                    ('content_status', './/cp:contentStatus', self.namespace)
                ]
                
                for meta_key, xpath, ns in elements_to_extract:
                    element = self.core_root.find(xpath, ns)
                    if element is not None and element.text:
                        metadata[meta_key] = element.text
                        
                # Convert numeric values
                if metadata['revision'] != 'Unknown':
                    metadata['revision'] = int(metadata['revision'])
            except Exception as e:
                st.warning(f"Error parsing core.xml: {str(e)}")
        
        # Parse app.xml
        if self.app_root is not None:
            try:
                elements_to_extract = [
                    ('company', './/ep:Company', self.namespace),
                    ('application', './/ep:Application', self.namespace),
                    ('app_version', './/ep:AppVersion', self.namespace),
                    ('total_edit_time', './/ep:TotalTime', self.namespace),
                    ('last_printed', './/ep:LastPrinted', self.namespace),
                    ('template', './/ep:Template', self.namespace)
                ]
                
                for meta_key, xpath, ns in elements_to_extract:
                    element = self.app_root.find(xpath, ns)
                    if element is not None and element.text:
                        metadata[meta_key] = element.text
                
                # Convert numeric values
                if metadata['total_edit_time'] != 'Unknown':
                    metadata['total_edit_time'] = int(metadata['total_edit_time'])
            except Exception as e:
                st.warning(f"Error parsing app.xml: {str(e)}")
                
        # Add document statistics if available
        if self.app_root is not None:
            try:
                stats_elements = [
                    ('pages', './/ep:Pages', self.namespace),
                    ('words', './/ep:Words', self.namespace),
                    ('characters', './/ep:Characters', self.namespace),
                    ('paragraphs', './/ep:Paragraphs', self.namespace)
                ]
                
                for meta_key, xpath, ns in elements_to_extract:
                    element = self.app_root.find(xpath, ns)
                    if element is not None and element.text:
                        metadata[meta_key] = int(element.text)
            except Exception:
                pass
        
        return metadata
    

    def detect_font_inconsistencies(self) -> Dict[str, Any]:
        """Detect inconsistencies in fonts and formatting that may indicate copy-paste."""
        if not self.document_root:
            return {'detected': False, 'details': {}}
        
        # Track font properties across the document
        fonts = set()
        font_sizes = set()
        languages = set()
        font_distribution = defaultdict(int)
        size_distribution = defaultdict(int)
        lang_distribution = defaultdict(int)
        
        # Analyze each run for font properties
        for run in self.document_root.findall('.//w:r', self.namespace):
            run_props = run.find('.//w:rPr', self.namespace)
            if run_props is None:
                continue
                
            # Check font
            font_element = run_props.find('.//w:rFonts', self.namespace)
            if font_element is not None:
                for font_attr in ['ascii', 'hAnsi', 'cs', 'eastAsia']:
                    font_name = font_element.attrib.get(f'w:{font_attr}', None)
                    if font_name:
                        fonts.add(font_name)
                        font_distribution[font_name] += 1
            
            # Check font size
            size_element = run_props.find('.//w:sz', self.namespace)
            if size_element is not None and 'val' in size_element.attrib:
                size_val = size_element.attrib['val']
                font_sizes.add(size_val)
                size_distribution[size_val] += 1
            
            # Check language
            lang_element = run_props.find('.//w:lang', self.namespace)
            if lang_element is not None:
                for lang_attr in ['val', 'eastAsia', 'bidi']:
                    lang_val = lang_element.attrib.get(f'w:{lang_attr}', None)
                    if lang_val:
                        languages.add(lang_val)
                        lang_distribution[lang_val] += 1
        
        # Convert to frequencies
        total_runs = sum(font_distribution.values()) or 1  # Avoid division by zero
        font_freq = {font: count/total_runs for font, count in font_distribution.items()}
        size_freq = {size: count/total_runs for size, count in size_distribution.items()}
        lang_freq = {lang: count/total_runs for lang, count in lang_distribution.items()}
        
        # Detect inconsistencies
        unusual_fonts = [font for font, freq in font_freq.items() if freq < 0.05]
        unusual_sizes = [size for size, freq in size_freq.items() if freq < 0.05]
        unusual_langs = [lang for lang, freq in lang_freq.items() if freq < 0.05]
        
        # Determine if inconsistencies are significant
        has_inconsistencies = (
            len(fonts) > 2 or  # More than two different fonts
            len(font_sizes) > 3 or  # More than three different sizes
            len(languages) > 2 or  # More than two different languages
            len(unusual_fonts) > 0 or  # Any unusual fonts
            len(unusual_sizes) > 0 or  # Any unusual sizes
            len(unusual_langs) > 0  # Any unusual languages
        )
        
        # Calculate severity score (0.0 to 1.0)
        severity = min(1.0, (
            (len(fonts) - 1) * 0.2 +
            (len(font_sizes) - 2) * 0.1 +
            (len(languages) - 1) * 0.2 +
            len(unusual_fonts) * 0.2 +
            len(unusual_sizes) * 0.1 +
            len(unusual_langs) * 0.2
        ))
        
        return {
            'detected': has_inconsistencies,
            'severity': severity,
            'details': {
                'fonts': list(fonts),
                'font_sizes': list(font_sizes),
                'languages': list(languages),
                'unusual_fonts': unusual_fonts,
                'unusual_sizes': unusual_sizes,
                'unusual_languages': unusual_langs,
                'font_distribution': dict(font_distribution),
                'size_distribution': dict(size_distribution),
                'language_distribution': dict(lang_distribution)
            }
        }


    def check_revision_tracking_status(self) -> Dict[str, Any]:
        """Check if revision tracking was enabled or disabled in the document."""
        tracking_status = {
            'tracking_enabled': False,
            'track_revisions': False,
            'track_moves': False,
            'track_format_changes': False,
            'rsidRoot': 'Unknown'
        }
        
        if self.settings_root is not None:
            try:
                # Check if tracking is enabled
                track_revisions = self.settings_root.find('.//w:trackRevisions', self.namespace)
                tracking_status['track_revisions'] = track_revisions is not None
                
                # Check if format changes are tracked
                track_format_changes = self.settings_root.find('.//w:trackFormatting', self.namespace)
                tracking_status['track_format_changes'] = track_format_changes is not None
                
                # Check if moves are tracked
                track_moves = self.settings_root.find('.//w:trackMoves', self.namespace)
                tracking_status['track_moves'] = track_moves is not None
                
                # Get the rsidRoot value
                rsid_root = self.settings_root.find('.//w:rsidRoot', self.namespace)
                if rsid_root is not None and 'val' in rsid_root.attrib:
                    tracking_status['rsidRoot'] = rsid_root.attrib['val']
                
                # Set overall tracking status
                tracking_status['tracking_enabled'] = any([
                    tracking_status['track_revisions'],
                    tracking_status['track_format_changes'],
                    tracking_status['track_moves']
                ])
            except Exception as e:
                st.warning(f"Error checking revision tracking status: {str(e)}")
                
        # Check for RSIDs in the document to verify if tracking was used
        if self.document_root is not None and tracking_status['rsidRoot'] == 'Unknown':
            try:
                rsid_attributes = [
                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidR',
                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidRPr',
                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidP'
                ]
                
                # Check first few paragraphs for RSIDs
                for para in self.document_root.findall('.//w:p', self.namespace):
                    for attr in rsid_attributes:
                        if attr in para.attrib:
                            tracking_status['tracking_enabled'] = True
                            return tracking_status
            except Exception:
                pass
                
        return tracking_status
    
    def parse_document_history(self) -> List[Dict]:
        """Parse document revision history and return a chronological timeline of edits."""
        if not self.document_root:
            return []
            
        history_events = []
        
        # Extract creation date from core.xml
        if self.core_root is not None:
            try:
                created_element = self.core_root.find('.//dcterms:created', self.namespace)
                if created_element is not None and created_element.text:
                    created_date = created_element.text
                    creator = self.core_root.find('.//dc:creator', self.namespace)
                    creator_text = creator.text if creator is not None else 'Unknown'
                    
                    history_events.append({
                        'date': created_date,
                        'author': creator_text,
                        'event_type': 'Document Creation',
                        'content': 'Document was created',
                        'detail': ''
                    })
            except Exception as e:
                st.warning(f"Error extracting creation date: {str(e)}")
        
        # Process insertions
        for ins in self.document_root.findall('.//w:ins', self.namespace):
            author = ins.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', 'Unknown')
            date = ins.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', 'Unknown')
            
            inserted_text = ''.join(t.text for r in ins.findall('.//w:r', self.namespace) 
                                  for t in r.findall('w:t', self.namespace) if t is not None and t.text)
            
            # Truncate long text
            display_text = inserted_text[:50] + "..." if len(inserted_text) > 50 else inserted_text
            
            history_events.append({
                'date': date,
                'author': author,
                'event_type': 'Insertion',
                'content': f'Added: "{display_text}"',
                'detail': inserted_text
            })
        
        # Process deletions
        for deletion in self.document_root.findall('.//w:del', self.namespace):
            author = deletion.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', 'Unknown')
            date = deletion.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', 'Unknown')
            
            deleted_text = ''.join(t.text for r in deletion.findall('.//w:r', self.namespace) 
                                  for t in r.findall('w:delText', self.namespace) if t is not None and t.text)
            
            # Truncate long text
            display_text = deleted_text[:50] + "..." if len(deleted_text) > 50 else deleted_text
            
            history_events.append({
                'date': date,
                'author': author,
                'event_type': 'Deletion',
                'content': f'Deleted: "{display_text}"',
                'detail': deleted_text
            })
        
        # Process formatting changes
        for fmt_change in self.document_root.findall('.//w:rPrChange', self.namespace):
            author = fmt_change.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', 'Unknown')
            date = fmt_change.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', 'Unknown')
            
            # Identify what formatting changed
            formatting_elements = [child.tag.split('}')[-1] for child in fmt_change]
            
            history_events.append({
                'date': date,
                'author': author,
                'event_type': 'Formatting',
                'content': f'Changed formatting: {", ".join(formatting_elements)}',
                'detail': str(formatting_elements)
            })
            
        # Process paragraph property changes
        for para_change in self.document_root.findall('.//w:pPrChange', self.namespace):
            author = para_change.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', 'Unknown')
            date = para_change.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', 'Unknown')
            
            # Identify what paragraph properties changed
            para_elements = [child.tag.split('}')[-1] for child in para_change]
            
            history_events.append({
                'date': date,
                'author': author,
                'event_type': 'Paragraph Formatting',
                'content': f'Changed paragraph format: {", ".join(para_elements)}',
                'detail': str(para_elements)
            })
        
        # Check for tracked moves
        for move_from in self.document_root.findall('.//w:moveFrom', self.namespace):
            author = move_from.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', 'Unknown')
            date = move_from.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', 'Unknown')
            
            moved_text = ''.join(t.text for r in move_from.findall('.//w:r', self.namespace) 
                                for t in r.findall('w:t', self.namespace) if t is not None and t.text)
            
            display_text = moved_text[:50] + "..." if len(moved_text) > 50 else moved_text
            
            history_events.append({
                'date': date,
                'author': author,
                'event_type': 'Move',
                'content': f'Moved: "{display_text}"',
                'detail': moved_text
            })
            
        # Process comments if available
        if 'comments.xml' in self.xml_files:
            try:
                comments_root = ET.fromstring(self.xml_files['comments.xml'])
                for comment in comments_root.findall('.//w:comment', self.namespace):
                    author = comment.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', 'Unknown')
                    date = comment.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', 'Unknown')
                    
                    comment_text = ''.join(t.text for p in comment.findall('.//w:p', self.namespace)
                                          for r in p.findall('.//w:r', self.namespace)
                                          for t in r.findall('w:t', self.namespace) if t is not None and t.text)
                    
                    display_text = comment_text[:50] + "..." if len(comment_text) > 50 else comment_text
                    
                    history_events.append({
                        'date': date,
                        'author': author,
                        'event_type': 'Comment',
                        'content': f'Comment: "{display_text}"',
                        'detail': comment_text
                    })
            except Exception as e:
                st.warning(f"Error parsing comments: {str(e)}")
                
        # Extract last modification date from core.xml
        if self.core_root is not None:
            try:
                modified_element = self.core_root.find('.//dcterms:modified', self.namespace)
                if modified_element is not None and modified_element.text:
                    modified_date = modified_element.text
                    last_modified_by = self.core_root.find('.//cp:lastModifiedBy', self.namespace)
                    last_modified_by_text = last_modified_by.text if last_modified_by is not None else 'Unknown'
                    
                    # Only add if it's different from the creation date
                    if not any(event['date'] == modified_date and event['event_type'] == 'Document Creation' for event in history_events):
                        history_events.append({
                            'date': modified_date,
                            'author': last_modified_by_text,
                            'event_type': 'Last Modification',
                            'content': 'Document was last saved',
                            'detail': ''
                        })
            except Exception as e:
                st.warning(f"Error extracting modification date: {str(e)}")
        
        # Sort events chronologically
        try:
            # Parse and standardize dates
            for event in history_events:
                try:
                    # Handle different date formats that might be in the document
                    if 'T' in event['date']:
                        # ISO format with timezone
                        event['parsed_date'] = datetime.fromisoformat(event['date'].replace('Z', '+00:00'))
                    else:
                        # Try a simpler format if the above fails
                        event['parsed_date'] = datetime.strptime(event['date'], '%Y-%m-%d')
                except Exception:
                    # Default to a distant past date if parsing fails
                    event['parsed_date'] = datetime(1900, 1, 1)
            
            history_events.sort(key=lambda x: x['parsed_date'])
        except Exception as e:
            st.warning(f"Error sorting timeline events: {str(e)}")
        
        return history_events
    
    def parse_rsid_data(self) -> Tuple[List, Dict, List, Dict]:
        """Parses the document XML to extract text runs and their RSIDs with enhanced metadata."""
        if not self.document_root:
            return [], {}, [], {}
            
        runs_data = []
        rsid_colors = {}
        rsid_timeline = []
        rsid_metadata = defaultdict(lambda: {
            'word_count': 0,
            'character_count': 0,
            'segment_count': 0,
            'consecutive_count': 0,
            'timestamps': set(),  # To track potential temporal information
            'authors': set(),     # To track potential authors
            'fonts': set(),       # To track font variations
            'font_sizes': set()   # To track font size variations
        })

        previous_rsid = None
        
        for paragraph in self.document_root.findall('.//w:p', self.namespace):
            para_rsid = paragraph.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidR', 'Unknown')
            para_rsid_p = paragraph.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidP', 'Unknown')
            
            # If paragraph doesn't have rsidR but has rsidP, use that
            if para_rsid == 'Unknown' and para_rsid_p != 'Unknown':
                para_rsid = para_rsid_p
                
            # Check for w14:textID attribute (potential merged document)
            text_id = paragraph.attrib.get('{http://schemas.microsoft.com/office/word/2010/wordml}textId', None)
            if text_id:
                rsid_metadata[para_rsid]['text_ids'] = rsid_metadata[para_rsid].get('text_ids', set())
                rsid_metadata[para_rsid]['text_ids'].add(text_id)
            
            if para_rsid not in rsid_colors:
                # Use golden ratio to create visually distinct colors
                hue = len(rsid_colors) * 0.618033988749895 % 1
                rgb = tuple(int(c * 255) for c in colorsys.hsv_to_rgb(hue, 0.8, 0.95))
                rsid_colors[para_rsid] = "#{:02x}{:02x}{:02x}".format(*rgb)

            rsid_timeline.append(para_rsid)
            
            if para_rsid == previous_rsid:
                rsid_metadata[para_rsid]['consecutive_count'] += 1
            previous_rsid = para_rsid
            
            # Process paragraph properties for style consistency
            para_props = paragraph.find('.//w:pPr', self.namespace)
            if para_props is not None:
                style_element = para_props.find('.//w:pStyle', self.namespace)
                if style_element is not None and 'val' in style_element.attrib:
                    style_val = style_element.attrib['val']
                    rsid_metadata[para_rsid]['styles'] = rsid_metadata[para_rsid].get('styles', set())
                    rsid_metadata[para_rsid]['styles'].add(style_val)
            
            for run in paragraph.findall('.//w:r', self.namespace):
                run_rsid = run.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidR', para_rsid)
                text_element = run.find('w:t', self.namespace)
                
                # Extract font information from run properties
                run_props = run.find('.//w:rPr', self.namespace)
                if run_props is not None:
                    # Check font
                    font_element = run_props.find('.//w:rFonts', self.namespace)
                    if font_element is not None:
                        for font_attr in ['ascii', 'hAnsi', 'cs', 'eastAsia']:
                            font_name = font_element.attrib.get(f'w:{font_attr}', None)
                            if font_name:
                                rsid_metadata[para_rsid]['fonts'].add(font_name)
                    
                    # Check font size
                    size_element = run_props.find('.//w:sz', self.namespace)
                    if size_element is not None and 'val' in size_element.attrib:
                        rsid_metadata[para_rsid]['font_sizes'].add(size_element.attrib['val'])
                
                if text_element is not None and text_element.text:
                    text = text_element.text
                    
                    # Store the run data with paragraph RSID for visualization
                    runs_data.append((text, para_rsid, rsid_colors[para_rsid]))
                    
                    # Count words accurately by splitting on whitespace
                    rsid_metadata[para_rsid]['word_count'] += len(text.split())
                    rsid_metadata[para_rsid]['character_count'] += len(text)
                    rsid_metadata[para_rsid]['segment_count'] += 1

        # Convert set values to lists for JSON serialization
        for rsid, meta in rsid_metadata.items():
            for key in ['timestamps', 'authors', 'fonts', 'font_sizes', 'text_ids', 'styles']:
                if key in meta and isinstance(meta[key], set):
                    meta[key] = list(meta[key])

        return runs_data, rsid_colors, rsid_timeline, dict(rsid_metadata)
    
    
    
    def analyze_typing_patterns(self, rsid_metadata: Dict, rsid_timeline: List) -> Tuple[str, Dict]:
        """Advanced analysis of typing patterns to detect manual typing vs copy-paste."""
        if not rsid_timeline:
            return "Not enough data for analysis", {}
            
        # Calculate key metrics
        total_rsids = len(set(rsid_timeline))
        total_segments = sum(meta['segment_count'] for meta in rsid_metadata.values())
        total_words = sum(meta['word_count'] for meta in rsid_metadata.values())
        
        # Compute average words and characters per RSID
        avg_words_per_rsid = total_words / total_rsids if total_rsids > 0 else 0
        
        # Calculate consecutive segments (runs with same RSID)
        consecutive_counts = []
        current_rsid = None
        current_count = 0
        
        for rsid in rsid_timeline:
            if rsid != current_rsid:
                if current_count > 0:
                    consecutive_counts.append(current_count)
                current_rsid = rsid
                current_count = 1
            else:
                current_count += 1
        
        # Add the last count
        if current_count > 0:
            consecutive_counts.append(current_count)
        
        # Calculate statistics on consecutive segments
        avg_consecutive = sum(consecutive_counts) / len(consecutive_counts) if consecutive_counts else 0
        max_consecutive = max(consecutive_counts) if consecutive_counts else 0
        
        # Calculate distribution of segment lengths
        segment_lengths = [meta['word_count'] / meta['segment_count'] if meta['segment_count'] > 0 else 0 
                        for meta in rsid_metadata.values()]
        
        avg_segment_length = sum(segment_lengths) / len(segment_lengths) if segment_lengths else 0
        
        # Find the most frequent RSIDs by word count
        top_rsids = sorted(rsid_metadata.items(), key=lambda x: x[1]['word_count'], reverse=True)[:5]
        top_rsid_percentages = [
            (rsid, meta['word_count'] / total_words * 100 if total_words > 0 else 0)
            for rsid, meta in top_rsids
        ]
        
        # Calculate standard deviation of words per RSID
        words_per_rsid = [meta['word_count'] for meta in rsid_metadata.values()]
        std_dev_words = np.std(words_per_rsid) if words_per_rsid else 0
        
        # Detect large blocks (potential copy-paste)
        large_blocks = []
        for rsid, meta in rsid_metadata.items():
            if meta['word_count'] > 50 and meta['consecutive_count'] > 3:
                large_blocks.append({
                    'rsid': rsid,
                    'word_count': meta['word_count'],
                    'consecutive_paragraphs': meta['consecutive_count']
                })
        
        # Analyze style consistency
        style_variations = []
        for rsid, meta in rsid_metadata.items():
            if 'styles' in meta and len(meta['styles']) > 1:
                style_variations.append({
                    'rsid': rsid,
                    'styles': meta['styles']
                })
        
        # Look for font variations within RSIDs
        font_variations = []
        for rsid, meta in rsid_metadata.items():
            if 'fonts' in meta and len(meta['fonts']) > 1:
                font_variations.append({
                    'rsid': rsid,
                    'fonts': meta['fonts']
                })
        
        # Calculate RSID frequency over time
        rsid_frequency = defaultdict(int)
        for rsid in rsid_timeline:
            rsid_frequency[rsid] += 1
        
        # Calculate entropy of RSID distribution (higher entropy = more randomness = more likely manual typing)
        probabilities = [count / len(rsid_timeline) for count in rsid_frequency.values()]
        entropy = -sum(p * np.log2(p) if p > 0 else 0 for p in probabilities)
        
        # Analysis and interpretation
        indicators = {
            'avg_words_per_rsid': avg_words_per_rsid,
            'max_consecutive_segments': max_consecutive,
            'avg_consecutive_segments': avg_consecutive, 
            'std_dev_words': std_dev_words,
            'entropy': entropy,
            'total_rsids': total_rsids,
            'large_blocks': large_blocks,
            'top_rsid_percentages': top_rsid_percentages,
            'style_variations': style_variations,
            'font_variations': font_variations
        }
        
        # Make a determination based on indicators
        # Higher values indicate more likely copy-paste
        copy_paste_score = 0.0
        
        # Large blocks of text with same RSID are suspicious
        if max_consecutive > 10:
            copy_paste_score += 0.3
        
        # High average words per RSID are suspicious
        if avg_words_per_rsid > 100:
            copy_paste_score += 0.2
        elif avg_words_per_rsid > 50:
            copy_paste_score += 0.1
        
        # High standard deviation indicates inconsistent typing
        if std_dev_words > 100:
            copy_paste_score += 0.15
        
        # Low entropy indicates less randomness, potential copy-paste
        if entropy < 2.0:
            copy_paste_score += 0.15
        
        # Large percentage of content from single RSID
        if top_rsid_percentages and top_rsid_percentages[0][1] > 40:
            copy_paste_score += 0.2
        
        # Style and font variations within same RSID
        if len(style_variations) > 2 or len(font_variations) > 2:
            copy_paste_score += 0.1
        
        # Normalize the score
        copy_paste_score = min(copy_paste_score, 1.0)
        
        # Determine the conclusion
        if copy_paste_score < 0.3:
            conclusion = "Strong indication of manual typing"
        elif copy_paste_score < 0.5:
            conclusion = "Mostly manual typing with some potential copy-paste"
        elif copy_paste_score < 0.7:
            conclusion = "Mixed typing and copy-paste behavior"
        elif copy_paste_score < 0.9:
            conclusion = "Strong indication of copy-paste behavior"
        else:
            conclusion = "Very strong indication of copy-paste behavior"
        
        indicators['copy_paste_score'] = copy_paste_score
        indicators['conclusion'] = conclusion
        
        return conclusion,{
            "manual_typing_score": 1.0 - copy_paste_score,
            "copy_paste_score": copy_paste_score
        }

    def analyze_editing_sessions(self, rsid_timeline: List, metadata: Dict) -> Dict[str, Any]:
        """Analyze the document for editing sessions based on RSID patterns."""
        if not rsid_timeline:
            return {
                'sessions': 0,
                'session_data': [],
                'analysis': "Not enough data for analysis"
            }
        
        # Define what constitutes a session break
        # If the same RSID repeats more than this threshold, it might be a new session
        session_break_threshold = 10
        
        sessions = []
        current_session = {
            'rsids': [rsid_timeline[0]],
            'unique_rsids': {rsid_timeline[0]},
            'start_idx': 0
        }
        
        # Detect session breaks
        for i in range(1, len(rsid_timeline)):
            current_rsid = rsid_timeline[i]
            current_session['rsids'].append(current_rsid)
            current_session['unique_rsids'].add(current_rsid)
            
            # Check for potential session break
            # 1. If we've seen a long run of the same RSID
            consecutive_same = 1
            for j in range(i-1, max(i-session_break_threshold, -1), -1):
                if rsid_timeline[j] == current_rsid:
                    consecutive_same += 1
                else:
                    break
            
            if consecutive_same >= session_break_threshold:
                # Finalize current session
                current_session['end_idx'] = i
                current_session['length'] = len(current_session['rsids'])
                current_session['unique_count'] = len(current_session['unique_rsids'])
                current_session['unique_rsids'] = list(current_session['unique_rsids'])  # Convert set to list for JSON
                sessions.append(current_session)
                
                # Start new session
                current_session = {
                    'rsids': [current_rsid],
                    'unique_rsids': {current_rsid},
                    'start_idx': i
                }
        
        # Add the last session
        if current_session['rsids']:
            current_session['end_idx'] = len(rsid_timeline) - 1
            current_session['length'] = len(current_session['rsids'])
            current_session['unique_count'] = len(current_session['unique_rsids'])
            current_session['unique_rsids'] = list(current_session['unique_rsids'])
            sessions.append(current_session)
        
        # Analyze sessions
        total_edit_time = metadata.get('total_edit_time', 0)
        estimated_time_per_session = total_edit_time / len(sessions) if sessions else 0
        
        session_analysis = {
            'sessions': len(sessions),
            'session_data': sessions,
            'avg_session_length': sum(s['length'] for s in sessions) / len(sessions) if sessions else 0,
            'avg_unique_rsids_per_session': sum(s['unique_count'] for s in sessions) / len(sessions) if sessions else 0,
            'estimated_time_per_session': estimated_time_per_session,
            'analysis': f"Document was edited in approximately {len(sessions)} sessions" if sessions else "Could not determine editing sessions"
        }
        
        return session_analysis

    def analyze_document_completeness(self) -> Dict[str, Any]:
        """Analyze document completeness and estimate completion percentage."""
        if not self.document_root:
            return {
                'is_complete': False,
                'completion_score': 0.0,
                'analysis': "Could not analyze document completeness"
            }
        
        # Check for completeness indicators
        indicators = {
            'has_conclusion': False,
            'has_references': False,
            'has_headers': False,
            'has_title': False,
            'has_body': False,
            'consistent_formatting': False
        }
        
        # Extract all paragraph text for analysis
        paragraphs = []
        for para in self.document_root.findall('.//w:p', self.namespace):
            text = ''.join(t.text for r in para.findall('.//w:r', self.namespace) 
                        for t in r.findall('w:t', self.namespace) if t is not None and t.text)
            paragraphs.append(text)
        
        if not paragraphs:
            return {
                'is_complete': False,
                'completion_score': 0.0,
                'analysis': "Document appears to be empty"
            }
        
        # Check for title (first non-empty paragraph)
        for para in paragraphs:
            if para.strip():
                indicators['has_title'] = True
                break
        
        # Check for body content (multiple paragraphs)
        indicators['has_body'] = len([p for p in paragraphs if p.strip()]) > 3
        
        # Check for conclusion indicators
        conclusion_terms = ['conclusion', 'summary', 'finally', 'in conclusion', 'to conclude']
        for para in paragraphs[-5:]:  # Check last 5 paragraphs
            lower_para = para.lower()
            if any(term in lower_para for term in conclusion_terms):
                indicators['has_conclusion'] = True
                break
        
        # Check for references/bibliography
        reference_terms = ['references', 'bibliography', 'works cited', 'sources']
        for para in paragraphs[-10:]:  # Check last 10 paragraphs
            lower_para = para.lower()
            if any(term in lower_para for term in reference_terms):
                indicators['has_references'] = True
                break
        
        # Check for headers
        if self.styles_root is not None:
            heading_styles = [style.attrib.get('w:styleId') for style in self.styles_root.findall('.//w:style', self.namespace)
                            if style.attrib.get('w:styleId', '').startswith('Heading')]
            
            for para in self.document_root.findall('.//w:p', self.namespace):
                style_elem = para.find('.//w:pStyle', self.namespace)
                if style_elem is not None and style_elem.attrib.get('w:val') in heading_styles:
                    indicators['has_headers'] = True
                    break
        
        # Check for consistent formatting
        font_inconsistencies = self.detect_font_inconsistencies()
        indicators['consistent_formatting'] = not font_inconsistencies['detected']
        
        # Calculate completion score
        indicator_weights = {
            'has_title': 0.1,
            'has_body': 0.3,
            'has_headers': 0.15,
            'has_conclusion': 0.2,
            'has_references': 0.15,
            'consistent_formatting': 0.1
        }
        
        completion_score = sum(value * indicator_weights[key] for key, value in indicators.items())
        
        # Determine if document appears complete
        is_complete = completion_score >= 0.7
        
        # Generate analysis
        missing_elements = [key.replace('has_', '') for key, value in indicators.items() 
                            if key.startswith('has_') and not value]
        
        if missing_elements:
            analysis = f"Document appears to be missing: {', '.join(missing_elements)}"
        else:
            analysis = "Document appears to have all standard structural elements"
        
        if not indicators['consistent_formatting']:
            analysis += ". Formatting inconsistencies detected."
        
        return {
            'is_complete': is_complete,
            'completion_score': completion_score,
            'indicators': indicators,
            'analysis': analysis
        }

    def detect_academic_misconduct(self, rsid_metadata: Dict, typing_analysis: Dict) -> Dict[str, Any]:
        """Provide a comprehensive assessment of potential academic misconduct."""
        # Default response if no data
        if not rsid_metadata:
            return {
                'misconduct_detected': False,
                'confidence': 0.0,
                'indicators': [],
                'analysis': "Not enough data for analysis"
            }
        
        # Initialize indicators
        indicators = []
        confidence_score = 0.0
        
        # 1. Check for copy-paste behavior
        copy_paste_score = typing_analysis.get('copy_paste_score', 0.0)
        if copy_paste_score > 0.7:
            indicators.append({
                'type': 'Copy-Paste',
                'severity': 'High',
                'description': "Document shows strong evidence of copy-paste behavior"
            })
            confidence_score += 0.3
        elif copy_paste_score > 0.5:
            indicators.append({
                'type': 'Copy-Paste',
                'severity': 'Medium',
                'description': "Document shows moderate evidence of copy-paste behavior"
            })
            confidence_score += 0.2
        
        # 2. Check for abnormally large content blocks
        large_blocks = typing_analysis.get('large_blocks', [])
        if len(large_blocks) > 0:
            block_word_count = sum(block['word_count'] for block in large_blocks)
            total_words = sum(meta['word_count'] for meta in rsid_metadata.values())
            large_block_percentage = block_word_count / total_words if total_words > 0 else 0
            
            if large_block_percentage > 0.4:
                indicators.append({
                    'type': 'Large Content Blocks',
                    'severity': 'High',
                    'description': f"Document contains {len(large_blocks)} large blocks covering ~{large_block_percentage:.1%} of content"
                })
                confidence_score += 0.25
            elif large_block_percentage > 0.2:
                indicators.append({
                    'type': 'Large Content Blocks',
                    'severity': 'Medium',
                    'description': f"Document contains {len(large_blocks)} large blocks covering ~{large_block_percentage:.1%} of content"
                })
                confidence_score += 0.15
        
        # 3. Check for font inconsistencies
        font_inconsistencies = self.detect_font_inconsistencies()
        if font_inconsistencies['detected']:
            if font_inconsistencies['severity'] > 0.7:
                indicators.append({
                    'type': 'Font Inconsistencies',
                    'severity': 'High',
                    'description': f"Document shows significant font and formatting inconsistencies"
                })
                confidence_score += 0.25
            elif font_inconsistencies['severity'] > 0.3:
                indicators.append({
                    'type': 'Font Inconsistencies',
                    'severity': 'Medium',
                    'description': f"Document shows some font and formatting inconsistencies"
                })
                confidence_score += 0.15
        
        # 4. Check for style variations
        style_variations = typing_analysis.get('style_variations', [])
        if len(style_variations) > 3:
            indicators.append({
                'type': 'Style Variations',
                'severity': 'Medium',
                'description': f"Document contains {len(style_variations)} instances of style variations within same content blocks"
            })
            confidence_score += 0.15
        
        # 5. Check for text ID variations (potential merged document)
        text_id_variations = []
        for rsid, meta in rsid_metadata.items():
            if 'text_ids' in meta and len(meta['text_ids']) > 1:
                text_id_variations.append({
                    'rsid': rsid,
                    'text_ids': meta['text_ids']
                })
        
        if len(text_id_variations) > 0:
            indicators.append({
                'type': 'Document Merging',
                'severity': 'Medium',
                'description': f"Evidence suggests document may have been merged from multiple sources"
            })
            confidence_score += 0.2
        
        # 6. Analyze metadata for inconsistencies
        metadata = self.parse_metadata()
        doc_history = self.parse_document_history()
        
        if len(doc_history) < 3 and copy_paste_score > 0.5:
            indicators.append({
                'type': 'Limited Edit History',
                'severity': 'Medium',
                'description': "Document has minimal editing history despite complexity, suggesting potential outsourcing"
            })
            confidence_score += 0.15
        
        # Combine indicators to make a determination
        # Cap the confidence score at 1.0
        confidence_score = min(confidence_score, 1.0)
        
        misconduct_detected = confidence_score > 0.5
        
        # Determine conclusion
        if confidence_score > 0.8:
            analysis = "There is very strong evidence of potential academic misconduct."
        elif confidence_score > 0.6:
            analysis = "There is substantial evidence suggesting potential academic misconduct."
        elif confidence_score > 0.4:
            analysis = "There are some concerning indicators of potential academic misconduct, but the evidence is not conclusive."
        elif confidence_score > 0.2:
            analysis = "There are minor indicators that raise some concerns, but most evidence suggests original work."
        else:
            analysis = "The document appears to be predominantly original work with no significant indicators of misconduct."
        
        return {
            'misconduct_detected': misconduct_detected,
            'confidence': confidence_score,
            'indicators': indicators,
            'analysis': analysis
        }

    def generate_visualization_data(self, runs_data, rsid_colors, rsid_timeline, rsid_metadata):
        """Generate data for visualizations of document revision patterns."""
        if not runs_data or not rsid_timeline:
            return {
                'word_distribution': None,
                'rsid_sequence': None,
                'timeline': None
            }
        
        # 1. Word distribution by RSID
        word_counts = {rsid: meta['word_count'] for rsid, meta in rsid_metadata.items()}
        sorted_rsids = sorted(word_counts.items(), key=lambda x: x[1], reverse=True)
        
        word_distribution = {
            'rsids': [rsid for rsid, _ in sorted_rsids],
            'counts': [count for _, count in sorted_rsids],
            'colors': [rsid_colors.get(rsid, '#CCCCCC') for rsid, _ in sorted_rsids]
        }
        
        # 2. RSID sequence for timeline visualization
        rsid_sequence = []
        current_rsid = None
        current_count = 0
        
        for rsid in rsid_timeline:
            if rsid != current_rsid:
                if current_rsid is not None:
                    rsid_sequence.append({
                        'rsid': current_rsid,
                        'count': current_count,
                        'color': rsid_colors.get(current_rsid, '#CCCCCC')
                    })
                current_rsid = rsid
                current_count = 1
            else:
                current_count += 1
        
        # Add the last sequence
        if current_rsid is not None:
            rsid_sequence.append({
                'rsid': current_rsid,
                'count': current_count,
                'color': rsid_colors.get(current_rsid, '#CCCCCC')
            })
        
        # 3. Generate text timeline for visualization
        unique_rsids = set(rsid_timeline)
        rsid_indices = {rsid: i for i, rsid in enumerate(unique_rsids)}
        
        timeline_data = {
            'positions': list(range(len(rsid_timeline))),
            'rsids': rsid_timeline,
            'y_positions': [rsid_indices.get(rsid, 0) for rsid in rsid_timeline],
            'colors': [rsid_colors.get(rsid, '#CCCCCC') for rsid in rsid_timeline]
        }
        
        return {
            'word_distribution': word_distribution,
            'rsid_sequence': rsid_sequence,
            'timeline': timeline_data
        }

    def generate_report_html(self, analysis_results):
        """Generate a comprehensive HTML report of all analysis results."""
        if not analysis_results:
            return "<html><body><h1>Analysis Failed</h1><p>Could not generate report.</p></body></html>"

        # Extract data from analysis results
        metadata = analysis_results.get('metadata', {})
        tracking_status = analysis_results.get('tracking_status', {})
        document_history = analysis_results.get('document_history', [])
        typing_analysis = analysis_results.get('typing_analysis', {})
        misconduct_analysis = analysis_results.get('misconduct_analysis', {})
        font_inconsistencies = analysis_results.get('font_inconsistencies', {})

        # Start building HTML report
        report = [
            '<!DOCTYPE html>',
            '<html>',
            '<head>',
            '<title>Document Analysis Report</title>',
            '<style>',
            'body { font-family: Arial, sans-serif; margin: 20px; line-height: 1.6; }',
            '.container { max-width: 1000px; margin: 0 auto; }',
            '.section { margin-bottom: 30px; border: 1px solid #ddd; padding: 20px; border-radius: 5px; }',
            '.header { background-color: #f8f9fa; padding: 10px; margin-bottom: 15px; border-radius: 5px; }',
            'h1 { color: #333; }',
            'h2 { color: #444; border-bottom: 1px solid #eee; padding-bottom: 10px; }',
            'h3 { color: #555; }',
            'table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }',
            'th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }',
            'th { background-color: #f2f2f2; }',
            'tr:nth-child(even) { background-color: #f9f9f9; }',
            '.high-severity { background-color: #ffcccc; }',
            '.medium-severity { background-color: #ffffcc; }',
            '.low-severity { background-color: #e6f2ff; }',
            '.progress-container { background-color: #f1f1f1; border-radius: 5px; }',
            '.progress-bar { background-color: #4CAF50; height: 24px; border-radius: 5px; text-align: center; line-height: 24px; color: white; }',
            '.progress-bar.warning { background-color: #ff9800; }',
            '.progress-bar.danger { background-color: #f44336; }',
            '</style>',
            '</head>',
            '<body>',
            '<div class="container">',
            f'<h1>Document Analysis Report</h1>',
            f'<p>Report generated on {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>'
        ]

        # Document Metadata Section
        report.extend([
            '<div class="section">',
            '<div class="header"><h2>Document Metadata</h2></div>',
            '<table>',
        ])
        for key, label in [
            ("title", "Title"), ("creator", "Author"), ("last_modified_by", "Last Modified By"),
            ("created", "Created"), ("modified", "Modified"), ("company", "Company"),
            ("application", "Application"), ("revision", "Revision"), ("total_edit_time", "Total Edit Time (mins)")
        ]:
            report.append(f'<tr><th>{label}</th><td>{metadata.get(key, "Unknown")}</td></tr>')
        
        report.append('</table></div>')

        # Revision Tracking Section
        report.extend([
            '<div class="section">',
            '<div class="header"><h2>Revision Tracking Status</h2></div>',
            '<table>'
        ])
        for key, label in [
            ("tracking_enabled", "Tracking Enabled"), ("track_revisions", "Track Revisions"),
            ("track_format_changes", "Track Format Changes"), ("track_moves", "Track Moves"),
            ("rsidRoot", "RSID Root")
        ]:
            value = "Yes" if tracking_status.get(key, False) else "No"
            report.append(f'<tr><th>{label}</th><td>{value}</td></tr>')

        report.append('</table></div>')

        # Academic Integrity Analysis
        report.append('<div class="section"><div class="header"><h2>Academic Integrity Analysis</h2></div>')

        if misconduct_analysis:
            misconduct_detected = misconduct_analysis.get('misconduct_detected', False)
            confidence = misconduct_analysis.get('confidence', 0.0)

            # Determine progress bar color
            bar_class = 'progress-bar'
            if confidence > 0.7:
                bar_class += ' danger'
            elif confidence > 0.4:
                bar_class += ' warning'

            report.extend([
                f'<h3>Summary Assessment</h3>',
                f'<p><strong>{"Potential misconduct detected" if misconduct_detected else "No significant misconduct detected"}</strong></p>',
                f'<p>{misconduct_analysis.get("analysis", "No details available.")}</p>',
                '<div class="progress-container">',
                f'<div class="{bar_class}" style="width:{confidence*100}%;">Confidence: {confidence*100:.1f}%</div>',
                '</div>'
            ])

            # Misconduct Indicators
            indicators = misconduct_analysis.get('indicators', [])
            if indicators:
                report.append('<h3>Detected Indicators</h3><table><tr><th>Type</th><th>Severity</th><th>Description</th></tr>')
                for indicator in indicators:
                    severity_class = 'high-severity' if indicator.get('severity') == 'High' else (
                        'medium-severity' if indicator.get('severity') == 'Medium' else '')
                    report.append(f'<tr class="{severity_class}"><td>{indicator.get("type", "")}</td>'
                                f'<td>{indicator.get("severity", "")}</td>'
                                f'<td>{indicator.get("description", "")}</td></tr>')
                report.append('</table>')

        report.append('</div>')

        # Typing Pattern Analysis
        if typing_analysis:
            report.extend([
                '<div class="section">',
                '<div class="header"><h2>Writing Pattern Analysis</h2></div>',
                '<table>',
                f'<tr><th>Conclusion</th><td>{typing_analysis.get("conclusion", "Unknown")}</td></tr>',
                f'<tr><th>Copy-Paste Score</th><td>{typing_analysis.get("copy_paste_score", 0.0) * 100:.1f}%</td></tr>',
                f'<tr><th>Average Words Per Edit</th><td>{typing_analysis.get("avg_words_per_rsid", 0.0):.1f}</td></tr>',
                f'<tr><th>Maximum Consecutive Segments</th><td>{typing_analysis.get("max_consecutive_segments", 0)}</td></tr>',
                '</table>',
                '</div>'
            ])

        # Font and Formatting Inconsistencies
        if font_inconsistencies and font_inconsistencies.get('detected', False):
            report.extend([
                '<div class="section">',
                '<div class="header"><h2>Font and Formatting Inconsistencies</h2></div>',
                f'<p>Font inconsistencies detected with severity score: {font_inconsistencies.get("severity", 0.0) * 100:.1f}%</p>',
                '<table><tr><th>Category</th><th>Details</th></tr>'
            ])
            for key in ["fonts", "font_sizes", "languages", "unusual_fonts", "unusual_sizes", "unusual_languages"]:
                if key in font_inconsistencies.get("details", {}):
                    report.append(f'<tr><td>{key.replace("_", " ").title()}</td><td>{", ".join(font_inconsistencies["details"][key])}</td></tr>')

            report.append('</table></div>')

        # Closing HTML tags
        report.append('</div></body></html>')

        return ''.join(report)


            

# Set Streamlit Page Configuration
# Set Streamlit Page Configuration
st.set_page_config(page_title="📄 Advanced Document Forensics", layout="wide")
st.title("📄 Advanced Document Forensics Tool")
st.markdown("---")

# File Upload
uploaded_file = st.file_uploader("📂 Upload a .docx file", type=["docx"], help="Max file size: 10MB. Supports .docx format.")

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_file:
        temp_file.write(uploaded_file.getvalue())
        temp_path = temp_file.name
    
    analyzer = WordDocumentAnalyzer(temp_path)
    
    # Sidebar Navigation
    st.sidebar.title("🔍 Analysis Modules")
    option = st.sidebar.radio("Choose Analysis Type", [
        "Document Metadata", "Revision Tracking", "Editing History",
        "RSID Analysis", "Typing Patterns", "Formatting Anomalies",
        "Plagiarism Detection", "Visualizations", "Comprehensive Report"
    ])
    
    if option == "Document Metadata":
        metadata = analyzer.parse_metadata()
        st.subheader("📑 Document Metadata")
        st.json(metadata)
    
    elif option == "Revision Tracking":
        tracking_status = analyzer.check_revision_tracking_status()
        st.subheader("🔍 Revision Tracking Status")
        st.json(tracking_status)
    
    elif option == "Editing History":
        history = analyzer.parse_document_history()
        st.subheader("📜 Document Editing History")
        history_df = pd.DataFrame(history)
        if not history_df.empty:
            st.dataframe(history_df)
        else:
            st.write("No significant revision history found.")
    
    elif option == "RSID Analysis":
        runs_data, rsid_colors, rsid_timeline, rsid_metadata = analyzer.parse_rsid_data()
        st.session_state["runs_data"] = runs_data
        st.session_state["rsid_colors"] = rsid_colors
        st.session_state["rsid_timeline"] = rsid_timeline
        st.session_state["rsid_metadata"] = rsid_metadata
        st.subheader("📊 RSID Analysis")
        st.write(f"Unique RSIDs: {len(rsid_colors)}")

        # 📝 RSID-Based Text Visualization
        st.subheader("📝 RSID-Based Text Visualization")
        if "runs_data" in st.session_state:
            rsid_text_html = "<div style='font-family: Arial, sans-serif; line-height: 1.5;'>"
            for text, rsid, color in st.session_state["runs_data"]:
                rsid_text_html += f"<span style='background-color:{color}; padding:3px; margin:2px; border-radius:3px;' title='RSID: {rsid}'>{text} </span>"
            rsid_text_html += "</div>"
            st.markdown(rsid_text_html, unsafe_allow_html=True)
        else:
            st.warning("⚠️ Perform RSID Analysis first.")

    
    elif option == "Typing Patterns":
        if "rsid_metadata" in st.session_state and "rsid_timeline" in st.session_state:
            analysis_result, confidence_scores = analyzer.analyze_typing_patterns(
                st.session_state["rsid_metadata"], st.session_state["rsid_timeline"]
            )
            st.session_state["confidence_scores"] = confidence_scores
            st.subheader("📝 Typing Pattern Analysis")
            st.write(analysis_result)
        else:
            st.error("⚠️ RSID Analysis must be performed first to enable Typing Patterns Analysis.")
    
    elif option == "Formatting Anomalies":
        formatting_issues = analyzer.detect_font_inconsistencies()
        st.subheader("🎨 Formatting Anomalies")
        st.json(formatting_issues)
    
    elif option == "Plagiarism Detection":
        st.subheader("📋 Plagiarism & Copy-Paste Detection")
        if "confidence_scores" in st.session_state:
            confidence_scores = st.session_state["confidence_scores"]
            manual_score = confidence_scores.get('manual_typing_score', 0.0)
            copy_paste_score = confidence_scores.get('copy_paste_score', 0.0)
            st.write(f"Manual Typing Confidence: {manual_score:.2f}")
            st.write(f"Copy-Paste Confidence: {copy_paste_score:.2f}")
        else:
            st.error("⚠️ Typing Pattern Analysis must be performed first to show Plagiarism Detection results.")
    
    elif option == "Visualizations":
        st.subheader("📊 Additional Visualizations")
        if "runs_data" in st.session_state and "rsid_metadata" in st.session_state:
            visualization_type = st.selectbox(
                "Select a visualization type:",
                ["Word Count per RSID", "RSID Sequences", "RSID Timeline" , "RSID Heatmap"]
            )
            visual_data = analyzer.generate_visualization_data(
                st.session_state["runs_data"],
                st.session_state["rsid_colors"],
                st.session_state["rsid_timeline"],
                st.session_state["rsid_metadata"]
            )
        else:
            st.error("⚠️ RSID Analysis must be performed first to enable visualizations.")
            visual_data = None
        
        if visual_data:
            if visualization_type == "Word Count per RSID" and visual_data.get('word_distribution'):
                fig = px.bar(
                    x=visual_data['word_distribution']['rsids'],
                    y=visual_data['word_distribution']['counts'],
                    labels={'x': 'RSID', 'y': 'Word Count'},
                    title="Word Count per RSID",
                    color=visual_data['word_distribution']['colors']
                )
                st.plotly_chart(fig)
            elif visualization_type == "RSID Sequences" and visual_data.get('rsid_sequence'):
                fig = px.bar(
                    x=[seq['rsid'] for seq in visual_data['rsid_sequence']],
                    y=[seq['count'] for seq in visual_data['rsid_sequence']],
                    labels={'x': 'RSID', 'y': 'Occurrences'},
                    title="RSID Sequences",
                    color=[seq['color'] for seq in visual_data['rsid_sequence']]
                )
                st.plotly_chart(fig)
            elif visualization_type == "RSID Timeline" and visual_data.get('timeline'):
                fig = px.scatter(
                    x=visual_data['timeline']['positions'],
                    y=visual_data['timeline']['y_positions'],
                    color=visual_data['timeline']['colors'],
                    title="RSID Timeline",
                    labels={'x': 'Document Position', 'y': 'RSID'}
                )
                st.plotly_chart(fig)

            elif visualization_type == "RSID Heatmap" and "rsid_metadata" in st.session_state:
                rsid_df = pd.DataFrame(
                    [(rsid, meta["word_count"]) for rsid, meta in st.session_state["rsid_metadata"].items()],
                    columns=["RSID", "Word Count"]
                )
                
                fig = px.imshow([rsid_df["Word Count"].values], 
                                labels=dict(x="RSID", y="", color="Word Count"),
                                x=rsid_df["RSID"],
                                color_continuous_scale="viridis")  # ✅ Use a valid color scale
                
                st.plotly_chart(fig)

            else:
                st.warning("⚠️ Selected visualization type has no data available.")
    
    elif option == "Comprehensive Report":
        st.subheader("📑 Comprehensive Report")
        if "confidence_scores" in st.session_state:
            confidence_scores = st.session_state["confidence_scores"]
            if confidence_scores['copy_paste_score'] > 0.7:
                st.error("⚠️ High likelihood of copy-paste detected!")
            elif confidence_scores['manual_typing_score'] > 0.7:
                st.success("✅ Text appears to be manually typed.")
            else:
                st.warning("⚠️ Unable to determine confidently.")
        else:
            st.error("⚠️ Typing Pattern Analysis must be performed first to generate a Comprehensive Report.")


        # 📥 Downloadable RSID Data
        st.subheader("📥 Downloadable RSID Data")
        if "rsid_metadata" in st.session_state:
            rsid_stats_df = pd.DataFrame(st.session_state["rsid_metadata"].items(), columns=["RSID", "Details"])
            rsid_stats_df["Word Count"] = rsid_stats_df["Details"].apply(lambda x: x["word_count"])
            rsid_stats_df.drop(columns=["Details"], inplace=True)

            csv_data = rsid_stats_df.to_csv(index=False).encode("utf-8")
            st.download_button("📥 Download RSID Statistics (CSV)", csv_data, "rsid_statistics.csv", "text/csv")

            report_html = "<html><head><title>RSID Report</title></head><body>"
            report_html += "<h1>RSID Word Count Statistics</h1>"
            report_html += "<table border='1'><tr><th>RSID</th><th>Word Count</th></tr>"
            for _, row in rsid_stats_df.iterrows():
                report_html += f"<tr><td>{row['RSID']}</td><td>{row['Word Count']}</td></tr>"
            report_html += "</table></body></html>"

            st.download_button("📄 Download Full RSID Report (HTML)", report_html, "rsid_report.html", "text/html")
        else:
            st.warning("⚠️ Perform RSID Analysis first.")

    
    if os.path.exists(temp_path):
        os.remove(temp_path)
