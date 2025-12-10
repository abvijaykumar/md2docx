
# -*- coding: utf-8 -*-
# MIT License
#
# Copyright (c) 2025 A B Vijay Kumar
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

import os
import sys
import argparse
import re
import zipfile
import io
from pathlib import Path
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.document import Document as DocumentType
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
try:
    from PIL import Image
except ImportError:
    Image = None

class EnhancedDocxToMarkdownConverter:
    """Enhanced DOCX to Markdown converter with advanced features"""
    
    def __init__(self, config=None):
        self.log_callback = None
        self.config = config or {}
        
        # Default configuration
        self.default_config = {
            'extract_images': True,
            'image_folder': 'assets',
            'image_format': 'png',
            'preserve_formatting': True,
            'table_alignment': True,
            'extract_hyperlinks': True,
            'extract_footnotes': True,
            'include_metadata': True,
            'list_style_preservation': True,
            'heading_detection_method': 'advanced',  # 'basic' or 'advanced'
            'output_format': 'github',  # 'github', 'commonmark', 'basic'
        }
        
        # Merge user config with defaults
        for key, value in self.default_config.items():
            if key not in self.config:
                self.config[key] = value
        
        # Initialize tracking variables
        self.image_counter = 0
        self.footnote_counter = 0
        self.extracted_images = []
    
    def log(self, message):
        """Log message using callback or print"""
        if self.log_callback:
            self.log_callback(message)
        else:
            print(message)
    
    def extract_images_from_docx(self, docx_path, output_dir):
        """Extract all images from DOCX file"""
        if not self.config['extract_images']:
            return {}
        
        images = {}
        image_folder = os.path.join(output_dir, self.config['image_folder'])
        os.makedirs(image_folder, exist_ok=True)
        
        try:
            with zipfile.ZipFile(docx_path, 'r') as docx_zip:
                # Get list of image files in the DOCX
                image_files = [f for f in docx_zip.namelist() if f.startswith('word/media/')]
                
                for image_file in image_files:
                    try:
                        # Extract image data
                        image_data = docx_zip.read(image_file)
                        
                        # Determine file extension
                        original_ext = os.path.splitext(image_file)[1].lower()
                        if not original_ext:
                            original_ext = '.png'  # Default
                        
                        # Generate new filename
                        self.image_counter += 1
                        base_name = os.path.splitext(os.path.basename(docx_path))[0]
                        new_filename = f"{base_name}_image_{self.image_counter:03d}{original_ext}"
                        image_path = os.path.join(image_folder, new_filename)
                        
                        # Save image
                        with open(image_path, 'wb') as img_file:
                            img_file.write(image_data)
                        
                        # Store mapping
                        images[os.path.basename(image_file)] = {
                            'path': image_path,
                            'relative_path': os.path.join(self.config['image_folder'], new_filename),
                            'filename': new_filename
                        }
                        
                        self.log(f"  Extracted image: {new_filename}")
                        
                    except Exception as e:
                        self.log(f"  Warning: Could not extract image {image_file}: {str(e)}")
                        
        except Exception as e:
            self.log(f"  Warning: Could not extract images: {str(e)}")
        
        return images
    
    def get_paragraph_style_info(self, paragraph):
        """Get detailed style information from paragraph"""
        style_info = {
            'is_heading': False,
            'heading_level': 0,
            'is_list': False,
            'list_level': 0,
            'alignment': 'left',
            'is_quote': False
        }
        
        # Check if it's a heading using multiple methods
        if self.config['heading_detection_method'] == 'advanced':
            # Method 1: Style name
            if paragraph.style.name.startswith('Heading'):
                style_info['is_heading'] = True
                try:
                    level_str = paragraph.style.name.split()[-1]
                    if level_str.isdigit():
                        style_info['heading_level'] = int(level_str)
                    else:
                        style_info['heading_level'] = 1
                except:
                    style_info['heading_level'] = 1
            
            # Method 2: Outline level
            elif hasattr(paragraph._element, 'pPr') and paragraph._element.pPr is not None:
                outline_lvl = paragraph._element.pPr.find(qn('w:outlineLvl'))
                if outline_lvl is not None:
                    style_info['is_heading'] = True
                    style_info['heading_level'] = int(outline_lvl.get(qn('w:val'))) + 1
        else:
            # Basic method - only style name
            if paragraph.style.name.startswith('Heading'):
                style_info['is_heading'] = True
                try:
                    level_str = paragraph.style.name.split()[-1]
                    style_info['heading_level'] = int(level_str) if level_str.isdigit() else 1
                except:
                    style_info['heading_level'] = 1
        
        # Check for list styles
        if 'List' in paragraph.style.name or 'Bullet' in paragraph.style.name:
            style_info['is_list'] = True
            # Try to determine list level from indentation
            if hasattr(paragraph._element, 'pPr') and paragraph._element.pPr is not None:
                ind = paragraph._element.pPr.find(qn('w:ind'))
                if ind is not None:
                    left_indent = ind.get(qn('w:left'))
                    if left_indent:
                        # Estimate level based on indentation (720 twips = 0.5 inch)
                        style_info['list_level'] = int(int(left_indent) / 720)
        
        # Check alignment
        if paragraph.alignment:
            if paragraph.alignment == WD_PARAGRAPH_ALIGNMENT.CENTER:
                style_info['alignment'] = 'center'
            elif paragraph.alignment == WD_PARAGRAPH_ALIGNMENT.RIGHT:
                style_info['alignment'] = 'right'
            elif paragraph.alignment == WD_PARAGRAPH_ALIGNMENT.JUSTIFY:
                style_info['alignment'] = 'justify'
        
        # Check for quote style
        if 'Quote' in paragraph.style.name or 'Block' in paragraph.style.name:
            style_info['is_quote'] = True
        
        return style_info
    
    def process_run_formatting(self, run):
        """Process formatting for a text run and return markdown"""
        text = run.text
        if not text:
            return ""
        
        # Handle different formatting combinations
        if self.config['preserve_formatting']:
            # Bold and italic
            if run.bold and run.italic:
                text = f"***{text}***"
            elif run.bold:
                text = f"**{text}**"
            elif run.italic:
                text = f"*{text}*"
            
            # Underline (HTML tags for now, as markdown doesn't support it natively)
            if run.underline:
                text = f"<u>{text}</u>"
            
            # Strikethrough
            if hasattr(run.font, 'strike') and run.font.strike:
                text = f"~~{text}~~"
            
            # Code/monospace
            if run.font.name in ['Courier New', 'Consolas', 'Monaco', 'Menlo']:
                text = f"`{text}`"
            
            # Superscript/Subscript (HTML tags)
            if hasattr(run.font, 'superscript') and run.font.superscript:
                text = f"<sup>{text}</sup>"
            elif hasattr(run.font, 'subscript') and run.font.subscript:
                text = f"<sub>{text}</sub>"
        
        return text
    
    def extract_hyperlinks(self, paragraph):
        """Extract hyperlinks from paragraph"""
        hyperlinks = []
        if not self.config['extract_hyperlinks']:
            return [], paragraph.text
        
        # This is a simplified version - full implementation would parse XML
        # For now, we'll return the text as-is and note hyperlinks in comments
        try:
            # Look for hyperlink elements in the paragraph XML
            xml = paragraph._element.xml
            if 'hyperlink' in xml.lower():
                # Placeholder for hyperlink extraction
                # Full implementation would parse the XML properly
                pass
        except:
            pass
        
        return hyperlinks, paragraph.text
    
    def convert_table_to_markdown(self, table):
        """Convert Word table to enhanced markdown table"""
        if not table.rows:
            return []
        
        markdown_lines = []
        rows = []
        
        # Process all rows
        for row in table.rows:
            cells = []
            for cell in row.cells:
                # Get cell text and clean it
                cell_text = cell.text.strip().replace('\n', ' ').replace('\r', '')
                
                # Handle empty cells
                if not cell_text:
                    cell_text = " "
                
                # Escape pipe characters in cell content
                cell_text = cell_text.replace('|', '\\|')
                
                cells.append(cell_text)
            rows.append(cells)
        
        if not rows:
            return []
        
        # Determine number of columns
        max_cols = max(len(row) for row in rows) if rows else 0
        if max_cols == 0:
            return []
        
        # Normalize all rows to have the same number of columns
        for row in rows:
            while len(row) < max_cols:
                row.append(" ")
        
        # Create header row (first row)
        header = "| " + " | ".join(rows[0]) + " |"
        markdown_lines.append(header)
        
        # Create separator row with alignment if configured
        if self.config['table_alignment']:
            # For now, default to left alignment
            # Future enhancement: detect cell alignment from Word document
            separator = "| " + " | ".join(["---"] * max_cols) + " |"
        else:
            separator = "| " + " | ".join(["---"] * max_cols) + " |"
        
        markdown_lines.append(separator)
        
        # Add data rows
        for row_data in rows[1:]:
            row_line = "| " + " | ".join(row_data) + " |"
            markdown_lines.append(row_line)
        
        markdown_lines.append("")  # Empty line after table
        return markdown_lines
    
    def process_list_paragraph(self, paragraph, style_info):
        """Process list paragraph with proper nesting"""
        level = style_info['list_level']
        indent = "  " * level  # 2 spaces per level
        
        # Determine bullet type
        if 'Bullet' in paragraph.style.name or paragraph.style.name.startswith('List Bullet'):
            if level == 0:
                bullet = "- "
            elif level == 1:
                bullet = "  - "
            else:
                bullet = "    - "
        else:
            # Numbered list
            bullet = "1. " if level == 0 else f"{indent}1. "
        
        # Process the text with formatting
        formatted_text = ""
        for run in paragraph.runs:
            formatted_text += self.process_run_formatting(run)
        
        return bullet + formatted_text
    
    def extract_document_metadata(self, doc):
        """Extract document metadata"""
        if not self.config['include_metadata']:
            return {}
        
        metadata = {}
        try:
            core_props = doc.core_properties
            metadata['title'] = core_props.title or ""
            metadata['author'] = core_props.author or ""
            metadata['subject'] = core_props.subject or ""
            metadata['created'] = core_props.created
            metadata['modified'] = core_props.modified
            metadata['category'] = core_props.category or ""
            metadata['comments'] = core_props.comments or ""
        except Exception as e:
            self.log(f"  Warning: Could not extract meta {str(e)}")
        
        return metadata
    
    def convert_file(self, docx_path, md_path):
        """Convert single DOCX file to enhanced Markdown"""
        try:
            self.log(f"Converting: {os.path.basename(docx_path)}")
            
            # Load document
            doc = Document(docx_path)
            
            # Prepare output directory
            output_dir = os.path.dirname(md_path) if os.path.dirname(md_path) else '.'
            os.makedirs(output_dir, exist_ok=True)
            
            # Extract images
            images = self.extract_images_from_docx(docx_path, output_dir)
            if images:
                self.log(f"  Extracted {len(images)} images")
            
            # Extract metadata
            metadata = self.extract_document_metadata(doc)
            
            # Start building markdown content
            markdown_content = []
            
            # Add metadata header if enabled
            if metadata and self.config['include_metadata']:
                markdown_content.append("---")
                for key, value in metadata.items():
                    if value:
                        markdown_content.append(f"{key}: {value}")
                markdown_content.append("---")
                markdown_content.append("")
            
            # Process document content
            for element in doc.element.body:
                if isinstance(element, CT_P):
                    # Paragraph
                    paragraph = Paragraph(element, doc)
                    self._process_paragraph(paragraph, markdown_content, images)
                elif isinstance(element, CT_Tbl):
                    # Table
                    table = Table(element, doc)
                    table_md = self.convert_table_to_markdown(table)
                    markdown_content.extend(table_md)
            
            # Write markdown file
            with open(md_path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(markdown_content))
            
            self.log(f"✓ Converted: {os.path.basename(docx_path)} -> {os.path.basename(md_path)}")
            
        except Exception as e:
            self.log(f"✗ Error converting {docx_path}: {str(e)}")
            raise
    
    def _process_paragraph(self, paragraph, markdown_content, images):
        """Process a paragraph and add to markdown content"""
        if not paragraph.text.strip():
            markdown_content.append("")
            return
        
        # Get style information
        style_info = self.get_paragraph_style_info(paragraph)
        
        # Handle different paragraph types
        if style_info['is_heading']:
            # Heading
            level = min(style_info['heading_level'], 6)  # Max 6 levels in markdown
            heading_text = ""
            for run in paragraph.runs:
                heading_text += self.process_run_formatting(run)
            markdown_content.append(f"{'#' * level} {heading_text}")
            
        elif style_info['is_list']:
            # List item
            list_line = self.process_list_paragraph(paragraph, style_info)
            markdown_content.append(list_line)
            
        elif style_info['is_quote']:
            # Quote/blockquote
            quote_text = ""
            for run in paragraph.runs:
                quote_text += self.process_run_formatting(run)
            markdown_content.append(f"> {quote_text}")
            
        else:
            # Regular paragraph
            paragraph_text = ""
            for run in paragraph.runs:
                paragraph_text += self.process_run_formatting(run)
            
            # Handle alignment if configured
            if self.config['preserve_formatting'] and style_info['alignment'] != 'left':
                if style_info['alignment'] == 'center':
                    paragraph_text = f"<div align='center'>{paragraph_text}</div>"
                elif style_info['alignment'] == 'right':
                    paragraph_text = f"<div align='right'>{paragraph_text}</div>"
            
            markdown_content.append(paragraph_text)
    
    def convert_multiple_files(self, docx_files, output_dir):
        """Convert multiple DOCX files to separate Markdown files"""
        os.makedirs(output_dir, exist_ok=True)
        
        for docx_file in docx_files:
            filename = os.path.splitext(os.path.basename(docx_file))[0]
            md_path = os.path.join(output_dir, f"{filename}.md")
            self.convert_file(docx_file, md_path)
    
    def convert_combined(self, docx_files, output_path):
        """Convert multiple DOCX files into a single combined Markdown file"""
        combined_content = []
        output_dir = os.path.dirname(output_path) if os.path.dirname(output_path) else '.'
        
        for i, docx_file in enumerate(docx_files):
            if i > 0:
                combined_content.append("\n---\n")  # Separator between files
            
            # Add file title
            filename = os.path.splitext(os.path.basename(docx_file))[0]
            combined_content.append(f"# {filename}\n")
            
            try:
                # Load document
                doc = Document(docx_file)
                
                # Extract images for this file
                images = self.extract_images_from_docx(docx_file, output_dir)
                
                # Extract metadata
                metadata = self.extract_document_metadata(doc)
                
                # Add metadata if enabled
                if metadata and self.config['include_metadata']:
                    combined_content.append("## Document Information")
                    for key, value in metadata.items():
                        if value:
                            combined_content.append(f"- **{key.title()}**: {value}")
                    combined_content.append("")
                
                # Process content
                for element in doc.element.body:
                    if isinstance(element, CT_P):
                        paragraph = Paragraph(element, doc)
                        self._process_paragraph_combined(paragraph, combined_content, images)
                    elif isinstance(element, CT_Tbl):
                        table = Table(element, doc)
                        table_md = self.convert_table_to_markdown(table)
                        combined_content.extend(table_md)
                
                self.log(f"✓ Processed: {os.path.basename(docx_file)}")
                
            except Exception as e:
                self.log(f"✗ Error processing {docx_file}: {str(e)}")
        
        # Write combined file
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(combined_content))
        
        self.log(f"✓ Created combined file: {output_path}")
    
    def _process_paragraph_combined(self, paragraph, combined_content, images):
        """Process paragraph for combined file (adjusts heading levels)"""
        if not paragraph.text.strip():
            combined_content.append("")
            return
        
        style_info = self.get_paragraph_style_info(paragraph)
        
        if style_info['is_heading']:
            # Adjust heading level for combined document (shift down by 1)
            level = min(style_info['heading_level'] + 1, 6)
            heading_text = ""
            for run in paragraph.runs:
                heading_text += self.process_run_formatting(run)
            combined_content.append(f"{'#' * level} {heading_text}")
            
        elif style_info['is_list']:
            list_line = self.process_list_paragraph(paragraph, style_info)
            combined_content.append(list_line)
            
        elif style_info['is_quote']:
            quote_text = ""
            for run in paragraph.runs:
                quote_text += self.process_run_formatting(run)
            combined_content.append(f"> {quote_text}")
            
        else:
            paragraph_text = ""
            for run in paragraph.runs:
                paragraph_text += self.process_run_formatting(run)
            
            if self.config['preserve_formatting'] and style_info['alignment'] != 'left':
                if style_info['alignment'] == 'center':
                    paragraph_text = f"<div align='center'>{paragraph_text}</div>"
                elif style_info['alignment'] == 'right':
                    paragraph_text = f"<div align='right'>{paragraph_text}</div>"
            
            combined_content.append(paragraph_text)


# Legacy class for backward compatibility
class DocxToMarkdownConverter(EnhancedDocxToMarkdownConverter):
    """Legacy class name for backward compatibility"""
    
    def __init__(self):
        # Initialize with basic configuration for backward compatibility
        super().__init__(config={
            'extract_images': False,
            'preserve_formatting': True,
            'table_alignment': False,
            'extract_hyperlinks': False,
            'extract_footnotes': False,
            'include_metadata': False,
            'list_style_preservation': False,
            'heading_detection_method': 'basic',
            'output_format': 'basic'
        })


def main():
    """Command line interface"""
    parser = argparse.ArgumentParser(description='Convert DOCX files to Markdown with enhanced features')
    parser.add_argument('input_path', help='Input DOCX file or folder')
    parser.add_argument('-o', '--output', help='Output directory (default: current directory)')
    parser.add_argument('-c', '--combine', action='store_true', help='Combine all files into single markdown')
    parser.add_argument('--extract-images', action='store_true', help='Extract images from DOCX files')
    parser.add_argument('--image-folder', default='assets', help='Folder name for extracted images')
    parser.add_argument('--preserve-formatting', action='store_true', default=True, help='Preserve text formatting')
    parser.add_argument('--table-alignment', action='store_true', help='Preserve table alignment')
    parser.add_argument('--extract-hyperlinks', action='store_true', help='Extract hyperlinks')
    parser.add_argument('--include-metadata', action='store_true', help='Include document metadata')
    
    args = parser.parse_args()
    
    # Build configuration
    config = {
        'extract_images': args.extract_images,
        'image_folder': args.image_folder,
        'preserve_formatting': args.preserve_formatting,
        'table_alignment': args.table_alignment,
        'extract_hyperlinks': args.extract_hyperlinks,
        'include_metadata': args.include_metadata,
        'heading_detection_method': 'advanced',
        'output_format': 'github'
    }
    
    converter = EnhancedDocxToMarkdownConverter(config)
    
    input_path = Path(args.input_path)
    output_dir = Path(args.output) if args.output else Path.cwd()
    
    if input_path.is_file() and input_path.suffix.lower() == '.docx':
        # Single file
        output_file = output_dir / f"{input_path.stem}.md"
        converter.convert_file(str(input_path), str(output_file))
    elif input_path.is_dir():
        # Directory
        docx_files = list(input_path.glob('*.docx'))
        if not docx_files:
            print("No DOCX files found in the specified directory")
            return
        
        if args.combine:
            output_file = output_dir / "combined.md"
            converter.convert_combined([str(f) for f in docx_files], str(output_file))
        else:
            converter.convert_multiple_files([str(f) for f in docx_files], str(output_dir))
    else:
        print("Invalid input path or not a DOCX file")
        return

if __name__ == "__main__":
    main()