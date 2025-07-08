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

import re
import os
import sys
import glob
from docx import Document
from docx.shared import Inches
import markdown
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

class MarkdownToWordConverter:
    def __init__(self):
        pass
        
    def extract_mermaid_diagrams(self, markdown_text):
        """Extract mermaid diagrams from markdown text"""
        pattern = r'```mermaid\n(.*?)\n```'
        diagrams = re.findall(pattern, markdown_text, re.DOTALL)
        return diagrams
    
    def render_mermaid_to_image(self, mermaid_code, output_path):
        """Render mermaid diagram to image using playwright"""
        html_content = '''
        <!DOCTYPE html>
        <html>
        <head>
            <script src="https://cdn.jsdelivr.net/npm/mermaid/dist/mermaid.min.js"></script>
        </head>
        <body>
            <div class="mermaid">''' + mermaid_code + '''</div>
            <script>
                mermaid.initialize({startOnLoad: true});
            </script>
        </body>
        </html>
        '''
        
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page()
            page.set_content(html_content)
            page.wait_for_selector('.mermaid svg')
            page.locator('.mermaid').screenshot(path=output_path)
            browser.close()
    
    def convert_file(self, md_file, output_file):
        """Convert single markdown file to Word document"""
        with open(md_file, 'r') as f:
            markdown_text = f.read()
        
        doc = Document()
        
        # Extract and render mermaid diagrams
        mermaid_diagrams = self.extract_mermaid_diagrams(markdown_text)
        diagram_images = []
        
        for i, diagram in enumerate(mermaid_diagrams):
            temp_image = "temp_diagram_" + str(i) + ".png"
            self.render_mermaid_to_image(diagram, temp_image)
            diagram_images.append(temp_image)
        
        # Remove mermaid blocks from markdown
        markdown_text = re.sub(r'```mermaid\n.*?\n```', '{{MERMAID_PLACEHOLDER}}', markdown_text, flags=re.DOTALL)
        
        # Convert markdown to HTML
        html = markdown.markdown(markdown_text)
        soup = BeautifulSoup(html, 'html.parser')
        
        # Process HTML elements
        diagram_index = 0
        for element in soup.find_all(['h1', 'h2', 'h3', 'p']):
            if '{{MERMAID_PLACEHOLDER}}' in element.get_text():
                if diagram_index < len(diagram_images):
                    doc.add_picture(diagram_images[diagram_index], width=Inches(6))
                    diagram_index += 1
            elif element.name.startswith('h'):
                level = int(element.name[1])
                doc.add_heading(element.get_text(), level)
            else:
                doc.add_paragraph(element.get_text())
        
        # Save document
        doc.save(output_file)
        
        # Cleanup temp files
        for img in diagram_images:
            if os.path.exists(img):
                os.remove(img)
    
    def convert_folder(self, folder_path):
        """Convert each markdown file in folder to separate Word documents"""
        md_files = sorted(glob.glob(os.path.join(folder_path, '*.md')))
        
        if not md_files:
            print("No markdown files found in " + folder_path)
            return
        
        for md_file in md_files:
            filename = os.path.splitext(os.path.basename(md_file))[0]
            output_file = filename + ".docx"
            self.convert_file(md_file, output_file)
            print("Converted: " + output_file)

def main():
    if len(sys.argv) != 2:
        print("Usage: python converter.py <folder_path>")
        sys.exit(1)
    
    folder_path = sys.argv[1]
    converter = MarkdownToWordConverter()
    converter.convert_folder(folder_path)
    print("Conversion completed for all files")

if __name__ == "__main__":
    main()
