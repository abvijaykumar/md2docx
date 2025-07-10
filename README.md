# Markdown to Word Converter

A powerful Python tool that converts markdown files containing mermaid diagrams into professional Word documents. Perfect for documentation, technical reports, and presentations that need to be shared in Microsoft Word format.

## Features

### Core Functionality
- **Markdown to Word Conversion**: Seamlessly converts `.md` files to `.docx` format while preserving formatting
- **Mermaid Diagram Rendering**: Automatically detects and renders mermaid diagrams as high-quality PNG images
- **Batch Processing**: Processes entire folders of markdown files in one command
- **Flexible Output**: Creates separate Word documents for each file OR combines all into a single document
- **Smart Formatting**: Maintains heading hierarchy, paragraphs, and text formatting

### Supported Elements
- Headers (H1-H6) with proper Word styling
- Paragraphs with text formatting
- Mermaid diagrams (flowcharts, sequence diagrams, etc.)
- Code blocks (converted to text)
- Lists and other markdown elements

## Prerequisites

- **Python 3.6+** (Python 3.8+ recommended)
- **Operating System**: macOS, Linux, or Windows
- **Internet Connection**: Required for downloading mermaid.js during diagram rendering

## Installation

### Step 1: Clone or Download
```bash
git clone <repository-url>
cd markdown-converter
```

### Step 2: Install Python Dependencies
```bash
pip install -r requirements.txt
```

### Step 3: Install Browser for Diagram Rendering
```bash
playwright install chromium
```

**Note**: The chromium browser is used headlessly to render mermaid diagrams. No GUI interaction required.

## Usage

### Basic Commands
```bash
# Convert each markdown file to separate Word documents
python md2docx.py <folder_path>

# Combine all markdown files into a single Word document
python md2docx.py -c <folder_path>
```

### Real-World Examples

#### Convert Documentation Folder (Separate Files)
```bash
python md2docx.py /Users/john/Documents/project-docs
```

#### Convert Book Chapters (Combined into Single Document)
```bash
python md2docx.py -c "/Users/jane/Books/my-technical-book"
```

#### Convert Current Directory
```bash
python md2docx.py .
python md2docx.py -c .  # Combined version
```

### What Happens During Conversion

1. **File Discovery**: Scans the specified folder for all `.md` files
2. **Mermaid Detection**: Identifies mermaid code blocks in each file
3. **Diagram Rendering**: Uses Playwright + Chromium to render diagrams as images
4. **Document Creation**: Converts markdown content to Word format
5. **Image Embedding**: Inserts rendered diagrams at appropriate locations
6. **File Output**: Saves each `.docx` file with the same name as the source `.md` file
7. **Cleanup**: Removes temporary image files

## File Structure

```
md2docx/
├── md2docx.py          # Main conversion script
├── requirements.txt      # Python dependencies
├── README.md            # This file
```

## Dependencies

| Package | Version | Purpose |
|---------|---------|----------|
| python-docx | 0.8.11 | Word document creation |
| markdown | 3.1.1 | Markdown parsing |
| beautifulsoup4 | 4.9.3 | HTML processing |
| playwright | 1.40.0 | Browser automation for diagram rendering |

## Output Details

### File Naming Convention
**Separate Files Mode (default):**
- Input: `chapter-1.md` → Output: `chapter-1.docx`
- Input: `README.md` → Output: `README.docx`
- Input: `technical-spec.md` → Output: `technical-spec.docx`

**Combined Mode (-c flag):**
- All `.md` files → Output: `combined.docx`

### Document Structure
**Separate Files Mode:**
- **Headings**: Properly formatted with Word's built-in heading styles
- **Body Text**: Clean paragraph formatting
- **Images**: Mermaid diagrams embedded as 6-inch wide images
- **Spacing**: Appropriate line spacing and margins

**Combined Mode:**
- **File Titles**: Each markdown filename becomes an H1 heading
- **Content Headings**: Original headings offset by one level (H1→H2, H2→H3, etc.)
- **Page Breaks**: Automatic separation between different source files
- **Images**: Mermaid diagrams with unique naming to avoid conflicts

### Mermaid Diagram Support
Supported mermaid diagram types:
- Flowcharts
- Sequence diagrams
- Class diagrams
- State diagrams
- Gantt charts
- And more!

## Troubleshooting

### Common Issues

**"No module named 'docx'"**
```bash
pip install python-docx
```

**"No markdown files found"**
- Ensure the folder path is correct
- Check that files have `.md` extension
- Use absolute paths if relative paths don't work

**Mermaid diagrams not rendering**
- Ensure internet connection is available
- Check that chromium is installed: `playwright install chromium`
- Verify mermaid syntax is correct

**Permission errors**
- Ensure write permissions in the output directory
- Run with appropriate user permissions

## Tips

- **Large Folders**: The tool processes files sequentially, so large folders may take time
- **File Organization**: Output files are created in the current working directory
- **Backup**: Always backup original markdown files before conversion
- **Testing**: Use the included `example.md` to test the setup

## Example Workflow

1. **Prepare**: Organize your markdown files in a folder
2. **Install**: Run the installation commands
3. **Convert**: Execute the converter command
4. **Review**: Check the generated Word documents
5. **Share**: Distribute the professional `.docx` files

## License

This project is open source. Feel free to modify and distribute according to your needs.

