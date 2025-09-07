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
- Paragraphs with text formatting (bold, italic, links, inline code)
- Unordered and ordered lists (with nesting support)
- Code blocks with monospace formatting
- Blockquotes with proper indentation
- Tables with header formatting
- Horizontal rules
- Mermaid diagrams (automatically rendered and extracted)

## Prerequisites

- **Python 3.7+** (Python 3.8+ recommended)
- **Operating System**: macOS, Linux, or Windows
- **Internet Connection**: Required for downloading mermaid.js during diagram rendering
- **GUI Support**: For the desktop interface, ensure your system supports Tkinter (usually included with Python)

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

### GUI Interface (Recommended)

The easiest way to use the converters is through the desktop GUI application:

```bash
# Start the GUI application
python run_ui.py
```

This provides a user-friendly desktop interface with:

#### Features:
- **Tabbed Interface**: Separate tabs for different conversion modes
- **File Selection**: Easy file and folder selection with visual feedback
- **Real-time Logging**: Live progress tracking and detailed status information
- **Batch Processing**: Convert entire folders with mixed file types
- **Cross-platform**: Works on Windows, macOS, and Linux

#### Interface Tabs:

1. **MD to DOCX Tab**:
   - Select individual Markdown files or entire folders
   - Option to combine all files into a single Word document
   - Visual file list showing selected files
   - One-click conversion with progress tracking

2. **MMD to Draw.io Tab**:
   - Select individual Mermaid files or entire folders
   - Option to combine all diagrams into a single Draw.io file
   - Support for all mermaid diagram types
   - Automatic layout and styling

3. **Batch Processing Tab**:
   - Process entire folders containing both .md and .mmd files
   - Analyze folders to see available files before processing
   - Simultaneous conversion of different file types
   - Perfect for documentation projects with mixed content

4. **Processing Log Tab**:
   - Real-time conversion progress and status
   - Detailed error reporting and success confirmations
   - Save logs to file for debugging
   - Clear logs between operations

### Web Interface (Alternative)

For browser-based usage, you can also use the web interface:

```bash
# Start the web application
python run_web_app.py
```

**Note**: The desktop GUI (`run_ui.py`) is generally recommended for local use as it provides better file system integration and doesn't require a web server.

### Command Line Interface

For automation and scripting, you can still use the command line:

#### MD2DOCX Commands
```bash
# Convert each markdown file to separate Word documents
python md2docx.py <folder_path>

# Combine all markdown files into a single Word document
python md2docx.py -c <folder_path>

# Specify custom output directory
python md2docx.py -t /custom/output/path <folder_path>
```

#### MMD2DRAWIO Commands
```bash
# Convert each mermaid file to separate Draw.io files
python mmd2drawio.py <folder_path>

# Combine all mermaid files into a single Draw.io file
python mmd2drawio.py -c <folder_path>

# Specify custom output directory
python mmd2drawio.py -o /custom/output/path <folder_path>
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
├── md2docx.py           # Main Markdown to DOCX conversion script
├── mmd2drawio.py        # Mermaid to Draw.io converter script
├── ui_app.py            # Desktop GUI application (main interface)
├── run_ui.py            # GUI application launcher
├── run_web_app.py       # Web interface launcher (if available)
├── requirements.txt     # Python dependencies
├── README.md           # This documentation
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

## Mermaid to Draw.io Conversion

The package also includes `mmd2drawio.py` for converting extracted mermaid files to Draw.io format.

### Usage

```bash
# Convert all .mmd files to separate .drawio files
python mmd2drawio.py ./docx/

# Combine all .mmd files into a single .drawio file with multiple pages
python mmd2drawio.py -c ./docx/

# Convert single mermaid file
python mmd2drawio.py diagram.mmd

# Specify output directory
python mmd2drawio.py -o ./output/ ./diagrams/
```

### Complete Workflow

1. **Convert Markdown to Word + Extract Mermaid**:
   ```bash
   uv run md2docx.py ./your-docs/
   ```
   This creates:
   - `.docx` files with embedded diagrams
   - `.mmd` files named as `ch<chapter>_<diagram>.mmd`

2. **Convert Mermaid to Draw.io**:
   ```bash
   uv run python mmd2drawio.py ./your-docs/docx/
   ```
   This creates:
   - `.drawio` files for each mermaid diagram
   - OR combined `.drawio` file with `-c` option

### Supported Mermaid Diagram Types & Notations

#### **Flowcharts** (`graph TD`, `graph LR`, `flowchart`)
**Node Shapes:**
- `A[Rectangle]` → Rounded rectangle
- `A(Round)` → Ellipse/Oval
- `A{Diamond}` → Diamond/Decision
- `A[[Subroutine]]` → Rectangle with double border
- `A[(Database)]` → Cylinder
- `A((Circle))` → Perfect circle
- `A>Flag]` → Parallelogram/Flag
- `A{{Hexagon}}` → Hexagon

**Arrow Types:**
- `A --> B` → Solid arrow
- `A -.-> B` → Dotted arrow
- `A ==> B` → Thick arrow
- `A --- B` → Line (no arrow)
- `A -.- B` → Dotted line (no arrow)
- `A === B` → Thick line (no arrow)

**Edge Labels:**
- `A -->|Label| B` → Arrow with label
- `A -.->|Dotted Label| B` → Dotted arrow with label

#### **Sequence Diagrams** (`sequenceDiagram`)
**Arrow Types:**
- `A -> B: Message` → Synchronous message (solid)
- `A ->> B: Message` → Asynchronous message (dashed)
- `A -.-> B: Message` → Dotted message
- `A -->> B: Message` → Return message (dashed)
- `A -x B: Message` → Cross ending

**Participants:**
- `participant A` → Simple participant
- `participant A as Actor` → Participant with alias
- `actor B as User` → Actor type participant

#### **ER Diagrams** (`erDiagram`)
**Relationship Types:**
- `USER ||--|| ORDER : places` → One-to-one
- `USER ||--o{ ORDER : places` → One-to-many
- `USER }o--|| ORDER : places` → Many-to-one
- `USER }o--o{ ORDER : places` → Many-to-many
- `USER ||..|| ORDER : places` → Dotted (identifying)

**Entity Attributes:**
```
USER {
    int id PK
    string name
    string email
}
```

#### **State Diagrams** (`stateDiagram-v2`)
**State Transitions:**
- `[*] --> State1` → Initial state
- `State1 --> State2: Event` → Labeled transition
- `State2 --> [*]` → Final state

## License

This project is open source. Feel free to modify and distribute according to your needs.

