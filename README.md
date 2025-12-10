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

- **UV Package Manager**: This project uses UV for dependency management and script execution
  - Install UV: https://docs.astral.sh/uv/getting-started/installation/
- **Python 3.12**: Automatically managed by UV (specified in [`.python-version`](.python-version))
- **Operating System**: macOS, Linux, or Windows
- **Internet Connection**: Required for downloading mermaid.js during diagram rendering
- **GUI Support**: For the desktop interface, ensure your system supports Tkinter (usually included with Python)

## Installation

> **Migration Note**: This project has been migrated from pip/virtualenv to UV for improved dependency management and development workflows. If you previously used `pip install -r requirements.txt`, please follow the new UV-based instructions below.

### Step 1: Install UV

If you don't have UV installed, follow the installation guide at https://docs.astral.sh/uv/getting-started/installation/

Quick installation:
```bash
# macOS/Linux
curl -LsSf https://astral.sh/uv/install.sh | sh

# Windows
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"
```

### Step 2: Clone or Download
```bash
git clone <repository-url>
cd md2docx
```

### Step 3: Install Dependencies
```bash
# Install all project dependencies
uv sync

# Or install with development dependencies
uv sync --extra dev
```

UV will automatically:
- Install Python 3.12 if not present
- Create a virtual environment
- Install all required dependencies from [`pyproject.toml`](pyproject.toml)

### Step 4: Install Browser for Diagram Rendering
```bash
uv run playwright install chromium
```

**Note**: The chromium browser is used headlessly to render mermaid diagrams. No GUI interaction required.

## Usage

### GUI Interface (Recommended)

The easiest way to use the converters is through the desktop GUI application:

```bash
# Start the GUI application
uv run md2docx-ui
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

**Note**: The desktop GUI is recommended for local use as it provides better file system integration.

### Command Line Interface

For automation and scripting, use the command line with UV:

#### MD2DOCX Commands
```bash
# Convert each markdown file to separate Word documents
uv run md2docx <folder_path>

# Combine all markdown files into a single Word document
uv run md2docx -c <folder_path>

# Specify custom output directory
uv run md2docx -t /custom/output/path <folder_path>
```

#### DOCX2MD Commands
```bash
# Convert Word documents back to Markdown
uv run docx2md <folder_path>
```

#### MMD2DRAWIO Commands
```bash
# Convert each mermaid file to separate Draw.io files
uv run mmd2drawio <folder_path>

# Combine all mermaid files into a single Draw.io file
uv run mmd2drawio -c <folder_path>

# Specify custom output directory
uv run mmd2drawio -o /custom/output/path <folder_path>
```

### Real-World Examples

#### Convert Documentation Folder (Separate Files)
```bash
uv run md2docx /Users/john/Documents/project-docs
```

#### Convert Book Chapters (Combined into Single Document)
```bash
uv run md2docx -c "/Users/jane/Books/my-technical-book"
```

#### Convert Current Directory
```bash
uv run md2docx .
uv run md2docx -c .  # Combined version
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
├── docx2md.py           # DOCX to Markdown conversion script
├── mmd2drawio.py        # Mermaid to Draw.io converter script
├── ui_app.py            # Desktop GUI application (main interface)
├── run_ui.py            # GUI application launcher
├── pyproject.toml       # Project configuration and dependencies
├── uv.lock              # UV lock file for reproducible builds
├── .python-version      # Python version specification (3.12)
├── requirements.txt     # Legacy requirements (for reference)
├── README.md            # This documentation
```

## Dependencies

All dependencies are managed through UV and defined in [`pyproject.toml`](pyproject.toml):

| Package | Purpose |
|---------|----------|
| python-docx | Word document creation and reading |
| markdown | Markdown parsing |
| beautifulsoup4 | HTML processing |
| playwright | Browser automation for diagram rendering |
| lxml | XML processing |

**Development Dependencies:**
| Package | Purpose |
|---------|----------|
| black | Code formatting |
| ruff | Fast Python linter |
| pytest | Testing framework (if applicable) |

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

## Development

This project uses UV for dependency management and development workflows.

### Adding Dependencies

```bash
# Add a new runtime dependency
uv add <package-name>

# Add a development dependency
uv add --dev <package-name>
```

### Updating Dependencies

```bash
# Update all dependencies to latest compatible versions
uv sync

# Update UV lock file
uv lock --upgrade
```

### Code Quality

```bash
# Format code with Black
uv run black .

# Lint code with Ruff
uv run ruff check .

# Auto-fix linting issues
uv run ruff check --fix .
```

### Running Tests

```bash
# Run tests (if test suite exists)
uv run pytest

# Run tests with coverage
uv run pytest --cov
```

## Troubleshooting

### Common Issues

**"Command not found: uv"**
- Install UV following the instructions at https://docs.astral.sh/uv/getting-started/installation/

**"No module named 'docx'"**
```bash
# Ensure dependencies are installed
uv sync
```

**"No markdown files found"**
- Ensure the folder path is correct
- Check that files have `.md` extension
- Use absolute paths if relative paths don't work

**Mermaid diagrams not rendering**
- Ensure internet connection is available
- Check that chromium is installed: `uv run playwright install chromium`
- Verify mermaid syntax is correct

**Permission errors**
- Ensure write permissions in the output directory
- Run with appropriate user permissions

**Python version mismatch**
- UV automatically manages Python 3.12 as specified in [`.python-version`](.python-version)
- Run `uv python list` to see available Python versions

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
   uv run md2docx ./your-docs/
   ```
   This creates:
   - `.docx` files with embedded diagrams
   - `.mmd` files named as `ch<chapter>_<diagram>.mmd`

2. **Convert Mermaid to Draw.io**:
   ```bash
   uv run mmd2drawio ./your-docs/docx/
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

