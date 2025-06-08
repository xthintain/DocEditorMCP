# Word Editor MCP Server

A Model Context Protocol (MCP) server for creating and editing Word documents programmatically.

## Features

- Create new Word documents and text files
- Read document content
- Format text (fonts, colors, styles)
- Adjust paragraph spacing
- Insert images and tables
- Edit table cells
- Convert to PDF/HTML/TXT formats
- Add headers/footers and page numbers
- Set page layout (margins, orientation)
- Merge multiple documents
- Batch process document structure

## Requirements

- Python 3.7+
- Required packages: `python-docx`
- Optional packages:
  - `Pillow` for image support
  - `pywin32` for advanced features on Windows

## Installation

1. Install requirements:
   ```bash
   pip install -r requirements.txt
   ```

2. (Optional) For Windows users needing advanced features:
   ```bash
   pip install pywin32
   ```

## Usage

Run the server:
```bash
python word_server.py
```

The server will be available to any MCP client with the following capabilities:

### Basic Operations
- Create empty TXT files
- Create new Word documents
- Read document content

### Formatting
- Text formatting (font, size, color, bold/italic/underline)
- Paragraph spacing (before/after, line spacing)
- Page layout (margins, orientation)

### Document Elements
- Insert images
- Insert and edit tables
- Add headers/footers
- Insert table of contents

### Document Conversion
- Save as PDF, DOCX, DOC, HTML, TXT
- Merge multiple documents

### Batch Processing
- Process document structure with JSON configuration

## Configuration

Set the `OFFICE_EDIT_PATH` environment variable to specify where documents should be saved (defaults to Desktop).

use for MCP:
{
    "mcpServers":{
        "wordEditor": {
        "command": "python",
        "args": ["path to word_server.py"]
        }
    }
}


## Notes

- Some advanced features require Windows and pywin32
- The server will warn if optional dependencies are missing
- Document paths can be absolute or relative to `OFFICE_EDIT_PATH`

## Note: This project is a secondary modification project.
source from https://github.com/theWDY/office-editor-mcp