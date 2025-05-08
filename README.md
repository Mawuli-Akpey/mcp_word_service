# MCP Word Service

A Python-based Microsoft Word document server powered by the MCP Python SDK. This service provides programmatic access to create, read, and modify `.docx` files, making it easy to automate document workflows.

## Features

**Tier 1: Essential Abilities**
- **Create a New Word Document:** Generate new `.docx` files with an optional title.
- **Read Document Content:** Extract all text from paragraphs and tables.
- **Add Paragraph:** Insert new paragraphs at the end of a document.
- **Add Heading:** Add headings (level 1, 2, or 3) to structure your document.
- **List Available Word Documents:** Browse all `.docx` files in a directory.

## Planned Features

- Replace text, add tables/images, set styles, create lists, and more (see `word-mcp-server-tool-priorities.md` for roadmap).

## Requirements

- Python 3.8+
- [MCP Python SDK](https://github.com/microsoft/mcp)
- [python-docx](https://python-docx.readthedocs.io/en/latest/)

## Installation

1. Clone this repository:
    ```sh
    git clone https://github.com/YOUR-USERNAME/mcp_word_service.git
    cd mcp_word_service
    ```
2. (Recommended) Create a virtual environment:
    ```sh
    python -m venv .venv
    source .venv/bin/activate  # On Windows: .venv\Scripts\activate
    ```
3. Install dependencies:
    ```sh
    pip install -r requirements.txt
    ```
    Or manually:
    ```sh
    pip install mcp python-docx
    ```

## Usage

1. Start the server in development mode:
    ```sh
    mcp dev main.py
    ```
2. Use the MCP Inspector or compatible client to call the available abilities.

### Example: Create a New Document

```python
# Example using MCP client (pseudo-code)
result = mcp.create_word_document(filename="MyReport.docx", title="Monthly Report")
print(result)
```

### Example: Add a Paragraph

```python
result = mcp.add_paragraph(filename="MyReport.docx", text="This is a new paragraph.")
print(result)
```

## Project Structure

```
mcp_word_service/
│
├── main.py                # MCP server and ability definitions
├── README.md              # This file
├── requirements.txt       # Python dependencies
├── word-mcp-server-tool-priorities.md  # Ability roadmap and priorities
└── ...
```

## Contributing

Contributions are welcome! Please open an issue or submit a pull request.

## License

[MIT License](LICENSE) (or your preferred license)