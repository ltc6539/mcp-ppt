# PPT maker MCP server

🌐 [中文版README](README_zh.md)

This MCP server enables dynamic creation, editing and saving of PowerPoint presentations. Built upon the [MCP](https://github.com/modelcontextprotocol/python-sdk) and using the [python-pptx](https://python-pptx.readthedocs.io/en/latest/) library, it provides a flexible interface to add slides, images, tables, and other elements. Users could effortlessly make, edit and save presentations by simply chatting with a large language model, streamlining the entire workflow

## Features

- **Create Presentations**  
  Initialize a new PowerPoint presentation using a title that generates a unique presentation ID.

- **Slide Operations**  
  - **Title Slide:** Add a title slide with an optional subtitle.
  - **Content Slide:** Create slides with a title and bullet-point content.
  - **Section Slide:** Insert a section divider slide with a large centered title and an optional background color.
  - **Image Slide:** Add slides featuring images from local files or URLs with titles and descriptive alt text.
  - **Table Slide:** Insert slides containing tables with defined headers and row data.

- **Presentation Management**  
  - **Save Presentation:** Write the presentation to a specified file path, handling temporary directories if needed.
  - **Download Link:** Generate a data URI with base64-encoded presentation content for direct download.
  - **Presentation Info:** Retrieve metadata about the presentation such as the number of slides and available slide layouts.
  - **Presentation Outline:** Obtain a text-based outline of the presentation structure via a dedicated resource endpoint.
  - **Remove Slide:** Delete a slide by its 1-based index.
  - **Export to Base64:** Export the complete presentation as a base64-encoded string for further processing.

## Installation

1. **Clone the Repository**  
   ```bash
   git clone https://github.com/ltc6539/mcp-ppt.git
   cd mcp-ppt
   ```

2. **Create a Virtual Environment (Optional but Recommended)**
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate  # On Windows use: .venv\Scripts\activate
   ```

3. **Then add MCP to your project dependencies**
   ```bash
   uv add "mcp[cli]"
   uv run mcp
   ```

You can install this server in [Claude Desktop](https://claude.ai/download) and interact with it right away by running:
```bash
mcp install server-local.py
```

Alternatively, you can test it with the MCP Inspector:
```bash
mcp dev server-local.py
```

If the claude desktop has error, You may need to put the full path to the uv executable in the command field. You can get this by running `which uv` on MacOS/Linux or `where uv` on Windows.
During startup, the server logs Python and python-pptx version information to stderr. Any errors during execution are also printed to stderr for easy debugging.

## Tool List

Each MCP tool function is directly accessible via the MCP server. Below are the primary operations available:

### 1. Create Presentation
- **Function:** `create_presentation(title: str) -> str`  
- **Description:** Initializes a new presentation and returns a unique presentation ID.
  
### 2. Add Title Slide
- **Function:** `add_title_slide(prs_id: str, title: str, subtitle: Optional[str] = None) -> str`  
- **Description:** Adds a title slide to the specified presentation.

### 3. Add Content Slide
- **Function:** `add_content_slide(prs_id: str, title: str, content: List[str]) -> str`  
- **Description:** Inserts a content slide with a title and multiple bullet points.

### 4. Add Section Slide
- **Function:** `add_section_slide(prs_id: str, section_title: str, background_color: Optional[str] = None) -> str`  
- **Description:** Creates a section divider slide with a customizable background color and large, centered text.

### 5. Add Image Slide
- **Function:** `add_image_slide(prs_id: str, title: str, image_path: str, image_description: str) -> str`  
- **Description:** Adds an image slide. The image can be loaded from a local file or downloaded from a URL.

### 6. Add Table Slide
- **Function:** `add_table_slide(prs_id: str, title: str, headers: List[str], rows: List[List[str]]) -> str`  
- **Description:** Inserts a slide containing a table defined by column headers and rows of data.

### 7. Save Presentation
- **Function:** `save_presentation(prs_id: str, output_path: str) -> str`  
- **Description:** Saves the presentation to the specified output path, managing temporary directories if necessary.

### 8. Get Presentation Download Link
- **Function:** `get_presentation_download_link(prs_id: str) -> str`  
- **Description:** Returns a data URI with base64-encoded presentation data for direct browser download.

### 9. Get Presentation Info
- **Function:** `get_presentation_info(prs_id: str) -> str`  
- **Description:** Retrieves metadata such as slide count and details on available slide layouts.

### 10. Get Presentation Outline
- **Resource Endpoint:** `presentation://{prs_id}/outline`  
- **Description:** Provides a text representation of the presentation structure, including slide titles and content summaries.

### 11. Remove Slide
- **Function:** `remove_slide(prs_id: str, slide_index: int) -> str`  
- **Description:** Removes a slide identified by its 1-based index from the presentation.

### 12. Export to Base64
- **Function:** `export_to_base64(prs_id: str) -> str`  
- **Description:** Exports the presentation as a base64-encoded string (with the first 100 characters shown as a sample).

### 13. SVG Generator Prompt Function
- **Function:** `svggenerator_prompt(description: str) -> list[base.Message]`
- **Description:** Creates a prompt that instructs Claude to generate an SVG image based on a natural language description. The function returns a list of two messages:
  1. A system message that sets Claude's role as an SVG expert
  2. A user message containing the specific SVG request

### 14. Generate SVG Function
- **Function:** `generate_svg(prs_id: str, svg_markup: str, title: str = None, width: float = 6.0) -> str`
- **Description:** Takes SVG markup and adds it to a PowerPoint presentation:
  - Requires a presentation ID and the SVG markup
  - Optionally accepts a title and width parameter (default 6 inches)
  - Writes the SVG to a temporary file
  - Converts the SVG to PNG using the rsvg-convert tool
  - Creates a new slide in the presentation
  - Adds the title to the slide if provided
  - Positions and adds the PNG image to the slide
  - Cleans up temporary files
  - Returns a confirmation message with the slide position

## Error Handling & Debugging

- **Error Checks:**  
  Each tool validates input (e.g., verifying presentation IDs or file existence) and returns descriptive error messages.
  
- **Temporary Directories:**  
  The server ensures that files are saved in writable directories (typically `/tmp`) and falls back accordingly if a provided path is read-only.

- **Logging:**  
  Errors and version information are output to stderr to aid in debugging and monitoring.

## Contributing

Contributions are welcome. If you encounter issues or have suggestions for improvements, please open an issue or submit a pull request.
