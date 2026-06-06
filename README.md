# document-md-mcp

MCP server that converts Word and PDF documents into:

- `example.md`
- `examplemd/` containing extracted images and optional rendered PDF pages
- `example.manifest.json` describing source, output files, pages, and images

This is meant for ingestion pipelines where the LLM reads Markdown text and a vision model reads the extracted images.

## Supported Inputs

- `.docx`
- `.pdf`

## Run

From this directory:

```powershell
python server.py
```

Example MCP client config:

```json
{
  "mcpServers": {
    "document-md": {
      "command": "python",
      "args": ["server.py"],
      "cwd": "C:\\LeoPortable\\apps\\MCPs\\document-md-mcp"
    }
  }
}
```

## Tools

### `convert_document`

Convert one `.docx` or `.pdf`.

Arguments:

- `input_path`: path to the source document
- `output_dir`: optional output directory; defaults to the source file directory
- `overwrite`: replace existing markdown/manifest/images when true
- `pdf_image_strategy`: `embedded`, `pages`, or `both`; defaults to `both`
- `pdf_page_dpi`: DPI used when rendering PDF pages; defaults to `180`
- `max_image_dimension`: max width/height for extracted images; defaults to `1600`
- `convert_vector_images`: convert DOCX EMF/WMF images to PNG when possible; defaults to `true`

Returns a JSON object with output paths and image metadata.

### `supported_document_types`

Returns supported file extensions.

## Output Layout

For `C:\docs\example.docx`:

```text
C:\docs\example.md
C:\docs\example.manifest.json
C:\docs\examplemd\
  image001.png
  image002.jpeg
```

For PDFs, rendered pages are named `page001.png`, `page002.png`, and embedded images are named like `page001-image001.png`.

DOCX vector images such as EMF/WMF are converted to PNG by default when Pillow can render them. This is usually better for AI vision ingestion and avoids sending large Office vector files.
