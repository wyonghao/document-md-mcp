# document-md-mcp

MCP server that converts Word and PDF documents into:

- `example.md`
- `examplemd/` containing extracted images and optional rendered PDF pages
- `example.manifest.json` describing source, output files, pages, and images

This is meant for ingestion pipelines where the LLM reads Markdown text and a vision model reads the extracted images.

## Why This Exists

Many AI agents can read plain text files well, but they often handle Word and PDF files inconsistently. DOCX files are especially tricky because important visual content may be stored as:

- normal PNG/JPEG images
- Office vector images such as EMF/WMF
- VML shapes, text boxes, arrows, and diagrams
- tables and layout-heavy content

This MCP server makes the ingestion step explicit:

1. Convert the document into Markdown for text-first reading.
2. Extract visual assets into a side folder.
3. Normalize images so vision models can read them efficiently.
4. Return a manifest so an agent can decide exactly which files to inspect.

The goal is to avoid asking an AI model to blindly ingest the original DOCX/PDF. Instead, the text model reads the Markdown and the vision model reads only the extracted images that matter.

## Update: DOCX EMF/WMF Image Handling

Office documents often contain EMF/WMF vector images. These can be several megabytes and are not always directly supported by AI vision tools or Markdown viewers.

This server now converts DOCX EMF/WMF images to PNG when Pillow can render them, then resizes large images to a configurable maximum dimension. The default maximum side length is `1600` pixels.

This improves ingestion because:

- PNG is widely supported by vision models.
- large Office vector files are reduced before upload.
- text inside diagrams remains readable at practical sizes.
- agents spend less bandwidth and fewer vision tokens on oversized images.
- the original document text still stays in Markdown, so vision is used only for visual content.

Example from testing:

```text
image002.emf  image/x-emf  3.6 MB
  -> image002.png  image/png  ~173 KB, max side 1600 px
```

The manifest records original and output metadata, including source content type, output content type, dimensions, byte size, and whether the image was converted or resized.

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

## Recommended Agent Workflow

For efficient ingestion:

1. Call `convert_document` on the source DOCX/PDF.
2. Read the generated `.md` file for text.
3. Read `example.manifest.json` to discover extracted images.
4. Send only the generated PNG/JPEG images to the vision model.
5. Avoid sending every page as an image unless the document has complex layout that text extraction cannot represent.

This keeps the default workflow token-efficient while still preserving diagrams, figures, and Office-specific visual content.
