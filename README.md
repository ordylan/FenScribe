# FenScribe - Smart PDF Layout Optimizer

Reduce printing costs by automatically detecting and removing blank spaces in PDF documents.

## Samples
![gui图片](https://github.com/user-attachments/assets/b6c8e163-81a9-4b5d-9baa-5d8947292634)
![up_1733669026353_1_695650](https://github.com/user-attachments/assets/9cccd0da-39c6-4cd4-909f-7bfa13cb0086)

---

## Installation Dependencies

```bash
pip install PyMuPDF Pillow python-docx tkinterdnd2
```

---

## Usage

1. Run `gui.pyw` for graphical interface operations (CLI version currently unavailable)
2. After processing, use the provided Office macro(`图片溢出缩小.bas`) to resize overflowing images
   - *The macro automatically detects first column width in Word and scales oversized images*

---

## Configuration Parameters

| Parameter      | Description                                                                 |
|----------------|-----------------------------------------------------------------------------|
| **threshold**  | Brightness threshold for blank line detection (0-255)                      |
|                | - Converts RGB pixels to grayscale (average value)                         |
|                | - Rows with average grayscale ≥ threshold are considered blank             |
| **dpi**        | Image resolution when converting PDF to images                             |
| **min_height** | Content validity filter (in pixels)                                        |
|                | - Only preserves content blocks with height ≥ specified value             |
| **blank_height** | Paragraph separation baseline                                          |
|                | - Content is considered separate paragraphs when blank lines ≥ this value |

---

## Important Notes

⚠️ **Potential Issues**:
- Small images/geometric shapes may be accidentally removed or cropped
- Narrow images/color blocks might be misidentified (adjust config parameters)
- Scanned documents must have horizontal text alignment (pre-process tilted pages)
- Manual margin trimming required to remove headers/footers

---

## License

MIT License

---

## Development Notes

This third-generation version features:
- GUI implementation for user-friendly operation
- Partial utilization of AI-assisted development tools
- Continuous optimization through multiple iterations
