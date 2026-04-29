# FenScribe - A Smart PDF Layout Optimizer

Automatically detect and remove blank areas from images or PDFs, for reducing printing costs.

[Home Page](https://ordylan.com/FenScribe)

![GUI screenshot](https://github.com/user-attachments/assets/048019cb-35ae-46c2-9a93-2c05c60f17b3)
![Output](https://github.com/user-attachments/assets/e494587b-dd31-4ece-b33d-b0674d8b311d)

## Key Features

- Detect and remove long blank regions (by brightness threshold).[New - We may now disregard marginal scan edges! (in advanced option)]
- Split pages into content segments(pre line) and save them as individual images.
- Optional Word export: insert segments into a `.docx` template.
- [New Update!]Automatic image width fitting for Word documents (python-docx), reducing the need for post-processing macros.

## Requirements

This project was developed on Windows and assumes Microsoft Word is available for macro injection features.

Python packages:

```bash
pip install PyMuPDF Pillow python-docx tkinterdnd2 pywin32
```

- `pywin32` is required only if you want the program to inject/run VBA macros in Word (Windows only).
- If you only need image output or python-docx insertion, Word and `pywin32` are optional.

---

## Quick Start

1. Install dependencies (see Requirements).
2. Run the GUI:

```bash
python gui.pyw
```

Or on Windows you can double-click `gui.pyw`.

## Typical Workflow

1. Select a PDF file, single image, or a folder containing images.
2. Tune parameters in Step 2 (threshold, DPI, min_height, blank_height).
3. Choose a `.docx` template in Step 3 if you want Word output.
4. Start processing. Output files:

- Images / segments are saved under `_temp/<input>_segments`.
- If Word insertion is enabled, the final document is written to `__Output/output_<basename>.docx`.

## Configuration Parameters

- `threshold` (20–255): Brightness threshold for blank-line detection. Larger values are more conservative.
- `dpi` (72–600): Resolution used when converting PDF pages to images.
- `min_height`: Minimum segment height (px) to keep a segment.
- `blank_height`: Number of consecutive blank rows that separate paragraphs.
- `lr_threshold_enabled`: Enable left/right margin trimming for blank detection only.
- `lr_threshold_percent`: Percent width on each side to ignore when detecting blanks.
- `no_insert_word`: When set, only images are saved and Word insertion is skipped.

## Macros and Word Automation

- The app can optionally inject and run a VBA macro (from the `_Macros` folder) after producing the `.docx` file. That requires Microsoft Word and `pywin32`.
- For macro injection to work, enable **Trust access to the VBA project object model** in Word: `File → Options → Trust Center → Trust Center Settings → Macro Settings → Trust access to the VBA project object model`.
- Note: the program now attempts to auto-fit images to the document/column width using `python-docx`. The macro remains available as an optional post-processing step but is no longer strictly required for common cases.

## Troubleshooting

- If Word macro injection fails with an access error, verify the VBA trust setting described above.
- If Word is not installed or `pywin32` is missing, use the Advanced option `Do not insert Word` to only save images.
- For scanned documents with slanted pages, run a deskew/preprocessing step first.

## License

MIT License

## Development Notes

- This version introduces batch image insertion, automatic width-fit using `python-docx`, and progress throttling to improve performance and UX.
- Contributions and bug reports are welcome.
