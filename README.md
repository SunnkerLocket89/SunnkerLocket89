# Idaho4 Exhibits Parser

This repository provides a command line helper that automates the task of
downloading and organising the public exhibits listed in the
`Idaho4_exhibits_with_full_metadata.xlsx` spreadsheet.  The script reads the
spreadsheet, downloads the referenced PDF files, and optionally extracts the
first *N* pages of each document into a dedicated folder.

## Installation

Install the required third-party libraries in your Python environment before
running the parser.  The script now reads the workbook directly via
[`openpyxl`](https://openpyxl.readthedocs.io/) so no `pandas` dependency is
needed.  You can install the libraries individually or via the provided
`requirements.txt` file:

```bash
pip install -r requirements.txt
```

The above is equivalent to running `pip install requests openpyxl PyMuPDF
PyPDF2 tqdm`.

## Usage

```bash
python run_idaho4_parser.py \
  --in-file Idaho4_exhibits_with_full_metadata.xlsx \
  --sheet Exhibits_With_Metadata \
  --workers 6 \
  --extract-pages 4
```

By default the script stores the downloaded PDFs in `idaho4_output/downloads`
and writes a JSON manifest plus a CSV summary to `idaho4_output`.  Downloaded
files are prefixed with the zero-padded Excel row number to guarantee
unique filenames while keeping the on-disk order aligned with the worksheet.
The manifest records whether each row succeeded, was skipped (for example
because it did not contain a URL), or failed, and includes the corresponding
Excel row number for quick cross-referencing.  Re-run the command with
`--resume` to continue from where a previous session stopped without
re-downloading files.

### Common flags

- `--url-column` – Set the spreadsheet column that contains the PDF URL.  When
  omitted the script attempts to infer a sensible column automatically.
- `--id-column` – Configure the column that uniquely identifies each exhibit.
  This identifier is used to name the downloaded files.
- `--out-dir` – Choose a different destination directory for all generated
  artefacts.
- `--manifest` / `--csv` – Override the default manifest output paths.
- `--verbose` – Enable verbose logging for troubleshooting.

Run `python run_idaho4_parser.py --help` to see the full list of supported
flags.
