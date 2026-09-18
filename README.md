# PdfSplice

PdfSplice is a command-line utility for building a PDF from selected pages of one or more existing PDF files. It can append pages, interleave pages from multiple PDFs, rotate pages, and help reassemble scanned booklets.

## What it does

Reads one or more source PDFs and creates an output PDF containing:

- specific page numbers
- page ranges such as `1-5` or `20-`
- rotated pages (`R90`, `R180`, `R270`, etc.)
- interleaved content from multiple PDFs using `=`
- booklet reassembly via `SPREADFIX` or `REASSEMBLE`

This is useful for tasks such as:

- extracting selected pages from PDFs
- combining PDFs in a custom order
- interleaving front and back scans
- restoring scan order for double-sided booklets

## Installation

From the project root:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Usage

```bash
python pdfsplice.py OUTFILE FILE1 [ROTATION] [PAGES] [ROTATION] [PAGES...] [FILE2 [=] [ROTATION] [PAGES] ...]
```

The script also supports two helper modes:

```bash
python pdfsplice.py OUTFILE FILE1 SPREADFIX
python pdfsplice.py OUTFILE fronts.pdf backs.pdf REASSEMBLE
```

Run with `-h` or `--help` to see the full built-in usage guide:

```bash
python pdfsplice.py --help
```

## Examples

Create a PDF from selected pages of one file:

```bash
python pdfsplice.py out.pdf input1.pdf 1-5 10 20-
```

Append or interleave pages from a second file:

```bash
python pdfsplice.py out.pdf input1.pdf 1-5 input2.pdf R90 2-2 R0 3-
python pdfsplice.py out.pdf fronts.pdf rears.pdf =
```

Reassemble a booklet from scanned spread-order pages:

```bash
python pdfsplice.py out.pdf scan1_split.pdf SPREADFIX
```

Restore a document from separately scanned fronts and backs:

```bash
python pdfsplice.py out.pdf fronts.pdf backs.pdf REASSEMBLE
```

## Rotation values

Available rotation values are:

- `R0` — no rotation
- `R90` — 90° clockwise
- `R180` — 180°
- `R270` or `R-90` — 270° clockwise / 90° counter-clockwise

## Notes

- Page numbering is 1-based in the CLI arguments, even though PDFs are internally indexed starting at 0.
