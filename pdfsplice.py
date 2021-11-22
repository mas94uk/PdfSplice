#!/usr/bin/env python3

# Tool to chop and splice PDFs.
# Takes ranges of pages from one or more PDFs and creates a new PDF from them.

import sys
import re
from PyPDF2 import PdfFileWriter, PdfFileReader

# Add the specified pages from the input pdf
def add_pages(pages, input_pdf, page_numbers, append):
    # If we did not specify any pages, use all of them
    if not page_numbers:
        print("No pages specified.  Using all pages.")
        page_numbers = range(0, input_pdf.numPages)

    # Add the pages from this file to the list
    print ("Mode: " + "Append" if append else "Interleave")
    if (append):
        # Simply add the new pages onto the end of the list
        for page_number in page_numbers:
            outputPdfPages.append(sourcePdf.getPage(page_number))
    else:
        # Interleave the pages into the existing array, starting AFTER the first existing item.
        # The first page is page 0, so we will insert odd-numbered pages.
        insertPosition = 1
        for page_number in page_numbers:
            pages.insert(insertPosition, sourcePdf.getPage(page_number))
            insertPosition = insertPosition + 2

# Parse the command line arguments

# We need at least one parameter - an input file
if len(sys.argv) < 3 or "--help" in sys.argv or "-h" in sys.argv:
    print("""Splice PDF files

usage: {0} OUTFILE FILE1 [PAGES] [PAGES...] [ FILE2 [=] [PAGES] [PAGES...] ...]

PAGES can be a single page (e.g. 7) or a range (e.g. 2-5 7- -11)

Examples:
  {0} out.pdf input1.pdf 1-5 10 20-")
  Create out.pdf from input1.pdf pages 1-5, 10 and 20 onwards

  {0} out.pdf input1.pdf 1-5 input2.pdf 2-
  Create out.pdf from input1.pdf pages 1-5 followed by input2.pdf pages 2 onwards

  {0} out.pdf input1.pdf input2.pdf =
  Create out.pdf by interleaving input1.pdf and input2.pdf, page by page
""".format(sys.argv[0]))
    exit(1)

# Get the output file
outfile_name = sys.argv[1]
print("Generating %s" % outfile_name)

# Output PDF as an array of pages
outputPdfPages = []

# The current source PDF
sourcePdf = None

# Each remaining parameter is either an input file, or a page specification
remainingArgs = sys.argv[2:]
append = True # Default behaviour
sourcePdf = None
pageNumbers = []
for arg in remainingArgs:
    # We may specify a mix spec - "+" for append or "=" for interleave
    if "=" == arg:
        append = False
    # If the argument is a single page number, e.g. "7"
    elif re.search("^(\d+)$", arg):
        match = re.search("(\d+)", arg)
        pageNumber = int(match.group(1))
        # Sanity checks
        if not sourcePdf:
            print("Must specify an input PDF before a page number")
            exit(20)
        if pageNumber < 1 or pageNumber > sourcePdf.numPages:
            print("Invalid page number %d: document has %d pages." % (pageNumber, sourcePdf.numPages))
            exit(30)
        # Start counting at 0
        print("Page %d" % pageNumber)
        pageNumbers.append(pageNumber-1)
    # If the argument(s) is a range page specifier, e.g. 1-7
    elif re.search("(\d*)\-(\d*)", arg):
        # Get the numeric values, if any
        matches = re.search("(\d*)\-(\d*)", arg)
        start = matches.group(1)
        end = matches.group(2)
        # If a value is omitted, use first or last page
        if start=="":
            start = 1
        if end=="":
            end = sourcePdf.numPages
        start = int(start)
        start = max(1,start)

        # Don't allow pages before the start or after the end
        end = int(end)
        end = min(sourcePdf.numPages, end)

        # Get the range. Note that humans number pages starting at 1, but PDF numbers them starting at 0
        print("Pages %d - %d" % (start, end))
        if end < start:
            # Allow us to reverse the order of a range of pages, e.g. 7-3
            pageNumbers.extend(range(start-1, end-2, -1))
        else:
            pageNumbers.extend(range(start-1, end))
    else:
        # The argument must be an input PDF file name.
        # We have finished with the previous file (if we have one) so get the specified pages from it,
        # and get everything ready for the next file.
        if sourcePdf:
            add_pages(outputPdfPages, sourcePdf, pageNumbers, append)
            pageNumbers = []
            append = True
        #  Open the new file.
        print("Input file: %s" % arg)
        infile = open(arg, "rb")
        sourcePdf = PdfFileReader(infile)

# Add pages from the last input file
add_pages(outputPdfPages, sourcePdf, pageNumbers, append)

# We have finished reading stuff.  Now compile a PDF from the array of pages.
outputPdf = PdfFileWriter()
for outputPdfPage in outputPdfPages:
    outputPdf.addPage(outputPdfPage)

# Write the generated PDF to disk
with open(outfile_name, "wb") as out:
    outputPdf.write(out)
