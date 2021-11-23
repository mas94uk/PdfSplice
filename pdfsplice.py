#!/usr/bin/env python3

# Tool to chop and splice PDFs.
# Takes ranges of pages from one or more PDFs and creates a new PDF from them.

# Requirements:
#  pip3 install pypdf2

import os
import sys
import re
from PyPDF2 import PdfFileWriter, PdfFileReader

# Add the specified pages from the input pdf
def add_pages(pages, input_pdf, page_numbers, append):
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

# We need at least two parameters - an output file and an input file
if len(sys.argv) < 3 or "--help" in sys.argv or "-h" in sys.argv:
    print("""Splice PDF files

usage: {0} OUTFILE FILE1 [PAGES] [PAGES...] [ FILE2 [=] [PAGES] [PAGES...] ...]

PAGES can be a single page (e.g. 7) or a range (e.g. 2-5 7- -11)

Examples:
  {0} out.pdf input1.pdf 1-5 10 20-
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

# Split the remaining parameters into sets, each starting with a file.
# Do this by seeing which are extant files.
remainingArgs = sys.argv[2:]
sourceSections = []
while remainingArgs:
    section = []
    # The next one should be a file
    file = remainingArgs.pop(0)
    if not os.path.isfile(file):
        print("%s is not a file" % file)
        exit(40)       
    section.append(file)
    
    # Append arguments until we find one which is the next file
    while remainingArgs and not os.path.isfile(remainingArgs[0]):
        section.append(remainingArgs.pop(0))

    sourceSections.append(section)

# Output PDF as an array of pages
outputPdfPages = []

# Process each input file in turn
for section in sourceSections:
    # The first item is the input file
    filename = section.pop(0)
    print("Input file: %s" % filename)
    infile = open(filename, "rb")
    sourcePdf = PdfFileReader(infile)

    append = True # Default behaviour
    pageNumbers = []

    # Each remaining item is either a mix spec or a page specification
    for arg in section:
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
            print("Unexpected argument %s" % arg)
            exit(50)

    # If we did not specify any pages, use all of them
    if not pageNumbers:
        print("No pages specified. Using all pages.")
        pageNumbers = range(0, sourcePdf.numPages)

    # We now know everything to retrieve from this file.
    add_pages(outputPdfPages, sourcePdf, pageNumbers, append)

# We have finished reading stuff.  Now compile a PDF from the array of pages.
outputPdf = PdfFileWriter()
for outputPdfPage in outputPdfPages:
    outputPdf.addPage(outputPdfPage)

# Write the generated PDF to disk
with open(outfile_name, "wb") as out:
    outputPdf.write(out)
