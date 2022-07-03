#!/usr/bin/env python3

# Tool to chop and splice PDFs.
# Takes ranges of pages from one or more PDFs and creates a new PDF from them.

# Requirements:
#  pip3 install pypdf2

import os
import sys
import re
from PyPDF2 import PdfFileWriter, PdfFileReader

# We need at least two parameters - an output file and an input file
if len(sys.argv) < 3 or "--help" in sys.argv or "-h" in sys.argv:
    print("""Splice PDF files

usage: {0} OUTFILE FILE1 [ROTATION] [PAGES] [ROTATION] [PAGES...] [ FILE2 [=] [ROTATION] [PAGES] [ROTATION] [PAGES...] ...]

       {0} OUTFILE FILE1 SPREADFIX

ROTATION can be R0 (no rotation), R90 (90 deg clockwise), R180 or R-180 (180 degrees), R270 or R-90 (270 deg clockwise or 90deg anti-clockwise) and applies to all subsequent pages from the source file. Note that a page can only appear in the ouput with a single rotation.

PAGES can be a single page (e.g. 7) or a range (e.g. 2-5 7- -11)

SPREADFIX reassembles (reorders and rotates) a scanned booklet from "printer spread" order, which must already have been split into individual pages. (Google 'briss' tool for the splitting.)

To interleave pages from a second (or later) file, for example to reassemble a set of front-side and rear-side scans, specify = after the filename.

Examples:
  {0} out.pdf input1.pdf 1-5 10 20-
  Create out.pdf from input1.pdf pages 1-5, 10 and 20 onwards

  {0} out.pdf input1.pdf 1-5 input2.pdf R90 2-2 R0 3-
  Create out.pdf from input1.pdf pages 1-5 followed by input2.pdf pages 2-3, rotated 90deg clockwise, and pages 3 onwards, not rotated.

  {0} out.pdf fronts.pdf rears.pdf =
  Create out.pdf by interleaving fronts.pdf and rears.pdf, page by page

  {0} out.pdf scan1_split.pdf SPREADFIX
  Restructure a pre-split scan in "printer spread" order, into normal reading order.  
""".format(sys.argv[0]))
    exit(1)

# Get the output file
outfile_name = sys.argv[1]
print("Generating %s" % outfile_name)

remainingArgs = sys.argv[2:]

# If we are in "spreadfix" mode, generate our own set of arguments
if len(sys.argv)==4 and sys.argv[3].upper()=="SPREADFIX":
    # Get the number of pages in the source file
    filename = sys.argv[2]
    tempReader = PdfFileReader(filename)
    numPages = tempReader.numPages
    tempReader = None

    # Page count must be a multiple of 4
    if (numPages % 4):
        print("Source file page count %d must be a multiple of 4" % numPages)
        exit(2)

    remainingArgs = [filename]

    # The page order (for e.g. a 36 page book) is 35, 33, 31 ... 1, 2, 4, 6 ... 36.
    # Because of the printer layout, rotation switches after each page, except between 1 and 2.
    # (Proper explanation of this is too complicated -- draw it out on paper.)
    rotation = 90
    for page in range (numPages-1, 0, -2):
        remainingArgs.append("R%d" % rotation)
        remainingArgs.append(str(page))
        rotation = -rotation
    rotation = -rotation
    for page in range (2, numPages+1, 2):
        remainingArgs.append("R%d" % rotation)
        remainingArgs.append(str(page))
        rotation = -rotation

# Split the remaining parameters into sets, each starting with a file.
# Do this by seeing which are extant files.
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
    rotation = 0

    # Each remaining item is either a mix spec or a page specification
    for arg in section:
        # We may specify a mix spec - "+" for append or "=" for interleave
        if "=" == arg:
            append = False
        # If the argument is a single page number, e.g. "7"
        elif re.match("^\d+$", arg):
            match = re.match("^(\d+)$", arg)
            pageNumber = int(match.group(1))
            # Sanity checks
            if not sourcePdf:
                print("Must specify an input PDF before a page number")
                exit(20)
            if pageNumber < 1 or pageNumber > sourcePdf.numPages:
                print("Invalid page number %d: document has %d pages." % (pageNumber, sourcePdf.numPages))
                exit(30)
            # Start counting at 0
            print("Page %d, rotation %d" % (pageNumber, rotation))
            pageNumbers.append((pageNumber-1, rotation))
        # If a rotation is specified
        elif re.match("^R\-?\d+$", arg):
            match = re.match("^R(\-?\d+)", arg)
            rotation = int(match.group(1))
        # If the argument(s) is a range page specifier, e.g. 1-7
        elif re.search("^\d*\-\d*$", arg):
            # Get the numeric values, if any
            matches = re.search("^(\d*)\-(\d*)$", arg)
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
            print("Pages %d - %d, rotation %d" % (start, end, rotation))
            if end < start:
                # Allow us to reverse the order of a range of pages, e.g. 7-3
                for p in range(start-1, end-2, -1):
                    pageNumbers.append((p, rotation))
            else:
                for p in range(start-1, end):
                    pageNumbers.append((p, rotation))
        else:
            print("Unexpected argument %s" % arg)
            exit(50)

    # If we did not specify any pages, use all of them
    if not pageNumbers:
        print("No pages specified. Using all pages.")
        for p in range(0, sourcePdf.numPages):
            pageNumbers.append((p, rotation))

    # We now know everything to retrieve from this file.
    # Add the specified pages from the input pdf
    print ("Mode: " + "Append" if append else "Interleave")
    increment = 2 if append else 1
    insertPosition = 1 if append else 0
    for page_number, rotation in pageNumbers:
        page = sourcePdf.getPage(page_number)
        if rotation == 0:
            pass
        elif rotation == 90:
            page.rotateClockwise(90)
        elif rotation == 180 or rotation == -180:
            page.rotateClockwise(180)
        elif rotation == 270 or rotation == -90:
            page.rotateCounterClockwise(90)
        else:
            print("Invalid rotation: %d" % rotation)

        outputPdfPages.insert(insertPosition, page)

        insertPosition += increment

# We have finished reading stuff.  Now compile a PDF from the array of pages.
outputPdf = PdfFileWriter()
for outputPdfPage in outputPdfPages:
    outputPdf.addPage(outputPdfPage)

# Write the generated PDF to disk
with open(outfile_name, "wb") as out:
    outputPdf.write(out)
