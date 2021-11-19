#!/usr/bin/env python3
# Has to run in Python2 because PyPdf does not support Python3.

# Tool to chop and splice PDFs.
# Takes ranges of pages from one or more PDFs and creates a new PDF from them.

import argparse
import sys
import re
from PyPDF2 import PdfFileWriter, PdfFileReader
import code

# Parse the command line arguments

# We need at least one parameter - an input file
if len(sys.argv) < 2:
    print("Run with --help for instructions.")
    exit(1)

parser = argparse.ArgumentParser(description='Splice PDF files.', usage="[--outfile OUTFILE] FILE1 [PAGES] [[+|=] FILE2 [PAGES] ...]")
parser.add_argument('--outfile', type=argparse.FileType('w'), default=sys.stdout, help="Specify output file - default STDOUT")
parser.add_argument('inputSpec', nargs=argparse.REMAINDER, help="Input files, ranges and mix specs. Use '=' to interleave pages, or + (default) to append. PAGES can be a range, e.g. 2-5, 7-, -11 etc.") 
args = parser.parse_args()

# Get the output file
outfile = args.outfile.name

# Output PDF as an array of pages
outputPdfPages = []

# The current source PDF
sourcePdf = None
pagesOutputFromCurrentPdf = 0

# Each remaining parameter is either an input file, or a page specification
remainingArgs = args.inputSpec
numArgs = len(remainingArgs)
nextArg=0
append=True # Default behaviour
while nextArg<numArgs:
    arg = remainingArgs[nextArg]
    nextArg = nextArg + 1

    # We may specify a mix spec - "+" for append or "=" for interleave
    if "=" == arg:
        append = False
    elif "+" == arg:
        append = True
    else:
        # The first (or next) argument must be an input PDF file name. Open it.
        print("New file: %s" % arg)
        infile = open(arg, "rb")
        sourcePdf = PdfFileReader(infile)
        pagesOutputFromCurrentPdf = 0

        # The next argument(s) may be page specifiers
        # Examples:
        #  1-7
        #  3
        #  2-
        #  -3
        #  -
        #  1,3-5,9-
        pageNumbers = []
        while nextArg<numArgs and re.search("^((\d*\-\d*)|\d+)(,(\d*\-\d*)|\d+)*$", remainingArgs[nextArg]):
            arg = remainingArgs[nextArg]
            nextArg = nextArg + 1

            # Generate the list of page numbers in this part
            terms = arg.split(",")
            for term in terms:
                # "1-9" or similar
                if re.search("(\d*)\-(\d*)", term):
                    # Get the numeric values, if any
                    matches = re.search("(\d*)\-(\d*)", term)
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
                    print("start %d, end %d" % (start, end))
                    if end < start:
                        # Allow us to reverse the order of a range of pages, e.g. 7-3
                        pageNumbers = pageNumbers + range(start-1, end-2, -1)
                    else:
                        pageNumbers = pageNumbers + range(start-1, end)
                # "7" or similar
                elif re.search("(\d+)", term):
                    match = re.search("(\d+)", term)
                    pageNumber = int(match.group(1))
                    # Sanity check
                    if pageNumber < 1 or pageNumber > sourcePdf.numPages:
                        print("Invalid page number %d: document has %d pages." % (pageNumber, sourcePdf.numPages))
                        exit(30)
                    # Start counting at 0
                    pageNumber = pageNumber - 1
                    pageNumbers = pageNumbers + [ pageNumber ]

        # If we did not specify any pages, use all of them
        if 0 == len(pageNumbers):
            print("No pages specified.  Using all pages.")
            pageNumbers = range(0, sourcePdf.numPages)

        # Add the pages from this file to the list
        print ("Pages: ", pageNumbers)
        print ("Mode: ", "Append" if append else "Interleave")
        if (append):
            # Simply add the new pages onto the end of the list
            for pageNumber in pageNumbers:
                outputPdfPages.append(sourcePdf.getPage(pageNumber))
                pagesOutputFromCurrentPdf = pagesOutputFromCurrentPdf + 1
        else:
            # Interleave the pages into the existing array, starting AFTER the first existing item.
            # The first page is page 0, so we will insert odd-numbered pages.
            insertPosition = 1
            for pageNumber in pageNumbers:
                outputPdfPages.insert(insertPosition, sourcePdf.getPage(pageNumber))
                insertPosition = insertPosition + 2

        # We have finished parsing this file.  Default back to 'append for the next one
        append=True


# We have finished reading stuff.  Now compile a PDF from the array of pages.
outputPdf = PdfFileWriter()
for outputPdfPage in outputPdfPages:
    outputPdf.addPage(outputPdfPage)

# Write the generated PDF to disk
with open(outfile, "wb") as out:
    outputPdf.write(out)
