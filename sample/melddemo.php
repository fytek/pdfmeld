<?php

# Sample of calling the DLL without a starting a server
$pdfobj = new COM('FyTek.PDFMeld');

# **** IMPORTANT ****
# set the location of the executable - change <my path> to where the program is actually installed on your system
$pdfobj->setExe("<my path>\pdfmeld64.exe");

# set your credentials to avoid the pop-up window
# PDF.setKeyName("your key name")
# PDF.setKeyCode("your key code")

# Get a temp area to work in - we'll place the output of the overlay here.
# The current working directory will change to this directory so you may
# create interim files without conflicting with any other instances of 
# the program.
# Let's start by getting a temp area to work in.  Any files in the temp directory
# are deleted after the call to run().
$tmpdir = $pdfobj->getTempDir();
$pdfobj->setInFile("c:\mydir\sample1.PDF");
$pdfobj->setInFile("c:\mydir\sample2.pdf");
# The out.pdf will be in the temp working area, we'll use it as input to the next call.
$pdfobj->setOutFile("out.pdf");
$pdfobj->setGUIOff();
$pdfobj->setOverlay();

# Stashes the above options for this part of the build and prepare for the next call.
# You may have more calls to apply() if you need to do more processing, there is no set limit.
$pdfobj->apply();

# Take the output of the above and use it as the input.
# We'll add page numbers and call the last output final.pdf.
$pdfobj->setInFile("out.pdf");
$pdfobj->setOutFile("c:\mydir\final.pdf");
$pdfobj->setPageNum();

# Set to true (default) to wait for the PDF to finish before continuing.
# run() should always be the last step.  It takes all of the settings
# and processes the inputs to create a PDF.
$res = $pdfobj->run(true);

print $res->cmd;

#####################################3

# Now let's get the results in memory instead.
# To do that, leave out the setOutFile call.
# If you have one, enter your key name/code again since this is a new call to the program.
# PDF.setKeyName("your key name")
# PDF.setKeyCode("your key code")

# Let's start by getting a temp area to work in.
$tmpdir = $pdfobj->getTempDir();
$pdfobj->setInFile("c:\mydir\sample1.PDF");
$pdfobj->setInFile("c:\mydir\sample2.pdf");
$pdfobj->setGUIOff();
$pdfobj->setOverlay();

# Pass true again to wait for the PDF, note there is no setOutFile() call.
$res = $pdfobj->run(true);

print $res->cmd;
print "\n";
# If you want to check the size of the PDF
print $res->getNumBytes();
print "\n";
# You may optionally pass a from/thru byte range to getHexChunk() if want to read it in sections.
# If you don't, like below, the entire PDF is returned as a hex string.
$pdf = hex2bin($res=>getHexChunk());

# You may stream $pdf to a browser, store it in a database, etc.





?>
