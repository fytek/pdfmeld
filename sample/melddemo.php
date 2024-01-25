<?php

# Sample of calling the DLL without a starting a server
$pdfobj = new COM('FyTek.PDFMeld');

# **** IMPORTANT ****
# set the location of the executable - change <my path> to where the program is actually installed on your system
$pdfobj->setExe("<my path>\pdfmeld64.exe");

# set your credentials to avoid the pop-up window
# PDF.setKeyName("your key name")
# PDF.setKeyCode("your key code")

# get a temp area to work in - we'll place the output of the overlay here.
# not mandatory, just one way of keeping temporary files separate.
$tmpdir = $pdfobj->getTempDir();
$pdfobj->setInFile("sample1.PDF");
$pdfobj->setInFile("sample2.pdf");
$pdfobj->setOutFile($tmpdir . "out.pdf");
$pdfobj->setGUIOff();
$pdfobj->setOverlay();

# stashes the above options for this part of the build
$pdfobj->apply();

# take the output of the above and use it as the input
# we'll add page numbers and call the last output final.pdf
$pdfobj->setInFile($tmpdir . "out.pdf");
$pdfobj->setOutFile("final.pdf");
$pdfobj->setPageNum();
$pdfobj->setForce();

# Set to true (default) to wait for PDF to finish before continuing
$res = $pdfobj->run(true);

print $res->cmd;
# leave $pdfobj->setOutFile("final.pdf"); out from the above to return the bytes of the PDF in $res->bytes

?>
