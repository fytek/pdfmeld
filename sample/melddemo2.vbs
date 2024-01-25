' Sample of calling the DLL without a starting a server
' run this as: wscript melddemo2.vbs
Dim PDF
Set PDF = CreateObject("FyTek.PDFMeld")

' **** IMPORTANT ****
' set the location of the executable - change <my path> to where the program is actually installed on your system
PDF.setExe("<my path>\pdfmeld64.exe")

' set your credentials to avoid the pop-up window
' PDF.setKeyName("your key name")
' PDF.setKeyCode("your key code")

' get a temp area to work in - we'll place the output of the overlay here.
' not mandatory, just one way of keeping temporary files separate.
tmpdir = PDF.getTempDir()
PDF.setInFile("sample1.PDF")
PDF.setInFile("sample2.pdf")
PDF.setOutFile(tmpdir + "out.pdf")
PDF.setGUIOff()
PDF.setOverlay()

' stashes the above options for this part of the build
PDF.apply()

' take the output of the above and use it as the input
' we'll add page numbers and call the last output final.pdf
PDF.setInFile(tmpdir + "out.pdf")
PDF.setOutFile("final.pdf")
PDF.setPageNum()
PDF.setForce()

' set to true (default) to wait for PDF to finish before continuing.
' the temp dir is cleaned up after run is called so make sure you don't put your finished output there.
Set res = PDF.run(true)

WScript.Echo("cmd=" & res.cmd)

' if you leave the PDF.setOutFile("final.pdf") out, you may reference the return PDF in memory like this: res.bytes
' WScript.Echo("len=" & LenB(res.bytes))

