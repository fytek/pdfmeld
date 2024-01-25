# pdfmeld
FyTek PDF Meld DLL

This program is supplied as a compiled .NET DLL assembly with the executable download of PDF Meld. You may download the source and compile yourself (instructions are at the top of the code) if you want to make any custom changes.

The purpose is to allow a DLL method interface to the executable program pdfmeld.exe or pdfmeld64.exe. This DLL, which is compiled as both a 32-bit and 64-bit version, may be used in Visual Basic program, ASP, C#, etc. to send the information for building and retreiving PDFs from PDF PDF Meld. This DLL also replaces the old DLL version that was limited to running PDF Meld on the same box and had occassional memory issues.

There are commands to start a PDF Meld server which loads the executable version of PDF Meld in memory and listens on the port you specify for commands. This also allows you to have PDF Meld running on a different server for load balancing by not building the PDF on the same box as the requestor. You may also run several instances of PDF Meld server all on different boxes if you wish and the DLL will cycle between them to handle requests.

To compile (use Run as Administrator):
<pre>
csc /target:library /platform:anycpu /out:pdfmeld_20.dll pdfmeld_20.cs /keyfile:mykey.snk
to create the DLL: C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /codebase c:pdfmeld_20.dll
to create a tlb file for old COM: C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe -tlb /codebase c:pdfmeld_20.dll
</pre>
Run the vbs samples like this: cscript.exe (or wscript.exe) testprog.vbs
