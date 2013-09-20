' Debenu Quick PDF Library Lite Sample

WScript.Echo("Debenu Quick PDF Library Lite - Hello World Sample")

Dim ClassName
Dim QP

Dim FileName

ClassName = "DebenuPDFLibraryLite0916.PDFLibrary"
FileName = "hello-world.pdf"

Set QP = CreateObject(ClassName)

' Draw text on the blank document that's already in memory
Call QP.DrawText(100, 500, "Hello world from VBScript")

' Save the document with the text you've just written to disk
If QP.SaveToFile(FileName) = 1 Then
WScript.Echo("File " + FileName + " written successfully")
Else
WScript.Echo("Error, file could not be written")
End If

Set QP = Nothing
WScript.Sleep(9000)