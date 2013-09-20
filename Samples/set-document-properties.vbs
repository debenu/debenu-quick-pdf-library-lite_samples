' Debenu Quick PDF Library Lite Sample

WScript.Echo("Debenu Quick PDF Library Lite - Set Document Properties Sample")

Dim ClassName
Dim QP

Dim FileName

ClassName = "DebenuPDFLibraryLite0916.PDFLibrary"
FileName = "set-document-properties.pdf"

Set QP = CreateObject(ClassName)

' Draw text on the blank document that's already in memory
Call QP.DrawText(50, 500, "Open this PDF in Adobe Reader and press Ctrl + D to see the document properties for this PDF.")

' Set the Author and Title for this document
Call QP.SetInformation(1, "Debenu")
Call QP.SetInformation(2, "Sample Document Properties")

' Save the document with the text you've just written to disk
If QP.SaveToFile(FileName) = 1 Then
WScript.Echo("File " + FileName + " written successfully")
Else
WScript.Echo("Error, file could not be written")
End If

Set QP = Nothing
WScript.Sleep(9000)