' Debenu Quick PDF Library Lite Sample

WScript.Echo("Debenu Quick PDF Library Lite - Draw Text Sample")

Dim ClassName
Dim QP

Dim FileName

ClassName = "DebenuPDFLibraryLite0916.PDFLibrary"
FileName = "text.pdf"

Set QP = CreateObject(ClassName)

' Set the origin for the co-ordinates to be the
' top left corner of the page. (optional)

Call QP.SetOrigin(1)

' Draw text on the blank document that's already in memory
Call QP.DrawText(100, 200, "Hello world from VBScript")

' Draw text in a text box. Specify width and height
' of the text box.

Call QP.DrawTextBox(350, 150, 200, 200, "This text was drawn using the DrawTextBox function. Similar to the DrawText function except that the alignment can be specified and line wrapping occurs.", 1)

Call QP.SetTextColor(0.9, 0.2, 0.5)
Call QP.SetTextSize(30)

Call QP.DrawText(100, 100, "Big and Colorful.")

' Save the document with the text you've just written to disk
If QP.SaveToFile(FileName) = 1 Then
WScript.Echo("File " + FileName + " written successfully")
Else
WScript.Echo("Error, file could not be written")
End If

Set QP = Nothing
WScript.Sleep(9000)