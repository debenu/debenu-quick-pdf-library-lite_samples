' Debenu Quick PDF Library Lite Sample

WScript.Echo("Debenu Quick PDF Library Lite - Add hyperlink to website Sample")

Dim ClassName
Dim QP

Dim FileName

ClassName = "DebenuPDFLibraryLite0916.PDFLibrary"
FileName = "web-link.pdf"

Set QP = CreateObject(ClassName)

' When the DQPL object is initiated a blank document
' is created and selected in memory by default.

' Set the origin for the co-ordinates to be the
' top left corner of the page.

Call QP.SetOrigin(1)

' Adding a link to the web is easy
' with the AddLinkToWeb function

Call QP.AddLinkToWeb(200, 100, 60, 20, "http://www.debenu.com", 1)

' Hyperlinks and text are two separate
' elements in a PDF, so we'll draw some
' text now so that you know where the
' hyperlink is located on the page.

Call QP.DrawText(205, 114, "Click me!")

' Save the document with the text you've just written to disk
If QP.SaveToFile(FileName) = 1 Then
WScript.Echo("File " + FileName + " written successfully")
Else
WScript.Echo("Error, file could not be written")
End If

Set QP = Nothing
WScript.Sleep(9000)