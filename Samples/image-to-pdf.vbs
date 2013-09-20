' Debenu Quick PDF Library Lite Sample

WScript.Echo("Debenu Quick PDF Library Lite - Hello World Sample")

Dim ClassName
Dim QP

Dim InFileName
Dim OutFileName
Dim lWidth
Dim lHeight

ClassName = "DebenuPDFLibraryLite0916.PDFLibrary"
InFileName = "image.png"
OutFileName = "image.pdf"

Set QP = CreateObject(ClassName)

' Load the image that you'd like to convert to PDF
Call QP.AddImageFromFile(InFileName, 0)

' Get the width and height of the image
lWidth = QP.ImageWidth()
lHeight = QP.ImageHeight()

' Reformat the size of the page in the selected document 
Call QP.SetPageDimensions(lWidth, lHeight)

' Draw the image onto the page using the specified width/height
Call QP.DrawImage(0, lHeight, lWidth, lHeight)

' Save the new PDF to disk
If QP.SaveToFile(OutFileName) = 1 Then
WScript.Echo("File " + OutFileName + " written successfully")
Else
WScript.Echo("Error, file could not be written")
End If

Set QP = Nothing
WScript.Sleep(9000)