' Debenu Quick PDF Library Lite Sample

WScript.Echo("Debenu Quick PDF Library Lite - Fonts and Text")

Dim ClassName
Dim QP

Dim FileName
Dim fontID1
Dim fontID2
Dim fontID3
Dim fontID4
Dim fontID5

ClassName = "DebenuPDFLibraryLite0916.PDFLibrary"
FileName = "different-fonts.pdf"

Set QP = CreateObject(ClassName)

'  Use the AddStandardFont function to add a font to
'  the default blank document and get the return
'  value which is the font ID.

fontID1 = QP.AddStandardFont(0)

'  Select the font using its font ID

Call QP.SelectFont(fontID1)

'  Draw some text onto the document to see if
'  everything is working OK.

Call QP.DrawText(100, 700, "Courier")

'  Repeat exercise to see what a couple of other
'  fonts will look like as well.

fontID2 = QP.AddStandardFont(1)
Call QP.SelectFont(fontID2)
Call QP.DrawText(100, 650, "CourierBold")

fontID3 = QP.AddStandardFont(2)
Call QP.SelectFont(fontID3)
Call QP.DrawText(100, 600, "CourierBoldOblique")

fontID4 = QP.AddStandardFont(3)
Call QP.SelectFont(fontID4)
Call QP.DrawText(100, 550, "Helvetica")

fontID5 = QP.AddStandardFont(4)
Call QP.SelectFont(fontID5)
Call QP.DrawText(100, 500, "HelveticaBold")

' Save the document with the text you've just written to disk
If QP.SaveToFile(FileName) = 1 Then
WScript.Echo("File " + FileName + " written successfully")
Else
WScript.Echo("Error, file could not be written")
End If

Set QP = Nothing
WScript.Sleep(9000)