Attribute VB_Name = "modDoc2Pdf"
Option Explicit

Function IMG2PDF(sPicFile, sPDFFile, Optional Silent As Boolean = True) As Boolean
On Error GoTo err:
Dim objDistiller As New ACRODISTXLib.PdfDistiller
Dim FSO As New Scripting.FileSystemObject
Dim objWord As New Word.Application
Dim objWordDoc As Word.Document
Dim objWordDocs As Word.Documents

Dim sPrevPrinter  As String
Dim sTempFile, sFolder

Set objWordDocs = objWord.Documents
FSO.CreateTextFile App.Path & "\Temp.doc", True
sTempFile = App.Path & "\Temp"
sPicFile = FSO.GetAbsolutePathName(sPicFile)
sFolder = FSO.GetParentFolderName(sPicFile)

If Len(sPDFFile) = 0 Then
  sPDFFile = FSO.GetBaseName(sPicFile) + ".pdf"
End If

If Len(FSO.GetParentFolderName(sPDFFile)) = 0 Then
  sPDFFile = sFolder + "\" + sPDFFile
End If

' Remember current active printer
sPrevPrinter = objWord.ActivePrinter

'objWord.ActivePrinter = "Acrobat PDFWriter"
objWord.ActivePrinter = "Acrobat Distiller"

' Open the Word document
Set objWordDoc = objWordDocs.Open(App.Path & "\Temp.doc")

'objWord.ActiveDocument.InlineShapes.AddPicture "C:\Documents and Settings\bijo\Desktop\Convert_Wo20248910142006\DOC To PDF\test.JPG"
objWordDoc.Shapes.AddPicture sPicFile, , True

' Print the Word document to the Acrobat Distiller - will generate a postscript (.ps) (temporary) file
objWord.ActiveDocument.PrintOut False, , , sTempFile

objWordDoc.Close wdSaveChanges  'wdDoNotSaveChanges
objWord.ActivePrinter = sPrevPrinter
objWord.Quit 'wdDoNotSaveChanges
Set objWord = Nothing

' Distill the postscript file to PDF
objDistiller.FileToPDF sTempFile, sPDFFile, "Print"
Set objDistiller = Nothing
FSO.DeleteFile (sTempFile)

FSO.DeleteFile App.Path & "\Temp.doc", True
Set FSO = Nothing

If Silent = False Then
    MsgBox "PDF File Created", vbInformation
End If
IMG2PDF = True
Exit Function

err:
If Silent = False Then
    MsgBox err.Description, vbExclamation
End If
IMG2PDF = False
End Function

Function DOC2PDF(sDocFile, sPDFFile)
On Error GoTo err:
Dim FSO
'if you want set the reference to > MS Word Object Library
Dim objWord 'As New Word.Application
Dim objWordDoc
Dim objWordDocs
Dim sPrevPrinter  As String
Dim objDistiller
Dim sTempFile, sFolder

Set objDistiller = CreateObject("PDFDistiller.PDFDistiller")
Set FSO = CreateObject("Scripting.FileSystemObject")
Set objWord = CreateObject("Word.Application")
Set objWordDocs = objWord.Documents

sTempFile = App.Path & "\Temp"
sDocFile = FSO.GetAbsolutePathName(sDocFile)
sFolder = FSO.GetParentFolderName(sDocFile)

If Len(sPDFFile) = 0 Then
  sPDFFile = FSO.GetBaseName(sDocFile) + ".pdf"
End If

If Len(FSO.GetParentFolderName(sPDFFile)) = 0 Then
  sPDFFile = sFolder + "\" + sPDFFile
End If

' Remember current active printer
sPrevPrinter = objWord.ActivePrinter

'objWord.ActivePrinter = "Acrobat PDFWriter"
objWord.ActivePrinter = "Acrobat Distiller"

' Open the Word document
Set objWordDoc = objWordDocs.Open(sDocFile, , True)

' Print the Word document to the Acrobat Distiller - will generate a postscript (.ps) (temporary) file
objWord.ActiveDocument.PrintOut False, , , sTempFile
objWordDoc.Close 'wdDoNotSaveChanges
objWord.ActivePrinter = sPrevPrinter
objWord.Quit 'wdDoNotSaveChanges
Set objWord = Nothing

' Distill the postscript file to PDF
objDistiller.FileToPDF sTempFile, sPDFFile, "Print"
Set objDistiller = Nothing
FSO.DeleteFile (sTempFile)

Set FSO = Nothing

MsgBox "PDF File Created", vbInformation
Exit Function

err:
MsgBox err.Description, vbExclamation
End Function
