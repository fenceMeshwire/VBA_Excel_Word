Option Explicit

Sub create_word_document()

Dim wkbBook As Excel.Workbook
Dim wksSheet As Excel.Worksheet
' ______________________________________________________________________
' Interface parameters:
Dim objWordApp As Object
Dim objWordDoc As Object
Dim strPath As String
' ______________________________________________________________________
strPath = ThisWorkbook.Path & "\" & "document.docx"

Set wkbBook = ThisWorkbook
Set wksSheet = Sheet1

' Create a Word document
Set objWordApp = CreateObject("Word.Application")

With objWordApp

  Set objWordDoc = .documents.Add
  ' Adding text to the document:
  .Selection.typetext Text:="Object Header"
  .Selection.typetext Text:=vbCrLf
  .Selection.typetext Text:="Next line..."
  ' ...
End With

' Save word document
objWordDoc.saveas2 Filename:=strPath
objWordApp.Quit

' Reset the objects
Set wkbBook = Nothing
Set wksSheet = Nothing
Set objWordApp = Nothing
Set objWordDoc = Nothing

End Sub
