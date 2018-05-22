' Name: Beier (Benjamin) Liu
' Date: 5/21/2018

' Remark:
' Word Object required
Option Explicit
' ===================================================================================================
' File content:
' Excel to word
' ===================================================================================================

Sub ExcelToWord(filePath As String, strStaCol As String, strEndCol As String, startRow As Integer) 
' ==================================================================================================
' ARGUMENTS:
' filePath	--string, path of the word document
' strStaCol	--string, name of the first col to be pasted
' strEndCol	--string, name of the last col to be pasted
' startRow	--integer, number of first row to be pasted 
' RETURNS:
' A word document with excel table as data 
' ==================================================================================================

' Preparation Phrase
Dim tbl As Range
Dim objWord As Object 
Dim target As Object 
Dim WordTable As Object
Dim wkSheet As Worksheet
Set wkSheet=activeworkbook.activesheet

Dim rowCount As Integer
Dim lastRow As Integer
rowCount=wkSheet.Cells(Rows.Count, strStaCol).End(xlUp).row 
lastRow=rowCount

Application.ScreenUpdating = False
Application.EnableEvents = False
Set tbl = wkSheet.Range(strStaCol & startRow & ":" & strEndCol & lastRow)

On Error Resume Next
Set objWord = GetObject(, "Word.Application")
If objWord Is Nothing Then
Set objWord = CreateObject("Word.Application")
End If
On Error GoTo 0

' Handling Phrase
Set target=objWord.Documents.Open(filePath)

tbl.copy

target.Paragraphs(1).Range.PasteExcelTable _
	LinkedToExcel:=True, _
	WordFormatting:=False, _
	RTF:=False

Set WordTable = target.Tables(1)
WordTable.AutoFitBehavior wdAutoFitWindow


' Checking Phrase
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.CutCopyMode = False

Set objWord=Nothing
Set target=Nothing
Set WordTable=Nothing 
End Sub

Function ExcelToWord_help() as String

ExcelToWord_help="filePath As String, strStaCol As String, strEndCol As String, startRow As Integer"

End Function 

' https://www.thespreadsheetguru.com/blog/2014/5/22/copy-paste-an-excel-table-into-microsoft-word-with-vba