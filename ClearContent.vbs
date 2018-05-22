' Name: Beier (Benjamin) Liu
' Date: 

' Remark:
' 
Option Explicit
' ===================================================================================================
' File content:
' Clear Content 
' ===================================================================================================

Sub ClearContent(strStaCol As String, strEndCol As String, startRow As Integer) 
' ==================================================================================================
' ARGUMENTS:
' strStaCol		--string, start col number, also used for finding the last row
' strEndCol		--string, end col number, the contents in which will be deleted
' startRow		--integer, the start row from which the contents will be deleted
' RETURNS:
' wkSheet.Range(strStaCol & startRow & ":" & strEndCol & lastRow)	--all contents in the targer area will be deleted, but the format will be retained 
' ==================================================================================================

' Preparation Phrase
Dim wkSheet As Worksheet
Set wkSheet = activeworkbook.activesheet

Dim rowCount As Integer
Dim lastRow As Integer
rowCount = wkSheet.Cells(Rows.Count, strStaCol).End(xlUp).row
lastRow = rowCount

' Handling Phrase
wkSheet.Range(strStaCol & startRow & ":" & strEndCol & lastRow).ClearContents

' Checking Phrase
End Sub

Function ClearContent_help() as String

ClearContent_help="strStaCol As String, strEndCol As String, startRow As Integer"

End Function 
