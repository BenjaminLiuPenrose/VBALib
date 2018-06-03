' Name: Beier (Benjamin) Liu
' Date:

' Remark:
'
Option Explicit
' ===================================================================================================
' File content:
' Clear Content
' ===================================================================================================

Sub SelectContent(strStaCol As String, strEndCol As String, startRow As Integer)
' ==================================================================================================
' ARGUMENTS:
' strStaCol		--string, start col number, also used for finding the last row
' strEndCol		--string, end col number, the contents in which will be deleted
' startRow		--integer, the start row from which the contents will be deleted
' RETURNS:
' wkSheet.Range(strStaCol & startRow & ":" & strEndCol & lastRow)	--all contents in the targer area will be selected
' ==================================================================================================

' Preparation Phrase
Dim wkSheet As Worksheet
Set wkSheet = activeworkbook.activesheet

Dim rowCount As Integer
Dim lastRow As Integer
rowCount = wkSheet.Cells(Rows.Count, strStaCol).End(xlUp).row
lastRow = rowCount

' Handling Phrase
wkSheet.Range(strStaCol & startRow & ":" & strEndCol & lastRow).Select

' Checking Phrase
End Sub

Function SelectContent_help() as String

SelectContent_help="strStaCol As String, strEndCol As String, startRow As Integer"

End Function
