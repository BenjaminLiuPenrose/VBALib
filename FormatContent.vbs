' Name: Beier (Benjamin) Liu
' Date:

' Remark:
'
Option Explicit
' ===================================================================================================
' File content:
' Clear Content
' ===================================================================================================

Sub FormatContent(strStaCol As String, strEndCol As String, startRow As Integer, endRow As Integer)
' ==================================================================================================
' ARGUMENTS:
' strStaCol		--string, start col number, also used for finding the last row
' strEndCol		--string, end col number, the contents in which will be deleted
' startRow		--integer, the start row from which the contents will be deleted
' RETURNS:
' wkSheet.Range(strStaCol & startRow & ":" & strEndCol & lastRow)	--all contents in the targer area will be formated with default format
' ==================================================================================================

' Preparation Phrase
Dim wkSheet As Worksheet
Set wkSheet = activeworkbook.activesheet

' Dim rowCount As Integer
' Dim lastRow As Integer
' rowCount = wkSheet.Cells(Rows.Count, strStaCol).End(xlUp).row
' lastRow = rowCount

' Handling Phrase
With wkSheet.Range(strStaCol & startRow & ":" & strEndCol & endRow).Interior
	.Pattern=xlSolid
	.PatternColorIndex=xlAutomatic
	.ThemeColor=xlThemeColorDark1
	.TintAndShade=-0.1499679555605 ' 0
	.PatternTintAndShade=0
End With

' wkSheet.Raneg("A3").Copy
' wkSheet.Range(strStaCol & startRow & ":" & strEndCol & endRow).PasteSpecial Paste:=xlPasteFormats

' Checking Phrase
End Sub

Function ClearContent_help() as String

ClearContent_help="strStaCol As String, strEndCol As String, startRow As Integer, endRow As Integer"

End Function

' roadmap	-- upgrade to optional ending point but with rowCount and colCount
