' Name: Beier (Benjamin) Liu
' Date: 

' Remark:
' 
Option Explicit
' ===================================================================================================
' File content:
' Populate Formula
' ===================================================================================================

Sub PopulateFormula(strFormula As String, strRefCol As String, strCol As String, startRow As Integer) 
' ==================================================================================================
' ARGUMENTS:
' strFormula	--string, formula
' strRefCol		--string, reference col number, used for finding the last row
' strCol 		--string, target col number, the formula will be populates along the col
' startRow 		--string, the start row from which the formula will be populated
' RETURNS:
' wkSheet.range(strCol & startRow & ":" & strCol & lastRow).strFormula	-target area filled with formula 
' ==================================================================================================

' Preparation Phrase
Dim wkSheet As Worksheet
Set wkSheet = activeworkbook.activesheet

Dim rowCount As Integer
Dim lastRow As Integer
rowCount = wkSheet.Cells(Rows.Count, strRefCol).End(xlUp).row
lastRow = rowCount

' Handling Phrase
wkSheet.Range(strCol & startRow & ":" & strCol & lastRow).Formula = strFormula

' Checking Phrase
End Sub

Function PopulateFormula_help() as String

PopulateFormula_help="strFormula As String, strRefCol As String, strCol As String, startRow As Integer"

End Function 
