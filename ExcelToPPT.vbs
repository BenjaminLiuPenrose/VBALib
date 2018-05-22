' Name: Beier (Benjamin) Liu
' Date: 

' Remark:
' 
Option Explicit
' ===================================================================================================
' File content:
' Write comments
' ===================================================================================================

Sub ExcelToPPT() 
' ==================================================================================================
' ARGUMENTS:
' 
' RETURNS:
' 
' ==================================================================================================

' Preparation Phrase
Dim tbl As Range
Dim objPPT As Object 
Dim target As Object
Dim PPTSlide As Object
Dim PPTShape As Object
Dim wkSheet As Worksheet
Set wkSheet=activeworkbook.activesheet

Application.ScreenUpdating = False
Application.EnableEvents = False
Set tbl = wkSheet.Range("A1:B6")

On Error Resume Next
Set objWord = GetObject(, "PowerPoint.Application")
If objWord Is Nothing Then
Set objWord = CreateObject("PowrPoint.Application")
End If
On Error GoTo 0

' Handling Phrase
Set target=objPPT.Presentations.Open("")
Set PPTSlide=target.Sildes(1)

tbl.copy

target.Sildes(1).Range.PasteExcelTable _
	LinkedToExcel:=True, _
	WordFormatting:=False, _
	RTF:=False
PPTSlide.Shapes.PasteSpecial DataType:=2

Set PPTShape = PPTSlide.Shapes(PPTShape.Shapes.Count)
PPTShape.Left=66
PPTShape.Top=152	

' Checking Phrase
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.CutCopyMode = False

Set objPPT=Nothing
Set target=Nothing
Set PPTSlide=Nothing
Set PPTShape=Nothing
End Sub


Function ExcelToPPT_help() as String

ExcelToPPT_help="helper"

End Function 

' https://www.thespreadsheetguru.com/blog/2014/3/17/copy-paste-an-excel-range-into-powerpoint-with-vba
