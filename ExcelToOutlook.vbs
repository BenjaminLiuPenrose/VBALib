' Name: Beier (Benjamin) Liu
' Date: 5/21/2018

' Remark:
' Outlook reference in VBA required 
Option Explicit
' ===================================================================================================
' File content:
' Excel to VBA
' ===================================================================================================

Sub ExcelToOutlook(strRange As String, strMailTo As String, strCCTo As String, strBCCTo As String, strSubject As String, strBody As String, strAttPath As String) 
' ==================================================================================================
' ARGUMENTS:
' strRange	--string, range of the excel table
' strMailTo	--string, email address
' strCCTo 	--string, email address, if no, use ""
' strBCCTo 	--string, email address, if no, use "" 
' strSubject	--string, email subject, if no, use ""
' strBody 	--string, email context, if no, use ""
' 			e.g. strBody = "您好," & "<br><br>" & _
' 							"报价如下:" & "<br><br>"
' strAttPath	--string, email attachment path, if no, use ""
' RETURNS:
' Send out an email with address, title, table body 
' ==================================================================================================

' Preparation Phrase
dim wkSheet As Object
Set wkSheet=activeworkbook.activesheet

dim rng As Range
dim objOutlook As Object
dim target As Object

Application.ScreenUpdating=False
Application.EnableEvents=False
Set rng=wkSheet.Range(strRange)

On Error Resume Next
Set objOutlook = GetObject(, "Outlook.Application")
If objOutlook Is Nothing Then
Set objOutlook = CreateObject("Outlook.Application")
End If
On Error GoTo 0

' Handling Phrase
Set target=objOutlook.CreateItem(0)

On Error Resume Next
With target
	.to= strMailTo
	.CC= strCCTo
	.BCC=strBCCTo
	.Subject= strSubject
	.HTMLBody=StrBody & RangeToHTML(rng)
	.Attachments.Add(strAttPath)
	.Send
End With
On Error GoTo 0


' Checking Phrase
Application.ScreenUpdating=True
Application.EnableEvents=True

Set objOutlook =Nothing
Set target=Nothing
End Sub

Function ExcelToOutlook_help() as String

ExcelToOutlook_help="strRange As String, strMailTo As String, strCCTo As String, strBCCTo As String, strSubject As String, strBody As String, strAttPath As String"

End Function 


Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
Dim fso As Object
Dim ts As Object
Dim TempFile As String
Dim TempWB As Workbook

TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

'Copy the range and create a new workbook to past the data in
rng.Copy
Set TempWB = Workbooks.Add(1)
With TempWB.Sheets(1)
	.Cells(1).PasteSpecial Paste:=8
	.Cells(1).PasteSpecial xlPasteValues, , False, False
	.Cells(1).PasteSpecial xlPasteFormats, , False, False
	.Cells(1).Select
	Application.CutCopyMode = False
	On Error Resume Next
	.DrawingObjects.Visible = True
	.DrawingObjects.Delete
	On Error GoTo 0
End With

'Publish the sheet to a htm file
With TempWB.PublishObjects.Add( _
	 SourceType:=xlSourceRange, _
	 Filename:=TempFile, _
	 Sheet:=TempWB.Sheets(1).Name, _
	 Source:=TempWB.Sheets(1).UsedRange.Address, _
	 HtmlType:=xlHtmlStatic)
	.Publish (True)
End With

'Read all data from the htm file into RangetoHTML
Set fso = CreateObject("Scripting.FileSystemObject")
Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
RangetoHTML = ts.readall
ts.Close
RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
				"align=left x:publishsource=")

'Close TempWB
TempWB.Close savechanges:=False

'Delete the htm file we used in this function
Kill TempFile

Set ts = Nothing
Set fso = Nothing
Set TempWB = Nothing
End Function

' https://stackoverflow.com/questions/48896499/copy-excel-range-as-picture-to-outlook/48897439#48897439