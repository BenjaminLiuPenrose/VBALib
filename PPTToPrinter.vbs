' Name: Beier (Benjamin) Liu
' Date: 

' Remark:
' 
Option Explicit
' ===================================================================================================
' File content:
' Write comments
' ===================================================================================================

Sub PPTToPrinter(filePath As String, numCopies As Integer) 
' ==================================================================================================
' ARGUMENTS:
' 
' RETURNS:
' 
' ==================================================================================================

' Preparation Phrase
dim objPPT as Object
On Error Resume Next 

Set objPPT = GetObject(, "PowerPoint.Application")
If objPPT Is Nothing Then
Set objPPT = CreateObject("PowerPoint.Application")
End If
On Error GoTo 0

' Handling Phrase
Set target=objPPT.Presentations.Open(filePath)
target.UpdateLinks

dim fileName2 As String
fileName2=Replace(target, "pptx", "pdf")
If newFilePath is not Nothing Then
	fileName2=newFilePath
End If

With target.PrintOptions
	.NumberOfCopies = numCopies
	.Collate = msoTrue
	.OutputType = ppPrintOutputSlides
	.PrintHiddenSlides = msoTrue
	.PrintColorType = ppPrintColor
	.FitToPage = msoFalse
	.FrameSlides = msoFalse
	.ActivePrinter = "\\SHCOLOR CANON C5551" ' To be set
End With
target.PrintOut

' Checking Phrase
target.Close
objPPT.Quit

Set objPPT=Nothing
Set target=Nothing
End Sub

Function PPTToPrinter_help() as String

PPTToPrinter_help="filePath As String, numCopies As Integer"

End Function 

' https://answers.microsoft.com/en-us/msoffice/forum/msoffice_powerpoint-mso_other-mso_2010/vba-powerpoint-need-help-printing-user-selects/2b9b8871-a819-4a91-9a41-0618b93e045e 
' Printer comments 
