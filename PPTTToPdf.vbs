' Name: Beier (Benjamin) Liu
' Date: 

' Remark:
' 
Option Explicit
' ===================================================================================================
' File content:
' Write comments
' ===================================================================================================

Sub PPTToPdf(filePath As String, Optional newFilePath As String) 
' ==================================================================================================
' ARGUMENTS:
' filePath	--string
' RETURNS: ***
' action 	--Save pptx as pdf under the same folder
' ==================================================================================================

' Preparation Phrase
dim target as Object 
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

target.ExportAsFixedFormat fileName2, ppSaveAsPDF ' ppFixedFormatTypePDF, ppFixedFormatIntentScreen ' ppSaveAsPDF ' SaveAs


' Checking Phrase
target.Close
objPPT.Quit

Set objPPT=Nothing
Set target=Nothing 
End Sub



Function PPTToPdf_help() as String

PPTToPdf_help="filePath As String, Optional newFilePath As String"

End Function 

