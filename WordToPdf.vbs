' Name: Beier (Benjamin) Liu
' Date: 5/21/2018

' Remark:
' Word Object required
Option Explicit
' ===================================================================================================
' File content:
' Write comments
' ===================================================================================================

Sub WordToPdf(filePath As String, Optional newFilePath As Variant) 
' ==================================================================================================
' ARGUMENTS:
' filePath	--string, path of the word file 
' newFilePath--string, path of the pdf file 
' RETURNS: 
' action 	--Save docx as pdf (if obmitted, under the same folder)
' ==================================================================================================

' Preparation Phrase
dim target as Object
dim objWord as Object

Application.ScreenUpdating=False
Application.EnableEvents=False

On Error Resume Next 
Set objWord = GetObject(, "Word.Application")
If objWord Is Nothing Then
Set objWord = CreateObject("Words.Application")
End If
On Error GoTo 0

' Handling Phrase
Set target=objWord.Documents.Open(filePath, ReadOnly:=False)
' target.UpdateLinks

dim fileName2 As String

If IsMissing(newFilePath) Then
	fileName2=newFilePath
Else 
	fileName2=Replace(target, "docx", "pdf")
End If

target.ExportAsFixedFormat fileName2, wdExportFormatPDF ' SaveAs


' Checking Phrase
target.Close
objWord.Quit

Application.ScreenUpdating=True
Application.EnableEvents=True

Set objWord=Nothing
Set target=Nothing 
End Sub



Function WordToPdf_help() as String

WordToPdf_help="filePath As String, Optional newFilePath As Variant"

End Function 

