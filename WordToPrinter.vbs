' Name: Beier (Benjamin) Liu
' Date: 5/21/2018

' Remark:
' 
Option Explicit
' ===================================================================================================
' File content:
' Print out word documents
' ===================================================================================================

Sub WordToPrinter(filePath As String, numCopies As Integer, collate As Bollean) 
' ==================================================================================================
' ARGUMENTS:
' filePath	--string, the path of the word file 
' numCopies	--integer, number of copies
' collate 	--bollean, true means collate
' RETURNS:
' action 	--print out word document
' ==================================================================================================

' Preparation Phrase
Dim objWord As Object
Dim target As Object

Application.ScreenUpdating=False
Application.EnableEvents=False

On Error Resume Next 
Set objWord = GetObject(, "Word.Application")
If objPPT Is Nothing Then
Set objWord = CreateObject("Word.Application")
End If
On Error GoTo 0

' Handling Phrase
Set target=objPPT.Presentations.Open(filePath)
' target.UpdateLinks


target.PrintOut Copies:=numCopies, Collate:=collate, ' ActivePrinter:=""

' Checking Phrase
target.Close
objPPT.Quit

Application.ScreenUpdating=True
Application.EnableEvents=True

Set objPPT=Nothing
Set target=Nothing
End Sub

Function WordToPrinter_help() as String

PPTToPrinter_help="filePath As String, numCopies As Integer, collate As Bollean"

End Function 

