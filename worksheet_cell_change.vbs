' ===================================================================================================
' Benjamin's VBA programming template file
' You can ignore this file
' ===================================================================================================

' Name: Beier (Benjamin) Liu
' Date: 5/14/2018

' Remark:
' 
Option Explicit
' ===================================================================================================
' File content:
' Write comments
' ===================================================================================================

Private Sub Worksheet_Change(ByVal Target As Range)

' ==================================================================================================
' ARGUMENTS:
'
' RETURNS:
'
' OPERATIONS:
' Create a chart with specific format
' ==================================================================================================

' Preparation Phrase

' Handling Phrase
If Not Intersect(Target, Range("T1")) Is Nothing Then
    If InStr(1, Range("T1"), "Yes") > 0 Then
        formatting 
    End If
End If
' Checking Phrase

End Sub

