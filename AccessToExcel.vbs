' Name: Beier (Benjamin) Liu
' Date: 5/14/2018

' Remark:
'
Option Explicit
' ===================================================================================================
' File content:
' Access to Excel 
' ===================================================================================================

Sub AccessToExcel (filePath As String, strSQL As String, outputCell As String)
' ======================================================================================================
' ARGUMENTS:
' filePath		--string, the path for the Access database
' strSQL		--string, the SQL sentense
' outputCell	--string, the start cell of output 
' RETURNS:	
' Sheets("Res").range(outputCell).CopyFromRecordset	-- Qurey a database with strSQL
'======================================================================================================

' Preparation Phrase
Dim wkSheet As Worksheet
Set wkSheet = activeworkbook.activesheet

Dim strFile As String
Dim strCon As String
Dim cn, rs As Object
strFile = filePath
strCon = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strFile

Set cn = CreateObject("ADODB.Connection")
cn.Open strCon
Set rs = CreateObject("ADODB.RECORDSET")
rs.activeconnection = cn

' Handling Phrase
rs.Open strSQL
wkSheet.Range(outputCell).CopyFromRecordset rs

' Checking Phrase
rs.Close
cn.Close
Set cn = Nothing
End Sub



Function AccessToExcel_help() As String

AccessToExcel_help = "path As String, strSQL As String, outputCell As String"

End Function