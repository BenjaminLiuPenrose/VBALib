' Name: Beier (Benjamin) Liu
' Date: 5/21/2018

' Remark:
' A.D.O reference required
Option Explicit
' ===================================================================================================
' File content:
' Append Excel data to Access
' ===================================================================================================

Sub ExcelToAccess (filePath As String, startRow As Integer)
' ======================================================================================================
' ARGUMENTS:
' filePath		--string, the path for the Access database
' startRow		--integer, number of first row to be inserted
' RETURNS:	
' Insert Excel data to database 
'======================================================================================================

' Preparation Phrase
Dim wkSheet As Worksheet
Set wkSheet = activeworkbook.activesheet

Dim rowCount As Integer
Dim i As Integer
rowCount=wkSheet.Cells(Rows.Count, "A").End(xlUp).row 			' can be changed

Dim cn As Object
Dim rs As Object 
Dim strFile As String
Dim strCon As String
strFile = filePath
strCon="Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strFile

Set cn = CreateObject("ADODB.Connection")
cn.Open strCon
Set rs = CreateObject("ADODB.RECORDSET")

' Handling Phrase
rs.Open "SELECT * FROM Test", cn, adOpenKeyset, adLockOptimistic ' can be changed

rs.AddNew

For i=startRow To rowCount												' can be changed
	rs.Fields("No_") = wkSheet.Range("A" & i)
	rs.Fields("Date_") = wkSheet.Range("B" & i)
	rs.Fields("Stock_Name") = wkSheet.Range("C" & i)
	rs.Fields("Notional") = wkSheet.Range("D" & i)
	rs.Fields("Tenor") = wkSheet.Range("E" & i)
	rs.Fields("Structure") = wkSheet.Range("F" & i)
	rs.Fields("Underlying") = wkSheet.Range("G" & i)
	rs.Fields("Strike1") = wkSheet.Range("H" & i)
	rs.Fields("Strike2") = wkSheet.Range("I" & i)
	rs.Fields("Comment") = wkSheet.Range("J" & i)
	rs.Fields("Margin") = wkSheet.Range("K" & i)
	rs.Fields("EDS_Quote") = wkSheet.Range("L" & i)
	rs.Fields("Fair_Value") = wkSheet.Range("M" & i)
	rs.Fields("Offer_Price") = wkSheet.Range("N" & i)
Next i

rs.Update

' Checking Phrase
rs.Close
cn.Close
Set rs = Nothing
Set cn = Nothing
End Sub



Function ExcelToAccess_help() As String

AccessToExcel_help = "filePath As String, startRow As Integer"

End Function