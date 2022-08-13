Option Explicit

Public Function Initiate(strTable As String, strSource As String)
  Dim strSQL As String
  Dim strAlter As String
  
  On Error GoTo ErrorHandler

  ' Purpose:	MS Access does not support the use of CREATE TABLE.
  '		This function creates a table from a specified source.
  ' Return:		N/A
  ' Arguments:	strTable = Table to be created
  '		strSource = Source to be used to create the table
  ' Notes:		Source should typically be a SQL query
  
  ' Delete the table if it already exists
  DoCmd.DeleteObject acTable, strTable

  ' Create the new table
  strSQL = "SELECT * INTO " & strTable & " FROM " & strSource
  CurrentDb.Execute strSQL
  Exit Function
  
ErrorHandler:
  ' Ignore if trying to drop a table that does not exist
  If Err.Number = 7874 Or Err.Number = 3084 Then
    Resume Next
  End If
  
End Function