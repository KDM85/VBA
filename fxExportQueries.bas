Option Explicit

Public Function ExportQueries(strPath As String)
  Dim db As Database
  Dim qdf As QueryDef
  Dim fh As Long

  ' Purpose:	Export SQL Queries from Access Database
  ' Return:		N/A
  ' Arguments:	strPath = File path for export location
  
  Set db = CurrentDb
  
  ' Look at each query in the database
  For Each qdf In db.QueryDefs
    ' Find an unused file number as a long integer
    fh = FreeFile
    ' Open [path\filename] For [mode] As [Freefile]
    Open strPath & "\" & qdf.Name & ".sql" For Append As fh
    ' Write the SQL to the file
    Print #fh, qdf.SQL
    ' Close the file
    Close fh
  Next
End Function