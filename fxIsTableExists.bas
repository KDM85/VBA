Option Explicit

Public Function IsTableExists(ByVal strTable As String) As Boolean

On Error Resume Next

  ' Purpose:	Determine if a table exists
  ' Return:		Boolean
  ' Arguments:	strTable = Table to be checked

  ' Set value to True if table exists
  IsTableExists = IsObject(CurrentDb.TableDefs(strTable))
End Function