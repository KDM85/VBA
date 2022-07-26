Option Explicit

Public Function ConvertToTime(intSec As Variant) As String
  Dim intMinutes As Integer
  Dim intSeconds As Integer

  ' Purpose:	Convert total seconds into minutes and seconds
  ' Return:		String
  ' Arguments:	intSec = Total seconds to be converted
  ' Notes:		Output values will be formatted as "0:00" or "00:00"
    
  ' Set null values or non-integer values to zero
  If Len(intSec) = 0 Or IsNull(intSec) Then
    ConvertToTime = 0
    Exit Function
  End If
  
  ' Convert seconds to minutes and seconds
  intMinutes = Int(intSec / 60)
  intSeconds = intSec - (intMinutes * 60)

  ' Concatenate seconds and minutes in text time format
  ConvertToTime = intMinutes & ":" & Format(intSeconds, "00")
End Function