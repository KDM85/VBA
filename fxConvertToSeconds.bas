Option Explicit

Public Function ConvertToSeconds(strNum As Variant) As Integer
  Dim intMinutes As Integer
  Dim intSeconds As Integer

  ' Purpose:	Convert minutes and seconds into total seconds
  ' Return:		Integer
  ' Arguments:	strNum = Time to be converted to seconds
  ' Notes:		Values should be entered as "0:00" or "00:00"

  ' Set null values or non-integer values to zero
  If Len(strNum) = 0 Or IsNull(strNum) Then
    ConvertToSeconds = 0
    Exit Function
  End If
  
  ' Convert text to minutes and seconds
  intMinutes = CInt(Left(strNum, InStr(1, strNum, ":", vbTextCompare) - 1))
  intSeconds = CInt(Mid(strNum, InStr(1, strNum, ":", vbTextCompare) + 1))

  ' Convert minutes and seconds into total seconds
  ConvertToSeconds = intMinutes * 60 + intSeconds
  
End Function