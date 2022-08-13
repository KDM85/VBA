Option Explicit

Public Function HexToDec(varInput As Variant) As Integer
  Dim i As Integer
  Dim strInput As String
  Dim intOutput As Integer

  ' Purpose:	Convert hexadecimal value to decimal value
  ' Return:		Integer
  ' Arguments:	varInput = hexadecimal value to be converted

  ' Counter to cycle through values
  i = 0
  
  ' Ensure varInput is a String
  strInput = CStr(varInput)

  ' Build the decimal value
  While i < Len(strInput)
    Select Case Mid(strInput, Len(strInput) - i, 1)
      Case "0"
        intOutput = intOutput + 0 * (16 ^ i)
      Case "1"
        intOutput = intOutput + 1 * (16 ^ i)
      Case "2"
        intOutput = intOutput + 2 * (16 ^ i)
      Case "3"
        intOutput = intOutput + 3 * (16 ^ i)
      Case "4"
        intOutput = intOutput + 4 * (16 ^ i)
      Case "5"
        intOutput = intOutput + 5 * (16 ^ i)
      Case "6"
        intOutput = intOutput + 6 * (16 ^ i)
      Case "7"
        intOutput = intOutput + 7 * (16 ^ i)
      Case "8"
        intOutput = intOutput + 8 * (16 ^ i)
      Case "9"
        intOutput = intOutput + 9 * (16 ^ i)
      Case "A"
        intOutput = intOutput + 10 * (16 ^ i)
      Case "B"
        intOutput = intOutput + 11 * (16 ^ i)
      Case "C"
        intOutput = intOutput + 12 * (16 ^ i)
      Case "D"
        intOutput = intOutput + 13 * (16 ^ i)
      Case "E"
        intOutput = intOutput + 14 * (16 ^ i)
      Case "F"
        intOutput = intOutput + 15 * (16 ^ i)
    End Select
    i = i + 1
  Wend

  ' Return the decimal value
  HexToDec = intOutput
End Function