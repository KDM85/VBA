Option Explicit

Public Function DecToHex(varInput As Variant) As String
  Dim strOutput As String
  Dim intRemainder As Integer
  Dim intQuotient As Integer

  ' Purpose:	Convert decimal value to hexadecimal value
  ' Return:		String
  ' Arguments:	varInput = decimal value to be converted
  
  ' Ensure varInput is an Integer
  intQuotient = Int(varInput)
  
  ' Build hexadecimal value
  While Not intQuotient = 0
    intRemainder = intQuotient Mod 16
    intQuotient = Int(intQuotient / 16)
    
    Select Case intRemainder
      Case 0
        strOutput = "0" & strOutput
      Case 1
        strOutput = "1" & strOutput
      Case 2
        strOutput = "2" & strOutput
      Case 3
        strOutput = "3" & strOutput
      Case 4
        strOutput = "4" & strOutput
      Case 5
        strOutput = "5" & strOutput
      Case 6
        strOutput = "6" & strOutput
      Case 7
        strOutput = "7" & strOutput
      Case 8
        strOutput = "8" & strOutput
      Case 9
        strOutput = "9" & strOutput
      Case 10
        strOutput = "A" & strOutput
      Case 11
        strOutput = "B" & strOutput
      Case 12
        strOutput = "C" & strOutput
      Case 13
        strOutput = "D" & strOutput
      Case 14
        strOutput = "E" & strOutput
      Case 15
        strOutput = "F" & strOutput
    End Select
  Wend
  
  ' Return hexadecimal value
  DecToHex = strOutput
End Function