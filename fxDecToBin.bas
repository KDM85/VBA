Option Explicit

Public Function DecToBin(varInput As Variant) As String
  Dim strOutput As String
  Dim intRemainder As Integer
  Dim intQuotient As Integer

  ' Purpose:	Convert decimal value to binary value
  ' Return:		String
  ' Arguments:	varInput = decimal value to be converted
  
  ' Ensure varInput is an Integer
  intQuotient = Int(varInput)
  
  ' Build binary value
  While Not intQuotient = 0
    intRemainder = intQuotient Mod 2
    intQuotient = Int(intQuotient / 2)
    strOutput = IIf(intRemainder = 0, "0", "1") & strOutput
  Wend
  
  ' Return binary value
  DecToBin = strOutput
End Function