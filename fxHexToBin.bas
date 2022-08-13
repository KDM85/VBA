Option Explicit

Public Function HexToBin(varInput As Variant) As String
  Dim i As Integer
  Dim strInput As String
  Dim strOutput As String

  ' Purpose:	Convert hexadecimal value to binary value
  ' Return:		String
  ' Arguments:	varInput = hexadecimal value to be converted

  ' Counter to cycle through input value
  i = 1
  
  ' Ensure varInput is a String
  strInput = CStr(varInput)

  ' Build the binary value
  While i <= Len(varInput)
    Select Case Mid(strInput, i, 1)
      Case "0"
        strOutput = strOutput & "0000"
      Case "1"
        strOutput = strOutput & "0001"
      Case "2"
        strOutput = strOutput & "0010"
      Case "3"
        strOutput = strOutput & "0011"
      Case "4"
        strOutput = strOutput & "0100"
      Case "5"
        strOutput = strOutput & "0101"
      Case "6"
        strOutput = strOutput & "0110"
      Case "7"
        strOutput = strOutput & "0111"
      Case "8"
        strOutput = strOutput & "1000"
      Case "9"
        strOutput = strOutput & "1001"
      Case "A"
        strOutput = strOutput & "1010"
      Case "B"
        strOutput = strOutput & "1011"
      Case "C"
        strOutput = strOutput & "1100"
      Case "D"
        strOutput = strOutput & "1101"
      Case "E"
        strOutput = strOutput & "1110"
      Case "F"
        strOutput = strOutput & "1111"
    End Select
    i = i + 1
  Wend

  ' Return the binary value
  HexToBin = strOutput
End Function