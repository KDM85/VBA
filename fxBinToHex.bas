Option Explicit

Public Function BinToHex(varInput As Variant) As String
  Dim i As Integer
  Dim strInput As String
  Dim strOutput As String

  ' Purpose:	Convert binary value to hexadecimal value
  ' Return:		String
  ' Arguments:	varInput = binary value to be converted

  ' Counter to cycle through "bytes"
  i = 1
  
  ' Ensure varInput is a String
  strInput = CStr(varInput)

  ' Add leading zeros to build four-bit groups
  Select Case (Len(strInput) Mod 4)
    Case 1
      strInput = "000" & strInput
    Case 2
      strInput = "00" & strInput
    Case 3
      strInput = "0" & strInput
  End Select

  ' Convert each character to its binary value
  While i <= Len(strInput)
    Select Case Mid(strInput, i, 4)
      Case "0000"
        strOutput = strOutput & "0"
      Case "0001"
        strOutput = strOutput & "1"
      Case "0010"
        strOutput = strOutput & "2"
      Case "0011"
        strOutput = strOutput & "3"
      Case "0100"
        strOutput = strOutput & "4"
      Case "0101"
        strOutput = strOutput & "5"
      Case "0110"
        strOutput = strOutput & "6"
      Case "0111"
        strOutput = strOutput & "7"
      Case "1000"
        strOutput = strOutput & "8"
      Case "1001"
        strOutput = strOutput & "9"
      Case "1010"
        strOutput = strOutput & "A"
      Case "1011"
        strOutput = strOutput & "B"
      Case "1100"
        strOutput = strOutput & "C"
      Case "1101"
        strOutput = strOutput & "D"
      Case "1110"
        strOutput = strOutput & "E"
      Case "1111"
        strOutput = strOutput & "F"
    End Select
    i = i + 4
  Wend

  ' Return the hexadecimal value
  BinToHex = strOutput
End Function