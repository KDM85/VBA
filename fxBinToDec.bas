Option Explicit

Public Function BinToDec(varInput As Variant) As Integer
  Dim i As Integer
  Dim strInput As String
  Dim intOutput As Integer

  ' Purpose:	Convert binary value to decimal value
  ' Return:		Integer
  ' Arguments:	varInput = binary value to be converted

  ' Counter to cycle through "bits"
  i = 0
  
  ' Ensure varInput is a String
  strInput = CStr(varInput)

  ' Cylce through each character of the String
  While i < Len(strInput)
    ' Find 1s in the String
    If (Mid(strInput, Len(strInput) - i, 1) = 1) Then
      ' Build the output value
      intOutput = intOutput + (2 ^ i)
    End If
    i = i + 1
  Wend

  ' Return the decimal value
  BinToDec = intOutput
End Function