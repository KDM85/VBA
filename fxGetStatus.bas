Option Explicit

Public Function getStatus(intNumerator As Integer, intDivisor As Integer, strStatus As String)
  Dim intStatus As Integer
  Dim intWidth As Integer

  ' Purpose:	Display status messages in a Status form in an Access Database
  ' Return:		N/A
  ' Arguments:	intNumerator = The current step number of the total number of steps
  '		intDivisor = The total number of steps
  '		strStatus = The status message to be displayed
  ' Notes:		Requires a form called "frmStatus" which has the following items:
  '		lblBar = Label set to a solid back of the desired color and a width of 0
  '		lblStatus = Label to display the status message

  intStatus = Forms!frmStatus.boxStatus.Width  ' Width of status bar in TWIP (1' = 1440 TWIP)
  
  ' Avoid division by zero
  If intDivisor = 0 Then
    intDivisor = 1
  End If

  ' Set the width of the bar to be the ratio of completion
  intWidth = (intNumerator / intDivisor) * intStatus
  
  ' Set the width of the status bar
  Forms!frmStatus.lblBar.Width = intWidth
  ' Set the caption of the status label
  Forms!frmStatus.lblStatus.Caption = strStatus
  ' Repaint the form so that the changes are displayed
  Forms!frmStatus.Repaint
End Function