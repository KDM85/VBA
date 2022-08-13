Option Explicit

Sub Wait(lngSeconds As Long)

  ' Purpose:	Force a delay
  ' Return:		N/A
  ' Arguments:	lngSeconds = Amount of seconds to delay

  Dim lngSec As Long
  lngSec = Timer + lngSeconds
  
  Do While Timer < lngSec
    DoEvents
  Loop
End Sub