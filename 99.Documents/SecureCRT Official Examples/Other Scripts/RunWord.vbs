# $language = "VBScript"
# $interface = "1.0"

' This script uses the GetObject() call to first attempt to connect to
' a running instance of MS Word. If that fails a call to CreateObject()
' attempts to start Word, make it visible and maximizes it.

Sub Main
	
  Dim obj

  ' Request to handle errors ourselves so we can handle possible failure
  ' of GetObject() to connect to Word...
  '
  On Error Resume Next

  ' Gets an instance of Word if it is already running
  '
  Set obj = GetObject(,"Word.Application")
  If TypeName(obj) <> "Application" Then
    ' It wasn't already running so start it running.
    '
    Set obj = CreateObject("Word.Application")
  End If

  obj.Visible = True
  obj.WindowState = 1 ' maximized

End Sub
