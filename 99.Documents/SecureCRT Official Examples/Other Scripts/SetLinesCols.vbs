# $language = "VBScript"
# $interface = "1.0"

' Use CRT's script object Rows and Columns properties to send settings
' for the LINES and COLUMNS environment variables when these variables
' aren't being set properly by the remote system.

Sub Main
  lines = crt.Screen.Rows
  cols = crt.Screen.Columns

  ' send bourne shell/korn shell command
  ' crt.Screen.Send "export LINES=" & lines & " COLUMNS=" & cols & Chr(13)

  ' send csh shell command
  '
  crt.Screen.Send "setenv LINES " & lines & "; setenv COLUMNS " & cols & Chr(13)

End Sub
