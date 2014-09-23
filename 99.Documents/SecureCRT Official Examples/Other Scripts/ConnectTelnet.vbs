# $language = "VBScript"
# $interface = "1.0"

' Connect to a telnet server and automate the initial login sequence.
' Note that synchronous mode is enabled to prevent server output from
' potentially being missed.

Sub Main

  crt.Screen.Synchronous = True

  ' connect to host on port 23 (the default telnet port)
  '
  crt.Session.Connect "/TELNET login.myhost.com 23"

  crt.Screen.WaitForString "ogin:"

  crt.Screen.Send "myusername" & vbCr

  crt.Screen.WaitForString "assword:"

  crt.Screen.Send "mypassword" & vbCr

  crt.Screen.Synchronous = False

End Sub
