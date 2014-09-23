# $language = "VBScript"
# $interface = "1.0"

' A VBScript login script that waits for a login prompt then sends
' a 'setenv DISPLAY <ipaddress>:0.0' command to direct X-clients to a
' locally run Xserver.

Sub Main

  crt.Screen.WaitForString("myhost$ ")

  Dim display
  display = "setenv DISPLAY " & crt.Session.LocalAddress & ":0.0" & vbCr
  crt.Screen.Send(display)

End Sub
