# $language = "VBScript"
# $interface = "1.0"

' Connect to an SSH server using the SSH2 protocol. Specify the
' username and password and hostname on the command line as well as
' some SSH2 protocol specific options.

Sub Main

  Dim host
  host = "ssh.somecompany.com"
  Dim user
  user = "myusername"

  ' Prompt for a password instead of embedding it in a script...
  '
  Dim passwd
  passwd = crt.Dialog.Prompt("Enter password for " & host, "Login", "", True)

  ' Build a command-line string to pass to the Connect method.
  '
  cmd = "/SSH2 /L " & user & " /PASSWORD " & passwd & " /C 3DES /M MD5 " & host

  crt.Session.Connect cmd

End Sub
