# $language = "VBScript"
# $interface = "1.0"

' This script sets a specific logfile, enables logging then connects
' to a server. After capturing the output of a command to the logfile
' logging is disabled and it disconnects.

Sub Main

  ' Turn on synchronous mode while performing Send/Wait sequences
  ' so no input is missed.
  '
  crt.Screen.Synchronous = True

  ' Connect using a pre-defined session that automatically logs me in. 
  '
  crt.Session.Connect "/s mysession"

  ' Wait for my unix login prompt or for 5 seconds whichever
  ' comes first.
  '
  crt.Screen.WaitForString "linux$", 5

  ' Set the name of the log file name "YYMMDD.log"
  '
  Dim logfile
  logfile = "C:\TEMP\mysession.log"
  crt.Session.LogFileName = logfile

  ' Enable logging
  '
  crt.Session.Log True

  ' Send a unix command. The output of the command will
  ' be captured to the logfile.
  '
  crt.Screen.Send "date" & vbCr

  ' Wait again for my login prompt or 5 seconds
  '
  crt.Screen.WaitForString "linux$", 5

  ' Turn off synchronous mode
  crt.Screen.Synchronous = false

  ' Stop logging and disconnect.
  '
  crt.Session.Log False
  crt.Session.Disconnect

End Sub

