# $language = "VBScript"
# $interface = "1.0"

' VBScript example that utilizes the 'Environment' method of the Windows
' Scripting Host (WSH) to read various settings from the Windows environment. 
' Note, WSH may need to be installed before this script will run successfully.

Sub Main

  Set shell = CreateObject("WScript.Shell")
  Set env = shell.Environment("process")
  MsgBox env("COMPUTERNAME")
  MsgBox env("USERNAME")

End Sub
