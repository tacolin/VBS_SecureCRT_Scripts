# $language = "VBScript"
# $interface = "1.0"

' Run.vbs demonstrates how to utilize the Windows Scripting Host (WSH) by using 
' its 'Run' method to execute other programs. Note the use of nested quotes to pass
' a path that contains spaces along with command line arguments.

Sub Main
	
  Dim shell
  Set shell = CreateObject("WScript.Shell")
  shell.Run """C:\Program Files\Internet Explorer\IExplore.exe"" http://www.vandyke.com"

End Sub
