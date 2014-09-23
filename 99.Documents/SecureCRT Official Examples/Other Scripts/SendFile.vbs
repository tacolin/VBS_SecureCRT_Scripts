# $language = "VBScript"
# $interface = "1.0"

' This script demonstrates one way to send a text file to a unix server
' This is equivalent to doing an "send ASCII" in a script.

' Constants used by OpenTextFile()
Const ForReading = 1
Const ForWriting = 2

Sub Main

  ' Create an instance of the filesystem object so we can open a file
  ' Note: A runtime exception will be generated if 'myfile.txt' doesn't exist.
  '
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set file = fso.OpenTextFile("c:\temp\myfile.txt", ForReading, False)

  ' Send a 'cat' command to direct everything we are about to send into a file.
  '
  crt.Screen.Send "cat > output.txt" & vbCR

  ' The next line causes the script to wait until the 'cat' command has been 
  ' sent (and executed) by the server. So that the data we're about to send
  ' doesn't get sent until redirected 'cat' command is actually in effect.
  '
  crt.Screen.WaitForString "output.txt"

  ' Send the file a line at a time.
  '
  Do While file.AtEndOfStream <> True
    str = file.Readline
    crt.Screen.Send str & Chr(13)
  Loop

  ' send an EOF character to end output to terminate 'cat'
  '
  crt.Screen.Send Chr(4)
 
End Sub

