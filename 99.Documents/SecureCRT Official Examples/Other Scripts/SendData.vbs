# $language = "VBScript"
# $interface = "1.0"

' This script demonstrates how to open a text file and it line by
' line to a server.

' Constants used by OpenTextFile()
'
Const ForReading = 1
Const ForWriting = 2

Sub Main

  Dim fso, file, str
  Set fso = CreateObject("Scripting.FileSystemObject")

  ' Note: A runtime exception will be generated if 'input.txt' doesn't exist.
  '
  Set file = fso.OpenTextFile("c:\temp\input.txt", ForReading, False)

  crt.Screen.Synchronous = True

  Do While file.AtEndOfStream <> True

    str = file.Readline

    ' Send the line with an appended CR
    '
    crt.Screen.Send str & Chr(13)
 
    ' Wait for my prompt before sending the next line
    '
    crt.Screen.WaitForString "prompt$"
  Loop

  crt.Screen.Synchronous = False

End Sub