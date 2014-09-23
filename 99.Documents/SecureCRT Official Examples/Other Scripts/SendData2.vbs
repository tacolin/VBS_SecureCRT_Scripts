# $language = "VBScript"
# $interface = "1.0"

' This script shows how to read in a file, and it demonstrates how to
' perform some preprocessing on data (splitting the file data into 
' separate strings) before sending it to a server.

' Constant used by OpenTextFile()
Const ForReading = 1

Sub main

  ' Open a file, read it in & send it one line at a time
  Dim fso, f

  Set fso = CreateObject("Scripting.FileSystemObject")

  Set f = fso.OpenTextFile("c:\temp\printers.txt", ForReading, 0)

  Dim line, params 
  Do While f.AtEndOfStream <> True

    ' Read each line of the printers file.
    '
    line = f.Readline

    ' Split the line up. Each line should contain 3 space-separated parameters
    params = Split( line )

    ' params(0) holds parameter 1, params(1) holds parameter 2, etc.
    '
    ' Send "mycommand" with the appended parameters from the file with
    ' an appended CR.
    '
    crt.Screen.Send "mycommand" & params(0) & " " & params(1) & " " & params(2) & " " & vbCR 

    ' Cause a 3-second pause between sends by waiting for something "unexpected"
    ' with a timeout value.
    crt.Screen.WaitForString "something_unexpected", 3
  Loop

End Sub
