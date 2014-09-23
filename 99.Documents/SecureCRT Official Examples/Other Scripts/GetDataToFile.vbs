# $language = "VBScript"
# $interface = "1.0"

' This script demonstrates how to capture line by line output
' from a command sent to a server. It then saves each line of output
' to a file. This script shows how the 'WaitForStrings' command can be
' use to wait for multiple possible outputs.

' Constants used by OpenTextFile()
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Sub Main

  crt.Screen.Synchronous = True

  ' Create an instance of the scripting filesystem runtime so we can
  ' manipulate files.
  '
  Dim fso, file
  Set fso = CreateObject("Scripting.FileSystemObject")

  ' Open a file for writing. The last True parameter causes the file
  ' to be created if it doesn't exist.
  '
  Set file = fso.OpenTextFile("c:\temp\output.txt", ForWriting, True)

  ' Send the initial command then throw out the first linefeed that we
  ' see by waiting for it.
  '
  crt.Screen.Send "a.out" & Chr(10)
  crt.Screen.WaitForString Chr(10)

  ' Create an array of strings to wait for.
  '
  Dim waitStrs
  waitStrs = Array( Chr(10), "linux$" )
  
  Dim row, screenrow, readline, items
  row = 1

  Do
    While True

      ' Wait for the linefeed at the end of each line, or the shell prompt
      ' that indicates we're done.
      '	
      result = crt.Screen.WaitForStrings( waitStrs )

      ' If we saw the prompt, we're done.
      If result = 2 Then
        Exit Do
      End If

      ' The result was 1 (we got a linefeed, indicating that we received 
      ' another line of of output). Fetch current row number of the 
      ' cursor and read the first 20 characters from the screen on that row. 
      ' 
      ' This shows how the 'Get' function can be used to read line-oriented 
      ' output from a command, Subtract 1 from the currentRow to since the 
      ' linefeed moved currentRow down by one.
      ' 
      screenrow = crt.screen.CurrentRow - 1
      readline = crt.Screen.Get(screenrow, 1, screenrow, 20 )

      ' NOTE: We read 20 characters from the screen 'readline' may contain 
      ' trailing whitespace if the data was less than 20 characters wide.

      ' Write the line out with an appended '\r\n'
      file.Write readline & vbCrLf
    Wend
  Loop

  crt.screen.synchronous = false

End Sub

