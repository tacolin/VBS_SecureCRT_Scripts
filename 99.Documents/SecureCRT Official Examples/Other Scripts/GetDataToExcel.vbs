# $language = "VBScript"
# $interface = "1.0"

' This script demonstrates how ActiveX scripting can be used to
' interact with CRT and manipulate other programs such as Microsoft Excel
' through an OLE automation interface. This script creates an instance of Excel, 
' then it sends a command to a remote server (assuming we're already 
' connected). It reads the output, parses it and writes out some of the
' data to an Excel spreadsheet and saves it. This script also demonstrates
' how the WaitForStrings function can be used to wait for more than one
' output string.

Sub main

  crt.screen.synchronous = true

  ' Create an Excel workbook/worksheet
  '
  Dim app, wb, ws
  Set app = CreateObject("Excel.Application")
  Set wb = app.Workbooks.Add
  Set ws = wb.Worksheets(1)

  ' Send the initial command to run and wait for the first linefeed
  '
  crt.Screen.Send("cat /etc/passwd" & Chr(10) )
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

      ' We saw the prompt, we're done.
      '
      If result = 2 Then
        Exit Do
      End If

      ' Fetch current row and read the first 40 characters from the screen
      ' on that row. Note, since we read a linefeed character subtract 1 
      ' from the return value of CurrentRow to read the actual line.
      '
      screenrow = crt.screen.CurrentRow -1
      readline = crt.Screen.Get(screenrow, 1, screenrow, 40 )

      ' Split the line ( ":" delimited) and put some fields into Excel
      '
      items = Split( readline, ":", -1 )

      ws.Cells(row, 1).Value = items(0)
      ws.Cells(row, 2).Value = items(2)
      row = row + 1
    Wend
  Loop

  wb.SaveAs("C:\temp\chart.xls")
  wb.Close
  app.Quit

  Set ws = nothing
  Set wb = nothing
  Set app = nothing

  crt.screen.synchronous = false

End Sub