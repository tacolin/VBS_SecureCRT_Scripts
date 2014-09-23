# $language = "VBScript"
# $interface = "1.0"

' This VBScript shows how you can use the 'Include' subroutine below to
' let your script load other files and execute functions. This allows 
' common script code to be shared by more than one CRT script.

Include "Module.vbs"

Sub Main
  Doit "Hello"
  Doit "GoodBye"
End Sub

' This subroutine must be pasted into any VBScript that calls 'Include'.
' NOTE: you may need to update your script engines and scripting runtime
' in order to successfully create the 'Scripting.FileSystemObject'.
'
Sub Include(file)
  Dim fso, f
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set f = fso.OpenTextFile(file, 1)
  str = f.ReadAll
  f.Close
  ExecuteGlobal str
End Sub

