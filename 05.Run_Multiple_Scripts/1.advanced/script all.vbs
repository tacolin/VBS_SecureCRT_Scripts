#$language = "VBScript"
#$interface = "1.0"

'----------------------------------------------------------------------'
' re-define the functions which exist in scirpt1 and script2
'----------------------------------------------------------------------'
Function warning (text)
End Function

Function error (text)
End Function

Function info (text)
End Function

'----------------------------------------------------------------------'
' Include the header file
'----------------------------------------------------------------------'
Function Include(vbsFile)
    dim fso, f, s
    set fso = CreateObject("Scripting.FileSystemObject")
    set f = fso.OpenTextFile(vbsFile)
    s = f.ReadAll()
    f.Close
    ExecuteGlobal s
End Function

Include "runner.vbs"

'----------------------------------------------------------------------'
' main sub
'----------------------------------------------------------------------'
Sub Main ()
    dim taco : set taco = new runner

    result1 = taco.RunVbsFile("script1.vbs")
    result2 = taco.RunVbsFile("script2.vbs")
    result3 = taco.RunVbsFile("script3.vbs")

    msgbox "result1 = " & result1 & vbCrLf & _
            "result2 = " & result2 & vbCrlf & _
            "result3 = " & result3

End Sub
