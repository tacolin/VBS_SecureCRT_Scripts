#$language = "VBScript"
#$interface = "1.0"

' use this to get the output from script1 and script2 '
dim run_return : run_return = False

Sub Run(vbsFile)
    '-------------------------------------------------------------------------'
    ' Read the script text'
    '-------------------------------------------------------------------------'
    dim filename : filename = Split(vbsFile, ".")(0)
    dim fso, f, s
    set fso = CreateObject("Scripting.FileSystemObject")
    set f = fso.OpenTextFile(vbsFile)
    s = f.ReadAll()
    f.Close

    '-------------------------------------------------------------------------'
    ' Remove the un-necessary part :'
    ' 1. #$language = "VBScript" '
    ' 2. #$interface = "1.0"'
    ' '
    ' Replace the Sub Main () to Sub 'filename' ()
    '-------------------------------------------------------------------------'
    lines = Split(s, vbCrLf)
    dim beginLine, endLine
    For i=0 To Ubound(lines)
        If InStr(LCase(lines(i)), "language") > 0 and InStr(LCase(lines(i)), "vbscript") Then
            lines(i) = ""
        ElseIf InStr(LCase(lines(i)), "interface") > 0 and InStr(LCase(lines(i)), "1.0") Then
            lines(i) = ""
        ElseIf InStr(LCase(lines(i)), "sub") > 0 and InStr(LCase(lines(i)), "main") > 0 Then
            lines(i) = "Sub " & filename & " ()"
        End If
    Next

    '-------------------------------------------------------------------------'
    ' Re-assemble the text '
    ' Append the execution of 'filename' () in the last line
    '-------------------------------------------------------------------------'
    s = ""
    For Each line In lines
        s = s & line & vbCrLf
    Next
    s = s & vbCrLf
    s = s & " " & filename

    '-------------------------------------------------------------------------'
    ' Execute the script '
    '-------------------------------------------------------------------------'
    ExecuteGlobal s
End Sub

Sub Main ()

    scripts = Array("script1.vbs", "script2.vbs")

    For Each script In scripts
        Run script
        msgbox script & " return value = " & run_return
    Next

End Sub
