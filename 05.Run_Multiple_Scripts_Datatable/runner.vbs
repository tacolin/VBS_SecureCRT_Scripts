'----------------------------------------------------------------------'
' return value of executing each script'
'----------------------------------------------------------------------'
dim return_value : return_value = false

Class runner

    Private Sub Class_Initialize
        ' Called automatically when class is created
    End Sub

    Private Sub Class_Terminate
        ' Called automatically when all references to class instance are removed
    End Sub

    '----------------------------------------------------------------------'
    ' Read the script text'
    '----------------------------------------------------------------------'
    Private Function m_getVbsText (vbsFile)
        dim filename : filename = Split(vbsFile, ".")(0)
        dim fso, file, text
        set fso = CreateObject("Scripting.FileSystemObject")

        set file = fso.OpenTextFile(vbsFile)
        text = file.ReadAll()
        file.Close

        m_getVbsText = text
    End Function

    '----------------------------------------------------------------------'
    ' Only catch text from Sub Main () to End Sub '
    ' Replace the Sub Main () to Sub 'RunnerEngine' ()
    '----------------------------------------------------------------------'
    Private Function m_catchMainSubTextOnly (origin_text)
        lines = Split(origin_text, vbCrLf)
        startline = -1
        For i=0 To Ubound(lines)
            If InStr(LCase(lines(i)), "sub") > 0 and _
            InStr(LCase(lines(i)), "main") > 0 Then
                lines(i) = "Sub RunnerEngine" & " ()"
                startline = i
            End If
        Next

        text = ""
        For i=startline To Ubound(lines)
            text = text & lines(i) & vbCrLf
        Next
        text = text & vbCrLf
        text = text & "RunnerEngine"

        m_catchMainSubTextOnly = text
    End Function

    '----------------------------------------------------------------------'
    ' Remove the un-necessary part :'
    ' 1. #$language = "VBScript" '
    ' 2. #$interface = "1.0"'
    ' '
    ' Replace the Sub Main () to Sub 'RunnerEngine' ()
    '----------------------------------------------------------------------'
    Private Function m_removeDeclarationTextOnly (origin_text)
        lines = Split(origin_text, vbCrLf)
        startline = 0
        For i=0 To Ubound(lines)
            If InStr(LCase(lines(i)), "language") > 0 and _
                InStr(LCase(lines(i)), "vbscript") Then

                lines(i) = ""

            ElseIf InStr(LCase(lines(i)), "interface") > 0 and _
                InStr(LCase(lines(i)), "1.0") Then

                lines(i) = ""

            ElseIf InStr(LCase(lines(i)), "sub") > 0 and _
                InStr(LCase(lines(i)), "main") > 0 Then

                lines(i) = "Sub RunnerEngine" & " ()"

            End If
        Next

        text = ""
        For Each line In lines
            text = text & line & vbCrLf
        Next
        text = text & vbCrlf
        text = text & "RunnerEngine"

        m_removeDeclarationTextOnly = text
    End Function

    Private Function m_runVbsText (script_text)
        ExecuteGlobal script_text
        m_runVbsText = return_value
    End Function

    Private Function m_insertVariables (text)
        insertion = "return_value = false" & vbCrLf & vbCrLf
        m_insertVariables = insertion & text
    End Function

    Public Function RunVbsFile (vbsFile)
        text = m_getVbsText(vbsFile)
        text = m_catchMainSubTextOnly(text)
        text = m_insertVariables(text)
        RunVbsFile = m_runVbsText(text)
    End Function

End Class

