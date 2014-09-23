#$language = "VBScript"
#$interface = "1.0"

Class IeLogWindow
    ' ie object'
    Private ie
    Public  maxLogLineNum

    ' when you new IeLogWindow, constructor will be excuted.'
    Private Sub Class_Initialize
        maxLogLineNum = 10000

        set ie = CreateObject("InternetExplorer.Application")
        ie.Navigate "about:blank"

        Do
            crt.sleep 100
        Loop While ie.Busy

        ie.menubar = False
        ie.statusbar = False
        ie.toolbar = False
        ie.document.title = "Logging"
        ie.width = 800
        ie.height = 540

        ie.document.body.innerHTML = _
        "<h1 id='LogTitle'>Logging</h1>" & _
        "<div style='width:100%; height:85%;'>" & _
        "    <textarea READONLY id='LogContent' wrap='off'" & _
        "    " & _
        "    style='width:100%; height:100%;" & _
        "    font-famliy:consolas; font-size:18; " & _
        "    line-height:1.5; '></textarea>" & _
        "</div>"

        ie.visible = True
        ie.document.parentWindow.focus
    End Sub

    ' when you set ieobj to nothing, desctructor will be excuted.'
    Private Sub Class_Terminate
        ' ie.quit
    End Sub

    Public Function quit ()
        ie.quit
    End Function

    Public Function append (text)
        dim textarea : set textarea = ie.Document.All("LogContent")
        currentText = ie.Document.All("LogContent").Value

        ' when line number > maxLogLinNum, delete the first line'
        ' but why vbLf? i don't know
        ' we append text and add vbcr in each line end'
        ' but split lines by vblf ... '

        lines = uBound(Split(currentText, vbLf)) + 1
        If lines > maxLogLineNum Then
            newStart = InStr(currentText, vbLf) + 1
            currentText = Mid(currentText, newStart)
        End If
        textarea.Value = currentText & text & vbCr
        textarea.scrollTop = textarea.scrollHeight
    End Function

End Class


Sub Main ()

    dim logwindow : set logwindow = new IeLogWindow

    logwindow.append "logwindow log window created."
    logwindow.append "line 1"
    logwindow.append "line 2"
    logwindow.append "line 3"
    logwindow.append "hello world"

    If msgbox("do you want to close this window?", vbYesNo + vbQuestion, "Quit?") = vbyes Then
        logwindow.quit
    Else
        logwindow.append "you should close it by yourself."
    End If

End Sub
