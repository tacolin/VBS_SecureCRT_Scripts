#$language = "VBScript"
#$interface = "1.0"

Function info (text, title)
    buttons = vbOKOnly
    icon = vbInformation
    msgbox text, buttons + icon, title
End Function

Function warning (text, title)
    buttons = vbOKOnly
    icon = vbExclamation
    msgbox text, buttons + icon, title
End Function

Function error (text, title)
    buttons = vbOKonly
    icon = vbCritical
    msgbox text, buttons + icon, title
End Function

Function confirm (text, title)
    buttons = vbYesNo
    icon = vbQuestion
    If msgbox(text, buttons + icon, title) = vbYes Then
        confirm = true
    Else
        confirm = false
    End If
End Function

Sub Main ()

    info "this is information.", "title of info"

    warning "this is warning!", "title of warning"

    error "this is error!!!", "title of error"

    answer = confirm("do you want to say goodbye?", "title of confirm")
    If answer Then
        info "goodbye", "sayorama"
    Else
        info "no goodbye", "madadesu"
    End If

End Sub
