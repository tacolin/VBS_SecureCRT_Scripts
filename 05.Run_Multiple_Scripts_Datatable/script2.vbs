#$language = "VBScript"
#$interface = "1.0"

'Scirpt1.vbs can be excuted by itself'

Function warning (text)
    title = "WARNING"
    buttons = vbOKOnly
    icon = vbExclamation
    msgbox text, buttons + icon, title
End Function

Function error (text)
    title = "ERROR"
    buttons = vbOKonly
    icon = vbCritical
    msgbox text, buttons + icon, title
End Function

Function info (text)
    title = "INFORMATION"
    buttons = vbOKOnly
    icon = vbInformation
    msgbox text, buttons + icon, title
End Function

return_value = false

Sub Main ()

    info "script2 : step 1"

    '... do something ...'

    info "script2 : step 2"

    '... do something ...'

    info "script2 : step 3"

    return_value = false

    error "script2 : result = " & return_value

End Sub
