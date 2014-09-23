#$language = "VBScript"
#$interface = "1.0"

'Scirpt2.vbs can be excuted by itself'

Function script2_func ()
    msgbox "[script2][func] this is script2-specific func"
End Function

Function common_func ()
    msgbox "[script2][common_func] this is common func"
End Function

param = 200

Sub Main ()
    ' notice: reset the run_return value at the beginning of script'
    run_return = False

    msgbox "[script2] enter"
    msgbox "[script2] param = " & param
    common_func
    script2_func

    If run_return = False Then
        msgbox "[script2] break"
        Exit Sub
    End If

    msgbox "[script2] exit"
End Sub
