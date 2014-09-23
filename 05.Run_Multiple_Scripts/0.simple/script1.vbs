#$language = "VBScript"
#$interface = "1.0"

'Scirpt1.vbs can be excuted by itself'

param = 100

Function common_func ()
    msgbox "[script1][common_func] this is common func"
End Function

Sub Main ()
    ' notice: reset the run_return value at the beginning of script'
    run_return = False

    msgbox "[script1] enter"
    msgbox "[script1] param = " & param
    common_func

    run_return = True

    msgbox "[script1] exit"
End Sub
