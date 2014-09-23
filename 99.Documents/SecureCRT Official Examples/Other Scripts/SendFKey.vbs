# $language = "VBScript"
# $interface = "1.0"

' This script simulates an F3 keypress (VT100 keyboard)

Sub Main
   
    ' Simulate F3 by sending ESC-OR
    '
    crt.screen.Send Chr(27) & "OR"

End Sub

