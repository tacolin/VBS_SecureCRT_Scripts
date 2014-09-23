# $language = "VBScript"
# $interface = "1.0"

' The Dialog object's MessageBox() function takes 3 parameters listed
' in order:
'
' text   - The message text to be displayed to users. This parameter is
'          mandatory.
'
' title  - A title or caption for the messagebox window. This parameter 
'          is optional and defaults to the application name.
'
' options - A numeric argument that represents a combination of various
'           ICON, BUTTON and DEFBUTTON values (shown below) that influence
'           the appearance and behavior of the messagebox dialog. This
'           parameter is optional and defaults to 0 (no icon, OK button).

' Constants for setting MessageBox options... 
'
Const ICON_STOP = 16		' Critical message; displays STOP icon.
Const ICON_QUESTION = 32	' Warning query; displays '?' icon.
Const ICON_WARN = 48		' Warning message; displays '!' icon.
Const ICON_INFO= 64		' Information message; displays 'i' icon.

Const BUTTON_OK = 0		' OK button only
Const BUTTON_CANCEL = 1		' OK and Cancel buttons
Const BUTTON_ABORTRETRYIGNORE = 2	' Abort, Retry, and Ignore buttons
Const BUTTON_YESNOCANCEL = 3	' Yes, No, and Cancel buttons
Const BUTTON_YESNO = 4		' Yes and No buttons
Const BUTTON_RETRYCANCEL = 5	' Retry and Cancel buttons

Const DEFBUTTON1 = 0	' First button is default
Const DEFBUTTON2 = 256	' Second button is default
Const DEFBUTTON3 = 512 	' Third button is default

' Possible return values of the MessageBox function...
'
Const IDOK = 1		' OK button clicked
Const IDCANCEL = 2	' Cancel button clicked
Const IDABORT =  3	' Abort button clicked
Const IDRETRY =  4	' Retry button clicked
Const IDIGNORE = 5	' Ignore button clicked
Const IDYES = 6		' Yes button clicked
Const IDNO = 7		' No button clicked

Sub Main
 
  Dim cap, result

  cap = "MessageBox Demo"

  result = crt.Dialog.MessageBox("OK button only, no icon, default caption")
  PrintResult(result)

  result = crt.Dialog.MessageBox("OK and Cancel buttons", cap, BUTTON_CANCEL)
  PrintResult(result)

  result = crt.Dialog.MessageBox("Abort, Retry, Ignore", cap, BUTTON_ABORTRETRYIGNORE Or ICON_QUESTION)
  PrintResult(result)

  result = crt.Dialog.MessageBox("Yes/No/Cancel buttons, No is default", cap, BUTTON_YESNOCANCEL Or DEFBUTTON2 Or ICON_INFO)
  PrintResult(result)

End Sub

' Print which result was returned.
'
Sub PrintResult(val)
  Select Case val
    Case IDOK     crt.Dialog.MessageBox("OK button pressed")
    Case IDCANCEL crt.Dialog.MessageBox("Cancel pressed")
    Case IDABORT  crt.Dialog.MessageBox("Abort pressed")
    Case IDRETRY  crt.Dialog.MessageBox("Retry pressed")
    Case IDIGNORE crt.Dialog.MessageBox("Ignore pressed")
    Case IDYES    crt.Dialog.MessageBox("Yes pressed")
    Case IDNO     crt.Dialog.MessageBox("No pressed")
    Case Else     crt.Dialog.MessageBox("Unknown result!")
  End Select
End Sub

