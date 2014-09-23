#$language = "VBScript"
#$interface = "1.0"

' Please following the steps of this page to enable active x on local file. '
' '
' http://goo.gl/xOruWy '
' '
' 1. Click the "Tools" menu in Internet Explorer
' 2. Select "Internet Options"
' 3. Click the "Advanced" tab
' 4. Scroll down to the "Security" section and place a check mark next to
'    "Allow active content to run in files on My Computer" '

isVbsEnd = true

Sub Main ()

    msgbox "please select your html file ..."

    filepath = crt.Dialog.FileOpenDialog("Select a HTML file", _
                                         "Select", _
                                         "html_js_side.html", _
                                         "Web Page (*.html, *.htm)|*.html;*.htm||")

    dim fso : set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filepath) Then
        msgbox filepath & vbCr & "not exist"
        Exit Sub
    End If

    dim ie : set ie = CreateObject("InternetExplorer.Application")
    ie.navigate filepath

    Do
        crt.sleep 100
    Loop While ie.Busy

    ie.visible = True
    ie.document.parentWindow.focus

    Do
        Select Case ie.document.parentWindow.vbsEvent

        Case 1
            data = Split(ie.document.parentWindow.params, ",")
            For Each d In data
                msgbox d
            Next

        Case 2
            text = ie.document.parentWindow.params
            msgbox "text = " & text

        Case Else
            'do nothing'
        End Select

        ie.document.parentWindow.vbsEvent = 0
        crt.sleep 100

    Loop While ie.document.parentWindow.ending <> -1

    ie.Quit

End Sub
