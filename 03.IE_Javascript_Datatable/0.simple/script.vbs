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

Sub Main ()

    msgbox "please select your html file ..."

    filepath = crt.Dialog.FileOpenDialog("Select a HTML file", _
                                         "Select", _
                                         "index.html", _
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

End Sub
