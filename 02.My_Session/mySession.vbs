#$language = "VBScript"
#$interface = "1.0"

' =========================================================================== '
' MySession class '
' =========================================================================== '
Class MySession

    Private mycrt   ' crt '
    Private mysess
    Private myname
    Private myscreen ' crt.Session.Screen'

    '-------------------------------------------------------------- '
    ' Default Constructor '
    '-------------------------------------------------------------- '
    Private Sub Class_Initialize
        set mycrt = Nothing
        set mysess = Nothing
        set myscreen = Nothing
        myname = ""
    End Sub

    '-------------------------------------------------------------- '
    ' Default Destructor '
    '-------------------------------------------------------------- '
    Private Sub Class_Terminate
    End Sub

    '-------------------------------------------------------------- '
    ' initilization & connect to session '
    '-------------------------------------------------------------- '
    Public Function connect (crt, name)
        set mycrt = crt
        myname = name

        '---------------------------------------------------------- '
        ' compare name with each existed session '
        ' if the sesseion exists, get the session '
        '---------------------------------------------------------- '
        count = mycrt.getTabCount()
        For i=1 To count
            set sess = mycrt.getTab(i)
            If LCase(sess.Caption) = LCase(myname) Then
                set mysess = sess
            End If
        Next

        '---------------------------------------------------------- '
        ' if the session doesn't exist, connect it '
        ' but the session should be pre-configured in the profile '
        '---------------------------------------------------------- '
        If mysess is Nothing Then
            set mysess = mycrt.Session.ConnectInTab("/S " & myname)
        End If

        set myscreen = mysess.Screen
        cleanScreen()
    End Function

    Public Function cleanScreen ()
        mysess.Activate
        myscreen.SendSpecial("MENU_CLEAR_SCREEN_AND_SCROLLBACK")
    End Function

    '-------------------------------------------------------------- '
    ' crt.Session.Screen.Send
    '-------------------------------------------------------------- '
    Public Function send (text)
        mysess.Activate
        myscreen.Send(text)
    End Function

    '-------------------------------------------------------------- '
    ' crt.Session.Screen.Send "????" & vbcr'
    '-------------------------------------------------------------- '
    Public Function sendline (text)
        send(text & vbcr)
    End Function

    '-------------------------------------------------------------- '
    ' CTRL + C '
    '-------------------------------------------------------------- '
    Public Function ctrlC ()
        send( chr(3) )
    End Function

    '-------------------------------------------------------------- '
    ' CTRL + D '
    '-------------------------------------------------------------- '
    Public Function ctrlD ()
        send( chr(4) )
    End Function

    '-------------------------------------------------------------- '
    ' crt.Session.Screen.WaitforString '
    '-------------------------------------------------------------- '
    Public Function wait (text)
        mysess.Activate
        myscreen.WaitForString(text)
    End Function

    '-------------------------------------------------------------- '
    ' crt.Session.Screen.WaitforStrings '
    '-------------------------------------------------------------- '
    Public Function waitArray (textArray)
        mysess.Activate
        waitArray = myscreen.WaitForStrings(textArray)
    End Function

    '-------------------------------------------------------------- '
    ' crt.Session.Screen.ReadString '
    '-------------------------------------------------------------- '
    Public Function read (prompt)
        mysess.Activate
        read = myscreen.ReadString(prompt)
    End Function

    '-------------------------------------------------------------- '
    ' crt.Sleep : mini-second version'
    '-------------------------------------------------------------- '
    Public Function msleep (miniseconds)
        mysess.Activate
        mycrt.sleep(miniseconds)
    End Function

    '-------------------------------------------------------------- '
    ' crt.Sleep : second version'
    '-------------------------------------------------------------- '
    Public Function sleep (seconds)
        mysess.Activate
        mycrt.sleep(seconds*1000)
    End Function

End Class


' =========================================================================== '
' Main Sub '
' =========================================================================== '
Sub Main ()
    ' dim ptt : set ptt = new MySession
    ' ptt.connect crt, "ptt"
    ' ptt.sendline("guest")
End Sub
