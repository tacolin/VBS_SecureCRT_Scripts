Class table

    Public ieobj
    Public win

    Private Sub Class_Initialize
        ' Called automatically when class is created
    End Sub

    Private Sub Class_Terminate
        ' Called automatically when all references to class instance are removed
    End Sub

    Public Function openHtml (relative_htmlpath)
        ' get current directory'
        set fso = CreateObject("Scripting.FileSystemObject")
        set file = fso.GetFile(crt.ScriptFullName)
        currentFolder = fso.GetParentFolderName(file)

        ' get absolute html file path '
        absolute_htmlpath = currentFolder & "/" & relative_htmlpath
        ' check if file exists '
        If not fso.FileExists(absolute_htmlpath) Then
            msgbox absolute_htmlpath & " does not exist."
            Exit Function
        End If

        set ieobj = CreateObject("InternetExplorer.Application")
        ieobj.Navigate(absolute_htmlpath)

        ' wait for ie ready'
        Do
            crt.Sleep 100
        Loop While ieobj.Busy

        ieobj.MenuBar = false
        ieobj.StatusBar = false
        ieobj.AddressBar = false
        ieobj.Toolbar = false

        set win = ieobj.document.parentWindow

        ieobj.Visible = True
        win.focus()
    End Function

    Public Function setTitles (titles)
        win.setTitles( titles )
    End Function

    Public Function initialize ()
        win.initialize()
        ' ieobj.Visible = True
        ' win.focus()
    End Function

    Public Function addRow (rowDataArray)
        rowIndex = win.addRow(rowDataArray)
        addRow = rowIndex
    End Function

    Public Function updateCell (rowIndex, columnIndex, cellData)
        call win.updateCell(rowIndex, columnIndex, cellData)
    End Function

    Public Function updateRow (rowIndex, rowDataArray)
        call win.updateRow(rowIndex, rowDataArray)
    End Function

    Public Function alert (text)
        win.alert(text)
    End Function

End Class
