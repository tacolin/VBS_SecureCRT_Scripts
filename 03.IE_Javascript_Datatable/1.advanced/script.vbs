#$language = "VBScript"
#$interface = "1.0"

'----------------------------------------------------------------------'
' Include the header file
'----------------------------------------------------------------------'
Function Include(vbsFile)
    dim fso, f, s
    set fso = CreateObject("Scripting.FileSystemObject")
    set f = fso.OpenTextFile(vbsFile)
    s = f.ReadAll()
    f.Close
    ExecuteGlobal s
End Function

Include "table.vbs"

'----------------------------------------------------------------------'
' main sub
'----------------------------------------------------------------------'
Sub Main ()

    dim tblobj : set tblobj = new table
    tblobj.openHtml("table.html")
    tblobj.setTitles( Array("c1", "c2") )
    tblobj.initialize()

    For i=0 To 20
        tblobj.addRow( Array(i, i*100) )
        crt.Sleep 500
    Next

    crt.Sleep 2000

    call tblobj.updateCell(2, 0, "2: cell modified")

    crt.Sleep 2000

    call tblobj.updateCell(18, 1, "1800: cell modified")

    crt.Sleep 2000

    call tblobj.updateRow(3, Array("3: row modified", "300 : row modified"))

    crt.Sleep 2000

    tblobj.alert("test is over")

End Sub
