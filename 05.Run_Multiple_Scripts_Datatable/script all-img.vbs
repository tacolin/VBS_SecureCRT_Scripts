#$language = "VBScript"
#$interface = "1.0"

'----------------------------------------------------------------------'
' re-define the functions which exist in scirpt1 and script2
'----------------------------------------------------------------------'
Function warning (text)
End Function

Function error (text)
End Function

Function info (text)
End Function

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

Include "runner.vbs"
Include "table.vbs"

'----------------------------------------------------------------------'
' main sub
'----------------------------------------------------------------------'
Sub Main ()

    dim run : set run = new runner
    dim tbl : set tbl = new table

    tbl.openHtml("table.html")
    tbl.setTitles( Array("script name", "return value") )
    tbl.initialize()

    dim scripts : scripts = Array("script1.vbs", "script2.vbs", "script3.vbs")

    okImg = "<img src='table_media/images/icon_ok.png' height=28/>"
    failImg = "<img src='table_media/images/icon_fail.png' height=28/>"
    naImg = "<img src='table_media/images/icon_na.png' height=28/>"

    For Each s In scripts
        tbl.addRow( Array(s, naImg) )
        crt.Sleep 1000
    Next

    For i=0 To Ubound(scripts)
        result = run.RunVbsFile( scripts(i) )
        If result Then
            call tbl.updateCell(i, 1, okImg)
        Else
            call tbl.updateCell(i, 1, failImg)
        End If
        crt.Sleep 1000
    Next

    tbl.alert("all scripts is over.")

End Sub
