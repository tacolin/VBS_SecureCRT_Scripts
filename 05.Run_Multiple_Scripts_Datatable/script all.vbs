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

    For Each s In scripts
        tbl.addRow( Array(s, "N/A") )
        crt.Sleep 1000
    Next

    For i=0 To Ubound(scripts)
        result = run.RunVbsFile( scripts(i) )
        call tbl.updateRow( i, Array(scripts(i), result) )
        crt.Sleep 1000
    Next

    tbl.alert("all scripts is over.")

End Sub
