#$language = "VBScript"
#$interface = "1.0"

' =========================================================================== '
' Select the tshark.exe file path'
' default path is = c:\Program Files\Wireshark\tshark.exe'
' =========================================================================== '
Function SelectTsharkPath (fso)
    tsharkPath = crt.Dialog.FileOpenDialog("Select tshark.exe", _
                                         "Select", _
                                         "tshark.exe", _
                                         "exe (*.exe)|*.exe||")
    If not fso.FileExists(tsharkPath) Then
        msgbox tsharkPath & vbCr & "not exist."
        SelectTsharkPath = ""
    End If
    SelectTsharkPath = chr(34) & tsharkPath & chr(34)
End Function

' =========================================================================== '
' Select Network Interface Index'
' =========================================================================== '
Function SelectInterfaceIndex (wsh, tsharkPath)
    set exec = wsh.Exec(tsharkPath & " -D")
    Do While exec.Status = 0
        crt.Sleep 100
    Loop
    stdout = exec.StdOut.ReadAll()

    index = InputBox("Please Enter the Interface Index:" & vbcr & _
                     vbcr & stdout, _
                     "Select Index", 1)
    SelectInterfaceIndex = index
End Function

' =========================================================================== '
' Check if the Input Network Interface Index is valid '
' =========================================================================== '
Function CheckInterface (wsh, tsharkPath, index)
    command = tsharkPath & " -D | find /i " & chr(34) & index & "." & chr(34)
    command = "cmd /c " & chr(34) & command & chr(34)
    ret = wsh.Run(command, 1, True)
    CheckInterface = not CBool(ret)
End Function


' =========================================================================== '
' 1. Trim the text'
' 2. replace spaces to 1 space'
' =========================================================================== '
Function RemoveSpaces (text)
    ret = Trim(text)
    For i=len(ret) To 2 Step -1
        spaces = Space(i)
        ret = Replace(ret, spaces, " ")
    Next
    RemoveSpaces = ret
End Function


' =========================================================================== '
' Main Sub'
' =========================================================================== '
Sub Main ()

    dim binarylog : binarylog = "d:\tacolin.pcap"
    dim textlog   : textlog   = "d:\tacolin.txt"
    dim tsharkPath
    dim wsh, fso, exec, stdout, file
    dim index, duration, filters, filestring, outputstring
    dim command
    dim lines, params

    set wsh = CreateObject("WScript.Shell")
    set fso = CreateObject("Scripting.FileSystemObject")


    tsharkPath = SelectTsharkPath(fso)

    index = SelectInterfaceIndex(wsh, tsharkPath)
    If not CheckInterface(wsh, tsharkPath, index) Then
        msgbox "interface index = " & index & " is invalid."
        Exit Sub
    End If

    duration = InputBox("Please Enter the Reception Duration (in seconds) :", _
                        "Enter Duration", _
                        10)

    '-------------------------------------------------------------- '
    ' Run tshark to receive packets '
    '-------------------------------------------------------------- '
    command = tsharkPath & " -i " & index & _
              " -a duration:" & duration & _
              " -w " & chr(34) & binarylog & chr(34)
    wsh.Run command, 1, True


    '-------------------------------------------------------------- '
    ' binary log file contains all packets '
    ' use filters to filter what we want and save to text log file'
    '-------------------------------------------------------------- '
    filters = InputBox("Reception over." & vbcr & _
                       "Enter the filter if you want:", _
                       "Wireshark Filter", _
                       "ip.src==192.168.1.3 && icmp")
    filters = chr(34) & filters & chr(34)

    command = tsharkPath & " -r " & chr(34) & binarylog & chr(34) & _
              " -Y " & filters & _
              " > " & chr(34) & textlog & chr(34)
    wsh.Run "cmd /c " & chr(34) & command & chr(34), 1, True

    If not fso.FileExists(textlog) Then
        msgbox textlog & " does not exist."
        Exit Sub
    End If

    set file = fso.GetFile(textlog)
    If file.Size <= 0 Then
        msgbox textlog & " file size = 0."
        Exit Sub
    End If

    '-------------------------------------------------------------- '
    ' 1. remove the redundant spaces '
    ' 2. get the field we want '
    '-------------------------------------------------------------- '
    set file = fso.OpenTextFile(textlog)
    filestring = file.ReadAll()
    lines = Split(filestring, vbcrlf)
    For i=0 To uBound(lines)
        lines(i) = RemoveSpaces(lines(i))
    Next

    outputstring = ""
    For Each line In lines
        If len(line) > 0 Then
           params = Split(line)
           outputstring = outputstring & _
                          "count = " & params(0) & _
                          " | recv time = " & FormatNumber(params(1), 3) & vbcr
        End If
    Next
    msgbox outputstring

End Sub
