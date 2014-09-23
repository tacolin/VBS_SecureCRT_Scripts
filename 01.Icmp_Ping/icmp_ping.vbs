#$language = "VBScript"
#$interface = "1.0"

Function ICMP_Ping (address)
    dim wsh : set wsh = CreateObject("WScript.Shell")
    ret = wsh.Run("cmd /c ping -n 1 " & address & " | find /i " & chr(34) & "TTL" & chr(34), 1, True)
    ICMP_Ping = Not CBool(ret)
End Function

Sub Main ()
    dim wsh : set wsh = CreateObject("WScript.Shell")

    ret1 = ICMP_Ping("tw.yahoo.com")
    ret2 = ICMP_Ping("tw.xxxxx.com")

    msgbox "ping tw.yahoo.com, ret = " & ret1 & vbCr & _
           "ping tw.xxxxx.com, ret = " & ret2
End Sub
